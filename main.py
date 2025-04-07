import streamlit as st
import pandas as pd
import docx
import difflib
import logging
import os
import time
from PIL import Image
import imagehash
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
import io
from urllib.parse import urljoin
import re

# 基本設定與常數
CONFIG = {
    "MAX_RETRIES": 3,
    "SIMILARITY_THRESHOLD": 0.99,
    "REQUEST_TIMEOUT": 15,
    "SCREENSHOT_DIR": "screenshots",
    "FOOTER_PHRASES": [
        "隱私權政策", "著作權聲明", "服務條款", "回到頁面頂層",
        "高雄市政府", "本網站最佳瀏覽模式", "Copyright", "reCAPTCHA", "適用"
    ],
    "NEW_SITE_SELECTORS": {
        "article_list": "div.card.is-grid",
        "article_item": "a.card-item.card-link",
        "title_link": "h3.card-title",
        "content_area": "div.article-body"
    }
}

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
os.makedirs(CONFIG["SCREENSHOT_DIR"], exist_ok=True)

# Streamlit UI 初始化
def setup_ui():
    st.set_page_config(page_title="高雄畫刊智能比對工具", layout="wide")
    st.title("📘 高雄畫刊智能比對工具")

# 從上傳的 Word 檔中擷取「期數、新舊網址」組合
def extract_triples_from_docx(file):
    doc = docx.Document(file)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    triples = []
    i = 0
    url_pattern = re.compile(r'^https?://')
    while i < len(paras):
        if i + 2 >= len(paras):
            break
        title = paras[i]
        new_url = paras[i+1]
        middle = paras[i+2]
        if i + 3 < len(paras):
            old_url = paras[i+3]
            step = 4
            if not old_url.startswith("http") and i + 4 < len(paras) and paras[i+4].startswith("http"):
                old_url += paras[i+4]
                step = 5
        else:
            break
        if url_pattern.match(new_url) and "原期刊" in middle and url_pattern.match(old_url):
            triples.append((title, new_url, old_url))
            i += step
        else:
            i += 1
    return triples

# 清理段落（去除空白與頁尾標語）
def clean_paragraphs(paragraphs):
    return [p.strip() for p in paragraphs if p.strip() and not any(x in p for x in CONFIG["FOOTER_PHRASES"])]

# 擷取主要內容（例如遇到特定標識後開始擷取）
def extract_main_content(paragraphs):
    for i, p in enumerate(paragraphs):
        if "臉書粉絲團" in p:
            return paragraphs[i+1:]
    return paragraphs

# 截圖功能（方便除錯）
def take_screenshot(driver, name):
    path = os.path.join(CONFIG["SCREENSHOT_DIR"], name)
    driver.save_screenshot(path)
    try:
        img = Image.open(path)
        st.image(img, caption=name, use_container_width=True)
    except Exception as e:
        st.error(f"❌ 無法顯示截圖: {e}")

# 初始化 Chrome WebDriver
def init_driver():
    try:
        chromedriver_autoinstaller.install()
        opts = Options()
        opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options=opts)
        logger.info("WebDriver 初始化成功")
        return driver
    except Exception as e:
        st.error(f"❌ WebDriver 初始化失敗: {str(e)}")
        raise

# 根據是否新網站，分別擷取文章內容
def fetch_articles_by_driver(driver, url, is_new_site=False):
    if is_new_site:
        return fetch_new_site_articles(driver, url)
    else:
        return fetch_old_site_articles(driver, url)

# 舊網站文章擷取
def fetch_old_site_articles(driver, url):
    articles = {}
    try:
        driver.get(url)
        WebDriverWait(driver, 15).until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(1)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        list_div = soup.find("div", class_="sub_peaper_list_page")
        if not list_div:
            st.error("❌ 找不到舊網站的文章列表容器")
            return articles
        li_elements = list_div.find_all("li")
        titles, hrefs = [], []
        for li in li_elements:
            a = li.find("a")
            if a and "跳到主要內容區塊" not in a.get_text(strip=True):
                href = urljoin(url, a.get("href"))
                titles.append(a.get_text(strip=True))
                hrefs.append(href)
        
        visited_pages = set()
        for idx, href in enumerate(hrefs):
            all_paragraphs = []
            all_images = []
            for attempt in range(CONFIG["MAX_RETRIES"]):
                try:
                    driver.get(href)
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)
                    html = driver.page_source
                    soup_article = BeautifulSoup(html, "html.parser")
                    
                    page_title_div = soup_article.find("div", class_="page_title")
                    article_title = (page_title_div.get_text(strip=True)
                                     if page_title_div and "跳到主要內容區塊" not in page_title_div.get_text(strip=True)
                                     else titles[idx] if idx < len(titles) else f"文章{idx+1}")
                    
                    content_div = soup_article.find("div", class_="page") or soup_article.find("body")
                    paragraphs = [p.get_text(strip=True) for p in content_div.find_all("p") if p.get_text(strip=True)]
                    exclude_headers = ["活動名稱", "活動時間", "活動地點", "主辦單位", "活動說明"]
                    for td in content_div.find_all("td"):
                        text = td.get_text(strip=True)
                        if text and text not in exclude_headers and not any(header in text for header in exclude_headers):
                            paragraphs.append(text)
                    
                    cleaned = clean_paragraphs(paragraphs)
                    main_content = extract_main_content(cleaned)
                    all_paragraphs.extend(main_content)
                    
                    imgs = content_div.find_all("img")
                    for img in imgs:
                        src = img.get("src")
                        if src:
                            all_images.append(urljoin(driver.current_url, src))
                    
                    # 分頁處理
                    pagination = soup_article.find("div", class_="list_gotopage")
                    if pagination:
                        page_links = pagination.find_all("a")
                        for a in page_links:
                            page_href = urljoin(driver.current_url, a.get("href"))
                            if page_href not in visited_pages:
                                visited_pages.add(page_href)
                                driver.get(page_href)
                                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                                time.sleep(2)
                                html_page = driver.page_source
                                soup_page = BeautifulSoup(html_page, "html.parser")
                                content_div_page = soup_page.find("div", class_="page") or soup_page.find("body")
                                
                                paragraphs_page = [p.get_text(strip=True) for p in content_div_page.find_all("p") if p.get_text(strip=True)]
                                for td in content_div_page.find_all("td"):
                                    text = td.get_text(strip=True)
                                    if text and text not in exclude_headers and not any(header in text for header in exclude_headers):
                                        paragraphs_page.append(text)
                                
                                cleaned_page = clean_paragraphs(paragraphs_page)
                                main_content_page = extract_main_content(cleaned_page)
                                all_paragraphs.extend(main_content_page)
                                
                                imgs_page = content_div_page.find_all("img")
                                for img in imgs_page:
                                    src = img.get("src")
                                    if src:
                                        all_images.append(urljoin(driver.current_url, src))
                    
                    unique_paragraphs = list(dict.fromkeys(all_paragraphs))
                    unique_images = list(dict.fromkeys(all_images))
                    articles[article_title] = {"paragraphs": unique_paragraphs, "images": unique_images}
                    st.success(f"✅ 舊文章擷取成功：{article_title}")
                    break
                except Exception as e:
                    st.warning(f"⚠️ 第 {attempt+1} 次擷取失敗：{str(e)[:100]}")
                    if attempt == CONFIG["MAX_RETRIES"] - 1:
                        take_screenshot(driver, f"fail_{idx+1}.png")
    except Exception as e:
        st.error(f"❌ 舊網站解析錯誤: {str(e)[:200]}")
        take_screenshot(driver, "old_site_error.png")
    return articles

# 新網站文章擷取
def fetch_new_site_articles(driver, url):
    articles = {}
    try:
        driver.get(url)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, CONFIG["NEW_SITE_SELECTORS"]["article_list"]))
        )
        time.sleep(2)
        articles_blocks = driver.find_elements(By.CSS_SELECTOR,
            f"{CONFIG['NEW_SITE_SELECTORS']['article_list']} {CONFIG['NEW_SITE_SELECTORS']['article_item']}"
        )
        if not articles_blocks:
            st.error("❌ 未找到新網站文章區塊")
            return articles
        st.info(f"🔍 發現 {len(articles_blocks)} 篇新網站文章")
        main_window = driver.current_window_handle
        for idx, block in enumerate(articles_blocks, 1):
            try:
                title_element = block.find_element(By.CSS_SELECTOR, CONFIG["NEW_SITE_SELECTORS"]["title_link"])
                title = title_element.text.strip()
                article_url = block.get_attribute("href")
                if not title or not article_url:
                    continue
                driver.execute_script("window.open('');")
                WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
                new_window = [wh for wh in driver.window_handles if wh != main_window][-1]
                driver.switch_to.window(new_window)
                driver.get(article_url)
                try:
                    WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, CONFIG["NEW_SITE_SELECTORS"]["content_area"]))
                    )
                    content_div = driver.find_element(By.CSS_SELECTOR, CONFIG["NEW_SITE_SELECTORS"]["content_area"])
                    paragraphs = [p.text.strip() for p in content_div.find_elements(By.TAG_NAME, "p")]
                    cleaned = clean_paragraphs(paragraphs)
                    main_content = extract_main_content(cleaned)
                    
                    imgs = content_div.find_elements(By.TAG_NAME, "img")
                    images = [img.get_attribute("src") for img in imgs if img.get_attribute("src")]
                    unique_images = list(dict.fromkeys(images))
                    
                    if main_content:
                        articles[title] = {"paragraphs": main_content, "images": unique_images}
                        st.success(f"✅ 新文章擷取成功：{title}")
                    else:
                        st.warning(f"⚠️ 文章內容為空：{title}")
                        take_screenshot(driver, f"new_empty_{idx}.png")
                except Exception as e:
                    st.warning(f"⚠️ 擷取文章內容失敗：{str(e)[:100]}")
                    take_screenshot(driver, f"new_content_fail_{idx}.png")
                finally:
                    driver.close()
                    driver.switch_to.window(main_window)
                    time.sleep(1)
            except Exception as e:
                st.warning(f"⚠️ 處理文章區塊失敗：{str(e)[:100]}")
                continue
    except Exception as e:
        st.error(f"❌ 新網站解析錯誤: {str(e)[:200]}")
        take_screenshot(driver, "new_site_error.png")
    return articles

# 文字比對：採用 difflib 判斷相似度，不足則記錄「未找到匹配」
def compare_paragraphs(p1, p2, similarity_threshold=0.8):
    p1_cleaned = clean_paragraphs(p1)
    p2_cleaned = clean_paragraphs(p2)
    missing_segments = []
    image_caption_keywords = ["圖", "圖片", "說明", "來源", "註解"]
    p1_filtered = [para for para in p1_cleaned if not any(keyword in para for keyword in image_caption_keywords) or len(para) > 50]
    for old_para in p2_cleaned:
        best_match = max(p1_filtered, key=lambda x: difflib.SequenceMatcher(None, old_para, x).ratio(), default="")
        similarity = difflib.SequenceMatcher(None, old_para, best_match).ratio()
        if similarity < similarity_threshold:
            missing_segments.append(f"舊段落未找到匹配：{old_para[:50]}...")
    if not missing_segments:
        return "一致", ""
    else:
        diff_text = "\n".join(missing_segments)
        return "不一致", diff_text

# 圖片比對：下載圖片後利用多重哈希計算相似度
def compare_images(new_imgs, old_imgs, timeout=10):
    new_image_data = []
    old_image_data = []
    failed_downloads = []
    target_size = (256, 256)
    for url in new_imgs:
        try:
            response = requests.get(url, stream=True, timeout=timeout)
            response.raise_for_status()
            img = Image.open(response.raw)
            img_resized = img.resize(target_size, Image.Resampling.LANCZOS)
            avg_hash = imagehash.average_hash(img_resized)
            phash = imagehash.phash(img_resized)
            dhash = imagehash.dhash(img_resized)
            new_image_data.append((str(avg_hash), str(phash), str(dhash), url))
        except Exception as e:
            failed_downloads.append(f"無法下載新圖片：{url} - {str(e)[:50]}")
    for url in old_imgs:
        try:
            response = requests.get(url, stream=True, timeout=timeout)
            response.raise_for_status()
            img = Image.open(response.raw)
            img_resized = img.resize(target_size, Image.Resampling.LANCZOS)
            avg_hash = imagehash.average_hash(img_resized)
            phash = imagehash.phash(img_resized)
            dhash = imagehash.dhash(img_resized)
            old_image_data.append((str(avg_hash), str(phash), str(dhash), url))
        except Exception as e:
            failed_downloads.append(f"無法下載舊圖片：{url} - {str(e)[:50]}")
    missing_images = []
    similarity_scores = []
    for old_avg, old_phash, old_dhash, old_url in old_image_data:
        found = False
        best_similarity = 0.0
        for new_avg, new_phash, new_dhash, new_url in new_image_data:
            avg_similarity = 1.0 - (imagehash.hex_to_hash(old_avg) - imagehash.hex_to_hash(new_avg)) / 64.0
            phash_similarity = 1.0 - (imagehash.hex_to_hash(old_phash) - imagehash.hex_to_hash(new_phash)) / 64.0
            dhash_similarity = 1.0 - (imagehash.hex_to_hash(old_dhash) - imagehash.hex_to_hash(new_dhash)) / 64.0
            max_similarity = max(avg_similarity, phash_similarity, dhash_similarity)
            best_similarity = max(best_similarity, max_similarity)
            if max_similarity > 0.70:
                found = True
                break
        similarity_scores.append((old_url, best_similarity))
        if not found:
            missing_images.append(old_url)
    if not missing_images:
        result = "一致"
        diff_text = ""
    else:
        result = "不一致"
        diff_text = f"遺漏圖片數：{len(missing_images)}\n新圖片數：{len(new_imgs)}, 舊圖片數：{len(old_imgs)}"
        diff_text += "\n\n遺漏的圖片URL：\n" + "\n".join(missing_images[:5])
        if len(missing_images) > 5:
            diff_text += f"\n...以及其他 {len(missing_images) - 5} 張圖片"
    if failed_downloads:
        diff_text += "\n下載失敗的圖片：\n" + "\n".join(failed_downloads[:5])
        if len(failed_downloads) > 5:
            diff_text += f"\n...以及其他 {len(failed_downloads) - 5} 張圖片"
    diff_text += "\n\n圖片相似度分數（舊圖片與最佳匹配）：\n"
    for url, score in similarity_scores:
        diff_text += f"{url}: {score:.2%}\n"
    return result, diff_text

# 每一期的處理：新舊網站比對並回傳該期結果
def process_period(period_title, new_url, old_url):
    period_results = []
    st.subheader(f"📑 處理期數：{period_title}")
    st.markdown(f"新網址： {new_url}")
    st.markdown(f"舊網址： {old_url}")
    
    driver = init_driver()
    try:
        st.info("📄 正在分析舊網站連結...")
        old_articles = fetch_articles_by_driver(driver, old_url, is_new_site=False)
        st.info(f"舊網站文章數：{len(old_articles)}")
        st.info("📄 正在分析新網站連結...")
        new_articles = fetch_articles_by_driver(driver, new_url, is_new_site=True)
        st.info(f"新網站文章數：{len(new_articles)}")
        
        # 以進度條顯示處理進度
        progress_bar = st.progress(0)
        total = len(new_articles)
        for idx, (new_title, new_data) in enumerate(new_articles.items(), start=1):
            best_match = max(old_articles.keys(), key=lambda k: difflib.SequenceMatcher(None, new_title, k).ratio(), default=None)
            if best_match:
                old_data = old_articles[best_match]
                text_result, text_diff = compare_paragraphs(new_data["paragraphs"], old_data["paragraphs"])
                image_result, image_diff = compare_images(new_data["images"], old_data["images"])
                title_diff = difflib.unified_diff([new_title], [best_match])
                title_diff_text = "".join(title_diff) if new_title != best_match else ""
                
                # 若比對一致，對應欄位保留空白；若文字不一致且出現「未找到匹配」，則只呈現舊網站內容
                if text_result == "一致":
                    old_text_diff = ""
                    new_text_diff = ""
                else:
                    if "未找到匹配" in text_diff:
                        old_text_diff = text_diff
                        new_text_diff = ""
                    else:
                        old_text_diff = text_diff
                        new_text_diff = text_diff
                
                if image_result == "一致":
                    image_diff = ""
                
                period_results.append({
                    "期數": period_title,
                    "文章標題": new_title,
                    "標題差異": title_diff_text,
                    "文字比對結果": text_result,
                    "舊文字內容差異": old_text_diff,
                    "新文字內容差異": new_text_diff,
                    "新文字段數": len(new_data["paragraphs"]),
                    "舊文字段數": len(old_data["paragraphs"]),
                    "圖片比對結果": image_result,
                    "圖片內容差異": image_diff,
                    "新圖片數": len(new_data["images"]),
                    "舊圖片數": len(old_data["images"]),
                })
            progress_bar.progress(idx / total)
        return period_results
    except Exception as e:
        st.error(f"❌ 本期處理錯誤: {str(e)[:200]}")
        return period_results
    finally:
        driver.quit()
        logger.info("WebDriver 已關閉")

# 輸出本期結果 Excel 檔（使用 XlsxWriter 格式化）
def output_excel(results, period_title):
    if results:
        df = pd.DataFrame(results)
        df.sort_values(by=["期數", "文章標題"], inplace=True)
        st.subheader(f"📊 {period_title} 校對結果")
        st.dataframe(df)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="比對結果")
            workbook  = writer.book
            worksheet = writer.sheets["比對結果"]
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)
        buffer.seek(0)
        st.download_button("💾 下載本期結果 Excel", data=buffer, file_name=f"{period_title}_比對結果.xlsx")

# 主程式：上傳 Word 後依序處理每一期，並產生該期 Excel 檔，同時累計顯示所有期數結果總覽
def main():
    setup_ui()
    uploaded = st.file_uploader("📄 上傳包含新舊網址的 Word 檔", type=["docx"])
    if not uploaded:
        return
    entries = extract_triples_from_docx(uploaded)
    st.info(f"📋 偵測到 {len(entries)} 組網址")
    
    all_results = []
    for period_title, new_url, old_url in entries:
        period_results = process_period(period_title, new_url, old_url)
        # 每完成一期就輸出該期 Excel
        output_excel(period_results, period_title)
        all_results.extend(period_results)
    
    if all_results:
        df_total = pd.DataFrame(all_results)
        df_total.sort_values(by=["期數", "文章標題"], inplace=True)
        st.subheader("📊 全部期數校對結果總覽")
        st.dataframe(df_total)
    else:
        st.error("❌ 沒有任何文章成功比對，請確認網站內容是否正確或可讀取")

if __name__ == "__main__":
    main()

