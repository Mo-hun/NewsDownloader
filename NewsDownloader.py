import urllib.request
import urllib.parse
import urllib.error
import json
import pandas as pd
import re
import sys
import os
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from PressInfo import extract_press_info

DISPLAY = 100
SORT = "date"
NAVER_DATE_FORMAT = "%a, %d %b %Y %H:%M:%S %z"


def get_resource_path(*paths):
    if getattr(sys, "_MEIPASS", None):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, *paths)


def remove_html_tags(text):
    if not text:
        return ""
    return re.sub(r"<.*?>", "", text)


def parse_pub_date(pub_date_str):
    return datetime.strptime(pub_date_str, NAVER_DATE_FORMAT)


def fetch_news_page(client_id, client_secret, query, start):
    enc_text = urllib.parse.quote(query)
    url = (
        f"https://openapi.naver.com/v1/search/news.json"
        f"?query={enc_text}&display={DISPLAY}&start={start}&sort={SORT}"
    )

    request = urllib.request.Request(url)
    request.add_header("X-Naver-Client-Id", client_id)
    request.add_header("X-Naver-Client-Secret", client_secret)

    with urllib.request.urlopen(request) as response:
        body = response.read().decode("utf-8")
        return json.loads(body)


def collect_news(client_id, client_secret, query, mode="24h", target_date=None):
    rows = []

    now = datetime.now(timezone(timedelta(hours=9)))

    if mode == "24h":
        cutoff_start = now - timedelta(hours=24)
        cutoff_end = now

    elif mode == "date":
        start_dt = datetime.strptime(target_date, "%Y-%m-%d")
        cutoff_start = start_dt.replace(tzinfo=timezone(timedelta(hours=9)))
        cutoff_end = cutoff_start + timedelta(days=1)

    start = 1

    while start <= 1000:
        print(f"[{query}] 조회중 start={start}")

        try:
            data = fetch_news_page(client_id, client_secret, query, start)
        except urllib.error.HTTPError as e:
            print("HTTP ERROR:", e.code)
            print(e.read().decode("utf-8", errors="ignore"))
            break
        except Exception as e:
            print("ERROR:", str(e))
            break

        items = data.get("items", [])
        if not items:
            break

        page_24_count = 0

        for item in items:
            pub_raw = item.get("pubDate")
            if not pub_raw:
                continue

            try:
                pub_dt = parse_pub_date(pub_raw)
            except Exception:
                continue

            if cutoff_start <= pub_dt < cutoff_end:
                page_24_count += 1

                originallink = item.get("originallink", "")
                naverlink = item.get("link", "")

                press_name, press_category, press_domain = extract_press_info(originallink)

                rows.append({
                    "검색어": query,
                    "언론사카테고리": press_category,
                    "언론사명": press_name,
                    "언론사도메인": press_domain,
                    "제목": remove_html_tags(item.get("title")),
                    "요약": remove_html_tags(item.get("description")),
                    "언론사링크": originallink,
                    "네이버링크": naverlink,
                    "작성일": pub_dt.strftime("%Y-%m-%d %H:%M:%S"),
                })

        print(f"[{query}] 24시간 기사: {page_24_count}/{len(items)}")

        if len(items) == DISPLAY and page_24_count == DISPLAY:
            start += DISPLAY
        else:
            break

    return rows


def build_queries(required_keywords, optional_keywords):
    required_keywords = [k.strip() for k in required_keywords if k and k.strip()]
    optional_keywords = [k.strip() for k in optional_keywords if k and k.strip()]

    if not required_keywords:
        raise ValueError("필수 키워드는 1개 이상 필요합니다.")

    if not optional_keywords:
        return list(dict.fromkeys(required_keywords))

    queries = []
    for required_keyword in required_keywords:
        for optional_keyword in optional_keywords:
            queries.append(f"{required_keyword} {optional_keyword}".strip())

    return list(dict.fromkeys(queries))

def autosize_excel(output):
    wb = load_workbook(output)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)

        ws.column_dimensions[col_letter].width = min(max_length + 2, 60)

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 24
    ws.column_dimensions["E"].width = 55
    ws.column_dimensions["F"].width = 90
    ws.column_dimensions["G"].width = 45
    ws.column_dimensions["H"].width = 45
    ws.column_dimensions["I"].width = 20

    wb.save(output)


def load_template(template_path):
    with open(template_path, "r", encoding="utf-8") as f:
        return f.read()


def generate_html_review(rows, html_output_name, csv_output_name, json_output_name):
    template_path = get_resource_path("templates", "NewsReviewTemplate.html")
    template = load_template(template_path)

    rows_json = json.dumps(rows, ensure_ascii=False)

    html = (
        template
        .replace("__NEWS_ROWS_JSON__", rows_json)
        .replace("__CSV_OUTPUT_NAME__", csv_output_name)
        .replace("__JSON_OUTPUT_NAME__", json_output_name)
        .replace("__NEWS_COUNT__", str(len(rows)))
    )

    with open(html_output_name, "w", encoding="utf-8") as f:
        f.write(html)


def main():
    if len(sys.argv) < 5:
        print('사용법: NewsDownloader.exe "CLIENT_ID" "CLIENT_SECRET" "필수키워드1,필수키워드2" "선택키워드1,선택키워드2"')
        print('예시: NewsDownloader.exe "id" "secret" "한전,한국전력공사" "민원,건설,철탑,송전선로"')
        sys.exit(1)

    client_id = sys.argv[1].strip()
    client_secret = sys.argv[2].strip()
    required_arg = sys.argv[3].strip()
    optional_arg = sys.argv[4].strip()

    required_keywords = [q.strip() for q in required_arg.split(",") if q.strip()]
    optional_keywords = [q.strip() for q in optional_arg.split(",") if q.strip()]

    try:
        queries = build_queries(required_keywords, optional_keywords)
    except ValueError as e:
        print(str(e))
        sys.exit(1)

    print("생성된 검색어:")
    for query in queries:
        print("-", query)

    all_rows = []

    for query in queries:
        rows = collect_news(client_id, client_secret, query)
        all_rows.extend(rows)

    if not all_rows:
        print("수집된 뉴스가 없습니다.")
        sys.exit(0)

    df = pd.DataFrame(all_rows)

    column_order = [
        "검색어",
        "언론사카테고리",
        "언론사명",
        "언론사도메인",
        "제목",
        "요약",
        "언론사링크",
        "네이버링크",
        "작성일",
    ]

    df = df[column_order]

    if "네이버링크" in df.columns:
        df = df.drop_duplicates(subset=["네이버링크"])

    if "작성일" in df.columns:
        df = df.sort_values(by="작성일", ascending=False)

    rows_for_output = df.to_dict(orient="records")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_output = f"건설이슈언론모니터링_{timestamp}.xlsx"
    html_output = f"건설이슈언론모니터링_{timestamp}.html"
    csv_output = f"건설이슈언론모니터링_선별결과_{timestamp}.csv"
    json_output = f"{timestamp}.json"

    df.to_excel(excel_output, index=False)
    autosize_excel(excel_output)
    generate_html_review(rows_for_output, html_output, csv_output, json_output)

    print("엑셀 저장 완료:", os.path.abspath(excel_output))
    print("HTML 저장 완료:", os.path.abspath(html_output))
    print("HTML에서 CSV 저장 시 파일명:", csv_output)
    print("HTML에서 JSON 저장 시 파일명:", json_output)


if __name__ == "__main__":
    main()