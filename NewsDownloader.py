import urllib.request
import urllib.parse
import urllib.error
import json
import pandas as pd
import re
import sys
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook

DISPLAY = 100
SORT = "date"
NAVER_DATE_FORMAT = "%a, %d %b %Y %H:%M:%S %z"


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


def collect_news(client_id, client_secret, query):
    rows = []

    now = datetime.now(timezone(timedelta(hours=9)))
    cutoff = now - timedelta(hours=24)
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

            if pub_dt >= cutoff:
                page_24_count += 1
                rows.append({
                    "검색어": query,
                    "제목": remove_html_tags(item.get("title")),
                    "요약": remove_html_tags(item.get("description")),
                    "언론사링크": item.get("originallink"),
                    "네이버링크": item.get("link"),
                    "작성일": pub_dt.strftime("%Y-%m-%d %H:%M:%S")
                })

        print(f"[{query}] 24시간 기사: {page_24_count}/{len(items)}")

        if len(items) == DISPLAY and page_24_count == DISPLAY:
            start += DISPLAY
        else:
            break

    return rows


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

    # 자주 긴 컬럼은 고정폭으로 더 보기 좋게
    if "B" in ws.column_dimensions:
        ws.column_dimensions["B"].width = 50  # 제목
    if "C" in ws.column_dimensions:
        ws.column_dimensions["C"].width = 80  # 요약
    if "D" in ws.column_dimensions:
        ws.column_dimensions["D"].width = 40
    if "E" in ws.column_dimensions:
        ws.column_dimensions["E"].width = 40

    wb.save(output)


def main():
    if len(sys.argv) < 4:
        print('사용법: 건설이슈모니터링.exe "CLIENT_ID" "CLIENT_SECRET" "검색어1,검색어2"')
        sys.exit(1)

    client_id = sys.argv[1].strip()
    client_secret = sys.argv[2].strip()
    queries_arg = sys.argv[3].strip()

    queries = [q.strip() for q in queries_arg.split(",") if q.strip()]
    if not queries:
        print("검색어가 비어 있습니다.")
        sys.exit(1)

    all_rows = []

    for query in queries:
        rows = collect_news(client_id, client_secret, query)
        all_rows.extend(rows)

    if not all_rows:
        print("수집된 뉴스가 없습니다.")
        sys.exit(0)

    df = pd.DataFrame(all_rows)

    if "네이버링크" in df.columns:
        df = df.drop_duplicates(subset=["네이버링크"])

    if "작성일" in df.columns:
        df = df.sort_values(by="작성일", ascending=False)

    today = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = f"건설이슈모니터링_{today}.xlsx"

    df.to_excel(output, index=False)
    autosize_excel(output)

    print("저장 완료:", output)


if __name__ == "__main__":
    main()