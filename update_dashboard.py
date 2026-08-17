import html
import os
import re
from datetime import datetime
from urllib.parse import unquote

import folium
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


API_URL = "https://apis.data.go.kr/1613000/HWSPR02/rsdtRcritNtcList"
REGIONS = {"11": "서울", "41": "경기"}
REGION_COORDS = {"서울": (37.5665, 126.9780), "경기": (37.4138, 127.5183)}


def create_http_session():
    """공공데이터 서버의 일시적인 지연·오류를 자동 재시도합니다."""
    retry = Retry(
        total=5,
        connect=5,
        read=5,
        status=5,
        backoff_factor=3,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({"User-Agent": "housing-dashboard-github-actions/1.0"})
    return session


HTTP = create_http_session()


def find_item_list(value):
    """API 응답 안에서 공고 목록(list[dict])을 찾습니다."""
    if isinstance(value, list) and (not value or isinstance(value[0], dict)):
        return value
    if isinstance(value, dict):
        for key in ("item", "items"):
            if key in value:
                found = find_item_list(value[key])
                if found is not None:
                    return found
        for child in value.values():
            found = find_item_list(child)
            if found is not None:
                return found
    return None


def pick(item, exact_keys, fuzzy_groups=()):
    """응답 필드명이 조금 바뀌어도 원하는 값을 최대한 찾아냅니다."""
    for key in exact_keys:
        value = item.get(key)
        if value not in (None, ""):
            return str(value).strip()

    lowered = {str(key).lower(): value for key, value in item.items()}
    for group in fuzzy_groups:
        for key, value in lowered.items():
            if value not in (None, "") and all(token in key for token in group):
                return str(value).strip()
    return ""


def parse_date(value):
    if not value:
        return pd.NaT
    digits = re.sub(r"[^0-9]", "", str(value))
    if len(digits) >= 8:
        return pd.to_datetime(digits[:8], format="%Y%m%d", errors="coerce")
    return pd.to_datetime(value, errors="coerce")


def fetch_region(api_key, region_code, region_name):
    params = {
        "serviceKey": unquote(api_key),
        "brtcCode": region_code,
        "pageNo": 1,
        "numOfRows": 1000,
    }
    # 연결 30초, 응답 90초. 실패하면 위 설정에 따라 최대 5회 재시도합니다.
    response = HTTP.get(API_URL, params=params, timeout=(30, 90))
    response.raise_for_status()

    try:
        payload = response.json()
    except ValueError as exc:
        preview = response.text[:300].replace("\n", " ")
        raise RuntimeError(f"API가 JSON을 반환하지 않았습니다: {preview}") from exc

    items = find_item_list(payload) or []
    rows = []

    for item in items:
        title = pick(
            item,
            ["pblancNm", "rcritPblancNm", "pblancTitle", "title", "hsmpNm"],
            [("pblanc", "nm"), ("rcrit", "nm"), ("title",)],
        )
        district = pick(
            item,
            ["signguNm", "sggNm", "cityNm", "insttAdres"],
            [("signgu", "nm"), ("sgg", "nm")],
        )
        housing_type = pick(
            item,
            ["suplyTyNm", "houseTyNm", "rentTyNm", "bsnsTyNm"],
            [("suply", "nm"), ("house", "ty"), ("rent", "ty")],
        )
        start_raw = pick(
            item,
            ["rcritBeginDe", "pblancBeginDe", "beginDe", "pblancDe", "registDt"],
            [("begin", "de"), ("start",), ("regist", "dt")],
        )
        end_raw = pick(
            item,
            ["rcritEndDe", "pblancEndDe", "endDe", "rceptClosDe", "closDe"],
            [("end", "de"), ("clos", "de"), ("deadline",)],
        )
        link = pick(
            item,
            ["dtlUrl", "pblancUrl", "hmpgAdres", "detailUrl", "url", "link"],
            [("url",), ("hmpg", "adres")],
        )

        if title:
            rows.append(
                {
                    "지역": region_name,
                    "시군구": district,
                    "공고명": title,
                    "주택유형": housing_type,
                    "공고일": start_raw,
                    "마감일": end_raw,
                    "원문 링크": link,
                }
            )
    return rows


def collect_data():
    api_key = os.environ.get("PUBLIC_DATA_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("GitHub Secret PUBLIC_DATA_API_KEY가 없습니다.")

    rows = []
    for code, name in REGIONS.items():
        rows.extend(fetch_region(api_key, code, name))

    if not rows:
        raise RuntimeError("서울·경기 공고를 한 건도 받지 못해 기존 파일을 유지합니다.")

    df = pd.DataFrame(rows).drop_duplicates(subset=["지역", "공고명", "원문 링크"])
    today = pd.Timestamp.now(tz="Asia/Seoul").tz_localize(None).normalize()
    df["마감일_날짜"] = df["마감일"].apply(parse_date)
    df["D-Day"] = (df["마감일_날짜"] - today).dt.days
    df["상태"] = df["D-Day"].apply(
        lambda days: "마감일 미정"
        if pd.isna(days)
        else ("마감" if days < 0 else ("오늘 마감" if days == 0 else f"D-{int(days)}"))
    )
    df = df[(df["D-Day"].isna()) | (df["D-Day"] >= 0)].copy()
    df = df.sort_values(["D-Day", "지역", "공고명"], na_position="last")
    return df


def save_excel(df):
    export_df = df.drop(columns=["마감일_날짜"])
    export_df.to_excel("housing_notices.xlsx", index=False)


def save_map(df):
    housing_map = folium.Map(location=[37.55, 127.05], zoom_start=8)
    for region, (lat, lon) in REGION_COORDS.items():
        region_df = df[df["지역"] == region]
        folium.Marker(
            [lat, lon],
            tooltip=f"{region} 공고 {len(region_df)}건",
            popup=folium.Popup(f"<b>{region}</b><br>진행 중 공고 {len(region_df)}건", max_width=260),
            icon=folium.Icon(color="blue" if region == "서울" else "green", icon="home"),
        ).add_to(housing_map)
    housing_map.save("housing_map.html")


def save_html(df):
    updated_at = datetime.now().astimezone().strftime("%Y-%m-%d %H:%M")
    cards = []
    for _, row in df.iterrows():
        title = html.escape(str(row["공고명"]))
        region = html.escape(" ".join(filter(None, [str(row["지역"]), str(row["시군구"])])))
        kind = html.escape(str(row["주택유형"] or "유형 미표기"))
        deadline = html.escape(str(row["마감일"] or "미정"))
        status = html.escape(str(row["상태"]))
        link = str(row["원문 링크"] or "").strip()
        link_html = (
            f'<a class="button" href="{html.escape(link, quote=True)}" target="_blank" rel="noopener">공고 원문 보기</a>'
            if link.startswith(("http://", "https://"))
            else '<span class="button disabled">원문 링크 없음</span>'
        )
        urgent = " urgent" if pd.notna(row["D-Day"]) and 0 <= row["D-Day"] <= 7 else ""
        search_text = html.escape(f"{title} {region} {kind}".lower(), quote=True)
        cards.append(
            f'''<article class="card{urgent}" data-search="{search_text}">
              <div class="meta"><span>{region}</span><span>{kind}</span><strong>{status}</strong></div>
              <h2>{title}</h2>
              <p>접수 마감: {deadline}</p>
              {link_html}
            </article>'''
        )

    document = f'''<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>서울·경기 공공주택 뉴스레터</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; background: #f4f6f8; color: #17212b; font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Noto Sans KR", sans-serif; }}
    header {{ padding: 44px 20px 30px; color: white; background: linear-gradient(135deg, #153e75, #2563eb); }}
    header div, main {{ max-width: 960px; margin: auto; }}
    h1 {{ margin: 0 0 10px; font-size: clamp(28px, 5vw, 44px); }}
    header p {{ margin: 5px 0; opacity: .9; }}
    .toolbar {{ display: grid; grid-template-columns: 1fr auto auto; gap: 10px; margin: 22px 0; }}
    input, .toolbar a {{ min-height: 45px; border: 1px solid #d8dee8; border-radius: 10px; padding: 11px 14px; background: white; color: #17345f; text-decoration: none; }}
    .grid {{ display: grid; gap: 14px; padding-bottom: 50px; }}
    .card {{ background: white; border: 1px solid #e3e8ef; border-radius: 16px; padding: 20px; box-shadow: 0 5px 18px rgba(15, 23, 42, .06); }}
    .card.urgent {{ border-left: 6px solid #ef4444; }}
    .meta {{ display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }}
    .meta span, .meta strong {{ padding: 5px 9px; border-radius: 999px; background: #eaf1ff; font-size: 13px; }}
    .meta strong {{ background: #fff1f1; color: #c53030; }}
    h2 {{ font-size: 19px; line-height: 1.45; margin: 14px 0 8px; }}
    .button {{ display: inline-block; margin-top: 5px; padding: 9px 13px; border-radius: 9px; background: #1d4ed8; color: white; text-decoration: none; font-weight: 700; }}
    .disabled {{ background: #9ca3af; }}
    .empty {{ display: none; text-align: center; padding: 50px; }}
    @media (max-width: 650px) {{ .toolbar {{ grid-template-columns: 1fr 1fr; }} .toolbar input {{ grid-column: 1 / -1; }} main {{ padding: 0 14px; }} }}
  </style>
</head>
<body>
  <header><div><h1>서울·경기 공공주택 뉴스레터</h1><p>진행 중 공고 {len(df)}건</p><p>마지막 자동 갱신: {updated_at} (한국시간)</p></div></header>
  <main>
    <div class="toolbar">
      <input id="search" type="search" placeholder="지역·공고명·주택유형 검색">
      <a href="housing_notices.xlsx" download>엑셀 받기</a>
      <a href="housing_map.html">지도 보기</a>
    </div>
    <section class="grid" id="cards">{''.join(cards)}</section>
    <p class="empty" id="empty">검색 결과가 없습니다.</p>
  </main>
  <script>
    const input = document.getElementById('search');
    const cards = [...document.querySelectorAll('.card')];
    input.addEventListener('input', () => {{
      const keyword = input.value.trim().toLowerCase();
      let visible = 0;
      cards.forEach(card => {{
        const show = card.dataset.search.includes(keyword);
        card.style.display = show ? '' : 'none';
        if (show) visible++;
      }});
      document.getElementById('empty').style.display = visible ? 'none' : 'block';
    }});
  </script>
</body>
</html>'''
    with open("index.html", "w", encoding="utf-8") as file:
        file.write(document)


def main():
    df = collect_data()
    save_excel(df)
    save_map(df)
    save_html(df)
    print(f"자동 갱신 완료: {len(df)}건")


if __name__ == "__main__":
    main()
