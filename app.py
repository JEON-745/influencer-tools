"""
인플루언서 서칭 대시보드 (Streamlit 버전)
- 기존 HTML 대시보드의 검색/필터/저장/엑셀 기능을 그대로 옮긴 버전입니다.
- API 키는 코드에 넣지 않고 Streamlit Secrets(.streamlit/secrets.toml, 로컬 전용 / 배포 시 Cloud 설정)에서 읽습니다.

로컬 실행:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import re
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 기본 설정
# ============================================================

st.set_page_config(page_title="인플루언서 서칭 대시보드", layout="wide")

YOUTUBE_API_KEY = st.secrets.get("YOUTUBE_API_KEY", "")
YOUTUBE_BASE = "https://www.googleapis.com/youtube/v3"

CATEGORY_LIST = [
    "뷰티", "메이크업", "스킨케어", "패션", "데일리룩", "게임", "e스포츠",
    "요리", "베이킹", "여행", "캠핑", "육아", "반려동물", "운동/헬스",
]

CONTENT_TYPE_KEYWORDS = {
    "뷰티": ["뷰티", "화장품", "코스메틱", "뷰티템", "beauty"],
    "메이크업": ["메이크업", "화장법", "파운데이션", "아이섀도", "makeup"],
    "스킨케어": ["스킨케어", "피부관리", "클렌징", "세럼", "스킨"],
    "패션": ["패션", "스타일링", "코디", "ootd", "옷"],
    "데일리룩": ["데일리룩", "룩북", "데일리"],
    "브이로그": ["브이로그", "vlog", "일상"],
    "건강": ["건강", "웰빙", "웰에이징", "다이어트"],
    "리빙": ["리빙", "살림", "인테리어", "정리"],
    "푸드": ["푸드", "먹방", "맛집"],
    "게임": ["게임", "겜", "플레이", "gaming"],
    "e스포츠": ["e스포츠", "esports", "롤", "발로란트"],
    "요리": ["요리", "레시피", "쿠킹", "cooking"],
    "베이킹": ["베이킹", "제빵", "빵"],
    "여행": ["여행", "트래블", "travel"],
    "캠핑": ["캠핑", "차박", "camping"],
    "육아": ["육아", "아기", "육아템", "맘"],
    "반려동물": ["반려동물", "강아지", "고양이", "펫"],
    "운동/헬스": ["운동", "헬스", "홈트", "필라테스", "요가"],
}

AUDIENCE_KEYWORDS = [
    ("20대", ["20대", "이십대"]),
    ("30대", ["30대", "삼십대"]),
    ("40대 이상", ["40대", "50대", "중년", "웰에이징", "안티에이징"]),
    ("직장인", ["직장인", "회사원", "오피스"]),
    ("대학생", ["대학생", "캠퍼스"]),
    ("주부/엄마", ["주부", "엄마", "육아맘", "워킹맘"]),
    ("남성", ["남자", "남성", "그루밍"]),
    ("여성", ["여자", "여성", "언니", "누나"]),
    ("커플", ["커플", "연인"]),
]

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
INSTAGRAM_RE = re.compile(r"instagram\.com/([a-zA-Z0-9._]+)")
HANGUL_RE = re.compile(r"[\uac00-\ud7a3]")


# ============================================================
# 세션 상태 초기화
# ============================================================

if "data" not in st.session_state:
    st.session_state.data = []          # 검색된 인플루언서 전체 (dict 리스트)
if "saved_lists" not in st.session_state:
    st.session_state.saved_lists = []   # [{name, created_at, rows: [...]}]
if "history" not in st.session_state:
    st.session_state.history = []       # [{keyword, min_subs, max_subs, min_engagement, result_count, when}]
if "selected_channel_id" not in st.session_state:
    st.session_state.selected_channel_id = None


# ============================================================
# 분석 헬퍼 (한국인 판별 / 콘텐츠유형 / 특징 요약 / 건강도)
# ============================================================

def is_korean_channel(title, description, country, threshold=0.15):
    if country == "KR":
        return True
    text = f"{title} {description}"
    if not text.strip():
        return False
    hangul_count = len(HANGUL_RE.findall(text))
    return (hangul_count / max(len(text), 1)) >= threshold


def infer_content_type(title, description):
    text = f"{title or ''} {description or ''}".lower()
    scores = []
    for tag, keywords in CONTENT_TYPE_KEYWORDS.items():
        score = sum(1 for kw in keywords if kw.lower() in text)
        if score > 0:
            scores.append((tag, score))
    scores.sort(key=lambda x: -x[1])
    top = [t for t, _ in scores[:2]]
    return "+".join(top) if top else None


def infer_audience(title, description):
    text = f"{title or ''} {description or ''}".lower()
    return [label for label, kws in AUDIENCE_KEYWORDS if any(kw.lower() in text for kw in kws)]


def subscriber_tier(subs):
    if subs is None:
        return ""
    if subs < 10_000:
        return "마이크로 인플루언서(1만 미만)"
    if subs < 100_000:
        return "미드티어(1만~10만)"
    if subs < 500_000:
        return "매크로(10만~50만)"
    return "메가 인플루언서(50만 이상)"


def generate_features(row):
    parts = []
    mix = row.get("category_mix")
    if mix:
        top = sorted(mix.items(), key=lambda x: -x[1])[:3]
        parts.append("최근 업로드는 " + ", ".join(f"{t} {p}%" for t, p in top) + " 비중으로 구성돼 있어요.")
    elif row.get("content_type"):
        parts.append(f"주로 {row['content_type'].replace('+', ', ')} 콘텐츠를 업로드해요.")
    else:
        parts.append("채널명/소개글만으로는 주력 콘텐츠를 특정하기 어려워요.")

    audience = infer_audience(row.get("title"), row.get("description"))
    if audience:
        parts.append(f"설명에 언급된 내용으로 볼 때 {', '.join(audience)} 타겟으로 보여요.")

    tier = subscriber_tier(row.get("subscriber_count"))
    if tier:
        parts.append(f"구독자 규모는 {tier}예요.")

    return " ".join(parts)


def is_healthy(subs, comment_reactivity):
    threshold = 0.5 if subs <= 10_000 else 0.25 if subs <= 100_000 else 0.1
    return comment_reactivity >= threshold


def is_recently_active(published_at, days=90):
    if not published_at:
        return False
    latest = datetime.fromisoformat(published_at.replace("Z", "+00:00"))
    return (datetime.now(timezone.utc) - latest) <= timedelta(days=days)


def channel_link(row):
    if row.get("platform") == "youtube":
        return f"https://www.youtube.com/channel/{row['channel_id']}"
    handle = row.get("instagram_handle")
    return f"https://www.instagram.com/{handle}/" if handle else None


# ============================================================
# 유튜브 API 호출
# ============================================================

def yt_get(endpoint, params):
    params = {**params, "key": YOUTUBE_API_KEY}
    resp = requests.get(f"{YOUTUBE_BASE}/{endpoint}", params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def search_youtube(keyword, min_subs, max_subs, min_engagement, target_count, status=None):
    """키워드로 유튜브 채널을 검색해 조건에 맞는 채널을 target_count만큼 모을 때까지 페이지네이션."""
    results = []
    page_token = None
    max_pages = 15

    for page in range(max_pages):
        if len(results) >= target_count:
            break
        if status:
            status.update(label=f'"{keyword}" 검색 중... ({len(results)}/{target_count}건 확보, {page + 1}페이지째)')

        params = {
            "q": keyword, "type": "channel", "part": "snippet", "maxResults": 50,
            "relevanceLanguage": "ko", "regionCode": "KR",
        }
        if page_token:
            params["pageToken"] = page_token

        search_data = yt_get("search", params)
        channel_ids = list({it["snippet"]["channelId"] for it in search_data.get("items", [])})
        page_token = search_data.get("nextPageToken")
        if not channel_ids:
            break

        channels_data = yt_get("channels", {
            "id": ",".join(channel_ids), "part": "snippet,statistics,contentDetails",
        })

        for item in channels_data.get("items", []):
            if len(results) >= target_count:
                break
            snippet = item["snippet"]
            description = snippet.get("description", "")
            subs = int(item["statistics"].get("subscriberCount", 0))
            country = snippet.get("country")

            if not (min_subs <= subs <= max_subs):
                continue
            if not is_korean_channel(snippet["title"], description, country):
                continue

            uploads_playlist = item["contentDetails"]["relatedPlaylists"]["uploads"]
            try:
                pl_data = yt_get("playlistItems", {
                    "playlistId": uploads_playlist, "part": "contentDetails", "maxResults": 8,
                })
                videos = [
                    {"videoId": v["contentDetails"]["videoId"],
                     "publishedAt": v["contentDetails"].get("videoPublishedAt")}
                    for v in pl_data.get("items", [])
                ]
            except requests.HTTPError:
                videos = []

            if not videos or not is_recently_active(videos[0]["publishedAt"]):
                continue

            comment_reactivity = 0.0
            category_mix = None
            recent_videos = []
            if videos:
                videos_data = yt_get("videos", {
                    "id": ",".join(v["videoId"] for v in videos), "part": "statistics,snippet",
                })
                items = videos_data.get("items", [])
                rates = []
                tag_counts = {}
                for v in items:
                    comments = int(v["statistics"].get("commentCount", 0))
                    rates.append((comments / subs) * 100 if subs > 0 else 0)
                    v_title = v.get("snippet", {}).get("title", "").lower()
                    for tag, kws in CONTENT_TYPE_KEYWORDS.items():
                        if any(kw.lower() in v_title for kw in kws):
                            tag_counts[tag] = tag_counts.get(tag, 0) + 1
                if rates:
                    comment_reactivity = round(sum(rates) / len(rates), 3)
                total_tagged = sum(tag_counts.values())
                if total_tagged:
                    category_mix = {t: round(c / total_tagged * 100) for t, c in tag_counts.items()}
                recent_videos = [
                    {
                        "videoId": v["id"],
                        "title": v.get("snippet", {}).get("title", ""),
                        "thumbnail": (v.get("snippet", {}).get("thumbnails", {}).get("medium", {}) or {}).get("url", ""),
                    }
                    for v in items[:4]
                ]

            if comment_reactivity < min_engagement:
                continue

            email_match = EMAIL_RE.search(description)
            ig_match = INSTAGRAM_RE.search(description)

            row = {
                "channel_id": item["id"],
                "platform": "youtube",
                "title": snippet["title"],
                "description": description,
                "thumbnail_url": (snippet.get("thumbnails", {}).get("medium", {}) or {}).get("url", ""),
                "country": country,
                "subscriber_count": subs,
                "view_count": int(item["statistics"].get("viewCount", 0)),
                "video_count": int(item["statistics"].get("videoCount", 0)),
                "engagement_rate_pct": comment_reactivity,
                "email": email_match.group(0) if email_match else None,
                "instagram_handle": ig_match.group(1) if ig_match else None,
                "keyword": keyword,
                "category": keyword,
                "category_mix": category_mix or {keyword: 100},
                "recent_videos": recent_videos,
                "searched_at": datetime.now(timezone.utc).isoformat(),
            }
            row["content_type"] = infer_content_type(row["title"], row["description"])
            row["features"] = generate_features(row)
            row["healthy_account"] = is_healthy(subs, comment_reactivity)
            results.append(row)

        if not page_token:
            break

    return results


# ============================================================
# 엑셀 내보내기 (기존 템플릿과 동일한 서식)
# ============================================================

def export_to_excel(rows, title):
    wb = Workbook()
    ws = wb.active
    ws.title = "인플루언서 리스트"
    ws.sheet_view.showGridLines = False

    widths = [2.625, 31.375, 12.75, 10.75, 14.75, 19.0, 55.875, 16.5, 15.125, 28.75, 20.75, 10.75, 18.5, 8.875]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin_white = Side(style="thin", color="FFFFFFFF")
    thin_gray = Side(style="thin", color="FFA6A6A6")

    ws.row_dimensions[2].height = 27.75
    ws["B2"] = title or "인플루언서 서칭 결과"
    ws["B2"].font = Font(name="맑은 고딕", size=14, bold=True, color="FF000000")
    ws["B2"].alignment = Alignment(vertical="center")
    ws["B2"].border = Border(top=thin_white, bottom=thin_white, left=thin_white, right=thin_white)

    ws["M2"] = datetime.now()
    ws["M2"].number_format = "mm-dd-yy"
    ws["M2"].font = Font(name="맑은 고딕", size=10)
    ws["M2"].border = Border(top=thin_white, bottom=thin_white, left=thin_white, right=thin_white)

    ws.row_dimensions[3].height = 23.25
    headers = ["이름", "플랫폼", "카테고리", "채널링크", "주력 컨텐츠 유형", "특징",
               "팔로워/구독자", "댓글 반응도(%)", "이메일", "인스타그램 핸들", "계정 상태", "두 플랫폼 동시 운영"]
    header_fill = PatternFill("solid", fgColor="FF002060")
    for i, h in enumerate(headers, start=2):
        cell = ws.cell(row=3, column=i, value=h)
        cell.font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(top=thin_white, left=thin_white, right=thin_white)

    for r, row in enumerate(rows, start=4):
        link = channel_link(row)
        values = [
            row.get("title"),
            "유튜브" if row.get("platform") == "youtube" else "인스타그램",
            row.get("category"),
            link or "",
            row.get("content_type") or "",
            row.get("features") or "",
            row.get("subscriber_count"),
            row.get("engagement_rate_pct"),
            row.get("email") or "별도 확인이 필요함",
            row.get("instagram_handle") or "",
            "건강함" if row.get("healthy_account") else "저조함",
            "예" if row.get("cross_platform") else "아니오",
        ]
        for i, v in enumerate(values, start=2):
            cell = ws.cell(row=r, column=i, value=v)
            if i == 5 and link:
                cell.hyperlink = link
                cell.font = Font(name="맑은 고딕", size=11, color="FF0563C1", underline="single")
            else:
                cell.font = Font(name="맑은 고딕", size=11)
            cell.alignment = Alignment(horizontal="center", wrap_text=(i == 7))
            cell.border = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)

    last_data_row = 3 + len(rows)
    for r in range(4, last_data_row + 1):
        ws.cell(row=r, column=1).border = Border(left=thin_white, top=thin_white, bottom=thin_white)
        ws.cell(row=r, column=14).border = Border(right=thin_white, top=thin_white, bottom=thin_white)
    close_row = last_data_row + 1
    for c in range(2, 15):
        ws.cell(row=close_row, column=c).border = Border(left=thin_white, right=thin_white, bottom=thin_white)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# UI - 사이드바 네비게이션
# ============================================================

with st.sidebar:
    st.markdown("### Influencer\nTools")
    page = st.radio(
        "메뉴", ["인플루언서 검색", "저장한 리스트", "검색 히스토리"],
        label_visibility="collapsed",
    )

if not YOUTUBE_API_KEY:
    st.warning(
        "⚠️ YouTube API 키가 설정되지 않았어요. `.streamlit/secrets.toml`에 "
        "`YOUTUBE_API_KEY = \"발급받은 키\"` 형식으로 넣어주세요. "
        "(Streamlit Cloud에 배포한 경우 앱 설정 > Secrets에서 등록)"
    )

# ============================================================
# 페이지: 인플루언서 검색
# ============================================================

if page == "인플루언서 검색":
    st.title("크리에이터 검색")
    st.caption(f"유튜브 · 인스타그램 인플루언서를 조건별로 검색하고 확인하세요 (현재 {len(st.session_state.data)}건 보유)")

    with st.container(border=True):
        col1, col2, col3, col4, col5 = st.columns([1.3, 1.3, 1.2, 1.2, 1])
        with col1:
            categories = st.multiselect("카테고리", CATEGORY_LIST, key="f_categories")
        with col2:
            sub_range = st.slider("팔로워/구독자", 0, 1_000_000, (0, 500_000), step=1000, key="f_subs")
        with col3:
            min_engagement = st.slider(
                "최소 댓글 반응도(%)", 0.0, 6.0, 0.0, step=0.1, key="f_engagement",
                help="최근 업로드 영상 최대 8개의 평균 (댓글 수 ÷ 구독자 수 × 100). 좋아요는 포함하지 않습니다.",
            )
        with col4:
            target_count = st.number_input("표시할 계정 수", min_value=1, value=20, step=5, key="f_count")
        with col5:
            st.write("")
            st.write("")
            run = st.button("🔍 조회하기", type="primary", use_container_width=True)

    if run:
        if not categories:
            st.info("카테고리를 1개 이상 선택하면 유튜브에서 실시간으로 검색합니다.")
        elif not YOUTUBE_API_KEY:
            st.error("YouTube API 키가 없어서 검색할 수 없어요. secrets 설정을 먼저 확인해주세요.")
        else:
            with st.status("실시간 검색 중...", expanded=True) as status:
                all_new = []
                for cat in categories:
                    new_rows = search_youtube(
                        cat, sub_range[0], sub_range[1], min_engagement, target_count, status=status
                    )
                    all_new.extend(new_rows)
                # 기존 데이터와 병합 (같은 채널이면 최신 결과로 교체)
                existing_ids = {r["channel_id"] for r in all_new}
                st.session_state.data = [r for r in st.session_state.data if r["channel_id"] not in existing_ids] + all_new
                status.update(label=f"검색 완료 — {len(all_new)}건 추가 (전체 {len(st.session_state.data)}건)", state="complete")

            st.session_state.history.insert(0, {
                "categories": categories, "min_subs": sub_range[0], "max_subs": sub_range[1],
                "min_engagement": min_engagement, "result_count": len(all_new),
                "when": datetime.now().strftime("%Y.%m.%d %H:%M"),
            })

    # ---- 결과 필터링 & 표시 ----
    rows = st.session_state.data
    if categories:
        rows = [r for r in rows if r["category"] in categories]
    rows = [r for r in rows if sub_range[0] <= r["subscriber_count"] <= sub_range[1]]
    rows = [r for r in rows if r["engagement_rate_pct"] >= min_engagement]
    rows = sorted(rows, key=lambda r: r["engagement_rate_pct"], reverse=True)[:target_count]

    c1, c2 = st.columns(2)
    c1.metric("표시 중", f"{len(rows)}개")
    c2.metric("두 플랫폼 동시 운영", f"{sum(1 for r in rows if r.get('cross_platform'))}개")

    col_a, col_b, col_c = st.columns([1, 1, 6])
    with col_a:
        list_name = st.text_input("리스트 이름", placeholder="리스트명 입력", label_visibility="collapsed")
    with col_b:
        if st.button("☆ 리스트에 저장", disabled=not rows):
            name = list_name.strip() or f"리스트 {len(st.session_state.saved_lists) + 1}"
            st.session_state.saved_lists.insert(0, {
                "name": name, "created_at": datetime.now().strftime("%Y.%m.%d"), "rows": rows,
            })
            st.success(f'"{name}" 저장 완료! "저장한 리스트" 탭에서 확인하세요.')
    with col_c:
        if rows:
            excel_buf = export_to_excel(rows, list_name.strip() or "인플루언서 서칭 결과")
            st.download_button(
                "⬇ 엑셀로 내보내기", data=excel_buf,
                file_name=f"{(list_name.strip() or '인플루언서_서칭결과')}_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    if rows:
        df = pd.DataFrame([{
            "이름": r["title"],
            "플랫폼": "유튜브" if r["platform"] == "youtube" else "인스타그램",
            "카테고리": r["category"],
            "채널링크": channel_link(r),
            "주력 콘텐츠 유형": r.get("content_type") or "-",
            "특징": r.get("features") or "-",
            "팔로워/구독자": r["subscriber_count"],
            "댓글 반응도(%)": r["engagement_rate_pct"],
            "이메일": r.get("email") or "별도 확인이 필요함",
            "계정 상태": "건강함" if r["healthy_account"] else "저조함",
        } for r in rows])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("조건에 맞는 인플루언서가 없습니다. 필터를 조정해보세요.")

# ============================================================
# 페이지: 저장한 리스트
# ============================================================

elif page == "저장한 리스트":
    st.title("저장한 리스트")
    st.caption("캠페인별로 저장해둔 인플루언서 후보 리스트를 관리하세요")

    if not st.session_state.saved_lists:
        st.info("아직 저장한 리스트가 없어요. '인플루언서 검색' 탭에서 결과를 저장해보세요.")
    else:
        for i, lst in enumerate(st.session_state.saved_lists):
            with st.container(border=True):
                col1, col2, col3 = st.columns([4, 1, 1])
                col1.markdown(f"**{lst['name']}**  \n인플루언서 {len(lst['rows'])}명 · {lst['created_at']}")
                excel_buf = export_to_excel(lst["rows"], lst["name"])
                col2.download_button(
                    "⬇ 내보내기", data=excel_buf, file_name=f"{lst['name']}.xlsx",
                    key=f"export_{i}", use_container_width=True,
                )
                if col3.button("삭제", key=f"delete_{i}", use_container_width=True):
                    st.session_state.saved_lists.pop(i)
                    st.rerun()
                with st.expander("펼쳐보기"):
                    df = pd.DataFrame([{
                        "이름": r["title"], "플랫폼": r["platform"], "채널링크": channel_link(r),
                        "팔로워/구독자": r["subscriber_count"], "댓글 반응도(%)": r["engagement_rate_pct"],
                        "계정 상태": "건강함" if r["healthy_account"] else "저조함",
                    } for r in lst["rows"]])
                    st.dataframe(df, use_container_width=True, hide_index=True)

# ============================================================
# 페이지: 검색 히스토리
# ============================================================

elif page == "검색 히스토리":
    st.title("검색 히스토리")
    st.caption("지금까지 실행한 검색 조건을 확인할 수 있어요")

    if not st.session_state.history:
        st.info("아직 검색 기록이 없어요.")
    else:
        for h in st.session_state.history:
            with st.container(border=True):
                st.markdown(
                    f"**{', '.join(h['categories'])}** · "
                    f"팔로워/구독자 {h['min_subs']:,}~{h['max_subs']:,} · "
                    f"최소 댓글 반응도 {h['min_engagement']}% · "
                    f"결과 {h['result_count']}건 · {h['when']}"
                )
