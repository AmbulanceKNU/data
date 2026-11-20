import json
import random
import re
import time
import uuid

import requests
from openpyxl import Workbook, load_workbook

BASE = "https://knuin.knu.ac.kr"
PAGE_URL = BASE + "/public/stddm/lectPlnInqr.knu"
API_URL = BASE + "/public/web/stddm/lsspr/syllabus/lectPlnInqr/selectListLectPlnInqr"

# ===== 설정 =====
YEAR = "2026"
SEMESTERS = {                       # 이름: 학기 코드 (다른 학기는 같은 방식으로 추가)
    "2학기": "CMBS001400002",
    "계절학기": "CMBS001400004",
}
DEFAULT_SEMESTER = "2학기"

# 전처리에서 삭제할 컬럼
DROP_COLUMNS = [
    "rowIndex", "success", "message", "pageSize", "pageNum",
    "lctreMthodNm", "lctreLnggeSctcd", "gubun", "name", "code",
    "bldngSn", "bldngCd", "bldngNm", "lssnsLcttmUntcd", "sbjetSctcd2",
    "isApi", "rowStatus", "paging", "pageSort",
]
# ================

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# ---------- 1단계: 원본 데이터 수집 ----------
def build_search(semester_code):
    """필터를 모두 비운 '전체 조회' 조건."""
    return {
        "estblYear": YEAR, "estblSmstrSctcd": semester_code,
        "sbjetCd": "", "sbjetNm": "", "crgePrfssNm": "",
        "sbjetRelmCd": "", "sbjetSctcd": "", "estblDprtnCd": "",
        "rmtCrseYn": "", "rprsnLctreLnggeSctcd": "", "flplnCrseYn": "",
        "pstinNtnnvRmtCrseYn": "", "dgGbDstrcRmtCrseYn": "",
        "sugrdEvltnYn": "", "prctsExrmnYn": "",
        "gubun": "01", "isApi": "Y",
        "bldngSn": "", "bldngCd": "", "bldngNm": "",
        "lssnsLcttmUntcd": "", "sbjetSctcd2": "", "contents": "",
        "lctreLnggeSctcd": "ko", "knuFtrDesigYn": "", "cltreHmntsCltreYn": "",
        "sdgCltreYn": "", "rltmCrseYn": "", "riseRmtCrseYn": "", "coRcgnnSbjetYn": "",
    }


def fetch(semester_code):
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9"})
    s.get(PAGE_URL, timeout=30)                  # 세션 쿠키 수신
    time.sleep(random.uniform(2, 4))

    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/json",
        "Origin": BASE, "Referer": PAGE_URL,
        "X-Requested-With": "XMLHttpRequest", "isAjax": "true",
        "headerAuthStr": "undefined", "headerMenuId": "undefined",
        "headerRequestId": f"kuptl-kcore-web.cli.{uuid.uuid4().hex[:12]}.{int(time.time()*1000)}",
    }
    for attempt in range(1, 4):
        try:
            r = s.post(API_URL, json={"search": build_search(semester_code)},
                       headers=headers, timeout=120)
            if r.status_code == 200:
                return r.json()
            print(f"⚠️ 상태코드 {r.status_code} (시도 {attempt}/3)")
            if r.status_code in (401, 403, 429):
                raise SystemExit("❌ 서버가 요청을 제한했습니다. 한참 뒤에 다시 시도하세요.")
        except requests.RequestException as e:
            print(f"⚠️ 네트워크 오류: {e} (시도 {attempt}/3)")
        time.sleep(10 * attempt)
    raise SystemExit("❌ 요청 실패. 잠시 후 다시 시도하세요.")


def find_records(obj):
    """응답 JSON에서 가장 큰 'dict 리스트'(= 강의 목록)를 찾는다."""
    best = []
    if isinstance(obj, list):
        if obj and all(isinstance(x, dict) for x in obj):
            best = obj
        for x in obj:
            r = find_records(x)
            if len(r) > len(best):
                best = r
    elif isinstance(obj, dict):
        for v in obj.values():
            r = find_records(v)
            if len(r) > len(best):
                best = r
    return best


def flatten(d, prefix=""):
    out = {}
    for k, v in d.items():
        key = f"{prefix}{k}"
        if isinstance(v, dict):
            out.update(flatten(v, key + "."))
        elif isinstance(v, list):
            out[key] = json.dumps(v, ensure_ascii=False)
        else:
            out[key] = v
    return out


def clean(v):
    if isinstance(v, str):
        v = re.sub(r"<br\s*/?>", " ", v)         # 원본 그대로 두려면 이 줄 삭제
        return ILLEGAL.sub("", v).strip()
    return v


def safe_save(wb, path):
    try:
        wb.save(path)
    except PermissionError:                      # 엑셀로 열어 둔 경우
        path = path.replace(".xlsx", f"_{int(time.time())}.xlsx")
        wb.save(path)
    return path


def save_raw_excel(rows, sheet_title, path):
    """서버 응답 그대로(컬럼 이름/순서 유지) 원본 엑셀 저장."""
    cols = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.append(cols)
    for r in rows:
        ws.append([clean(r.get(c)) for c in cols])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return safe_save(wb, path), len(cols)


# ---------- 2단계: 전처리 (불필요한 컬럼 삭제) ----------
def preprocess_excel(raw_path, out_path, drop_columns):
    """원본 엑셀을 읽어서 drop_columns에 있는 컬럼을 삭제하고 새 파일로 저장."""
    wb = load_workbook(raw_path)
    ws = wb.active

    header = [c.value for c in ws[1]]
    drop_set = set(drop_columns)
    # 뒤에서부터 지워야 앞쪽 열 번호가 밀리지 않는다
    targets = [i for i, name in enumerate(header, start=1) if name in drop_set]
    removed = [header[i - 1] for i in targets]
    for idx in reversed(targets):
        ws.delete_cols(idx)

    missing = [c for c in drop_columns if c not in header]
    ws.auto_filter.ref = ws.dimensions           # 열이 줄었으니 필터 범위 갱신
    ws.freeze_panes = "A2"
    out = safe_save(wb, out_path)
    return out, removed, missing, ws.max_row - 1, ws.max_column


def main():
    name = input(f"학기를 입력하세요 ({', '.join(SEMESTERS)}) [Enter={DEFAULT_SEMESTER}]: ").strip()
    if name not in SEMESTERS:
        name = DEFAULT_SEMESTER

    # 1단계: 수집
    print(f"\n📡 [{YEAR} {name}] 전체 강의 조회 요청 1회 전송...")
    data = fetch(SEMESTERS[name])
    with open(f"raw_{name}_전체.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    records = find_records(data)
    if not records:
        raise SystemExit("❌ 강의 목록이 비어 있습니다. raw 파일의 구조를 확인하세요.")

    rows = [flatten(r) for r in records]
    raw_path, ncols = save_raw_excel(rows, f"{name}_전체", f"knu_{name}_전체_원본.xlsx")
    print(f"✅ [원본] {len(rows)}행 × {ncols}열 → {raw_path}")
    if len(rows) % 100 == 0:
        print("⚠️ 행 수가 100의 배수입니다. 서버가 100건씩 잘라 보냈을 수 있습니다. "
              "실제 강의 수와 비교해 보세요.")

    # 2단계: 전처리
    out_path, removed, missing, nrows, nleft = preprocess_excel(
        raw_path, f"knu_{name}_전체.xlsx", DROP_COLUMNS)
    print(f"🧹 [전처리] {len(removed)}개 컬럼 삭제 → {nrows}행 × {nleft}열 → {out_path}")
    if missing:
        print(f"   ℹ️ 원본에 없어서 건너뛴 컬럼: {', '.join(missing)}")

    print("\n🎉 완료")


if __name__ == "__main__":
    main()