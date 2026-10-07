# -*- coding: utf-8 -*-
"""
금형 예방점검(자주보전) 알림이 왜 안 뜨는지 MES를 '조회만' 해서 조사합니다. (MES는 바꾸지 않음)

- server.py 와 같은 폴더에서 실행:  py pm_check.py   (또는 pm_check.bat 더블클릭)
- 결과: 같은 폴더에 '예방점검_조사_날짜.xlsx' 가 생깁니다. 이 파일을 보내 주세요.
"""
import os
import sys
import glob
import datetime as dt
import importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
MOLD = sys.argv[1] if len(sys.argv) > 1 else "08DHA086"   # 화면에 보인 금형 (다른 금형: py pm_check.py 금형코드)

# 웹 서버 파일을 불러와서 같은 DB 설정·접속을 그대로 씀
srv_file = None
for p in glob.glob(os.path.join(HERE, "*.py")):
    if os.path.abspath(p) == os.path.abspath(__file__):
        continue
    try:
        with open(p, "rb") as f:
            if b"def quit_watch" in f.read():
                srv_file = p
                break
    except Exception:
        pass
if not srv_file:
    sys.exit("웹 서버 파일(server.py)을 이 파일과 같은 폴더에 두세요.")
sys.argv = [srv_file]
spec = importlib.util.spec_from_file_location("websrv", srv_file)
S = importlib.util.module_from_spec(spec)
spec.loader.exec_module(S)
APP, MES, openpyxl = S.APP, S.MES, S.openpyxl
if openpyxl is None:
    sys.exit("openpyxl 이 필요합니다: py -m pip install openpyxl")

OWNER = MES.qualify("X").split(".")[0] if "." in MES.qualify("X") else None
KW_COL = ("CYCLE", "PERIOD", "INTERVAL", "CHECK", "PM_", "PREV", "ALARM", "ALERT", "NOTI", "SHOT", "HIT", "STROKE",
          "INSP", "MAINT", "LIMIT", "TABAL", "JUGI")
KW_CMT = ("주기", "점검", "보전", "타발", "알림", "알람", "경고", "누적", "타수", "샷")
KW_NAME = ("PM", "PREV", "MAINT", "INSP", "CHECK", "CYCLE", "ALARM", "ALERT", "NOTI", "SHOT", "BOJEON", "JUGI")

sheets = {}     # 시트이름 -> (머리글, 줄들)
notes = []      # 요약


def say(t):
    print(t, flush=True)
    notes.append([t])


def q(cur, sql, **b):
    try:
        cur.execute(sql, b)
        names = [d[0] for d in cur.description]
        return names, cur.fetchall()
    except Exception as e:
        return ["오류"], [[str(e).splitlines()[0]]]


def text(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    try:
        if hasattr(v, "read"):
            v = v.read()
    except Exception:
        pass
    return str(v)[:500]


def work(conn):
    cur = conn.cursor()
    own = "NVL(:o, USER)"
    say(f"[조사 시각] {dt.datetime.now():%Y-%m-%d %H:%M}   스키마: {OWNER or '(접속 계정)'}   기준 금형: {MOLD}")

    # 1) 후보 테이블: 컬럼명/주석/테이블명/테이블주석에 주기·점검·보전·타발 등
    conds = " OR ".join([f"C.COLUMN_NAME LIKE '%{k}%'" for k in KW_COL] +
                        [f"CC.COMMENTS LIKE '%{k}%'" for k in KW_CMT])
    names, rows = q(cur, f"""
        SELECT C.TABLE_NAME, C.COLUMN_NAME, C.DATA_TYPE, CC.COMMENTS
          FROM ALL_TAB_COLUMNS C
          LEFT JOIN ALL_COL_COMMENTS CC ON CC.OWNER = C.OWNER AND CC.TABLE_NAME = C.TABLE_NAME AND CC.COLUMN_NAME = C.COLUMN_NAME
         WHERE C.OWNER = {own} AND ({conds})
           AND C.TABLE_NAME IN (SELECT TABLE_NAME FROM ALL_TAB_COLUMNS WHERE OWNER = {own} AND COLUMN_NAME LIKE '%MOLD%')
         ORDER BY C.TABLE_NAME, C.COLUMN_NAME""", o=OWNER)
    sheets["1_후보컬럼"] = (["테이블", "컬럼", "형식", "주석"], rows)
    tabs = sorted({r[0] for r in rows if names != ["오류"]})
    tconds = " OR ".join([f"T.TABLE_NAME LIKE '%{k}%'" for k in KW_NAME] + [f"TC.COMMENTS LIKE '%{k}%'" for k in KW_CMT])
    n2, r2 = q(cur, f"""
        SELECT T.TABLE_NAME, T.NUM_ROWS, TC.COMMENTS FROM ALL_TABLES T
          LEFT JOIN ALL_TAB_COMMENTS TC ON TC.OWNER = T.OWNER AND TC.TABLE_NAME = T.TABLE_NAME
         WHERE T.OWNER = {own} AND ({tconds}) ORDER BY T.TABLE_NAME""", o=OWNER)
    sheets["2_후보테이블"] = (["테이블", "줄수(통계)", "주석"], r2)
    if n2 != ["오류"]:
        tabs = sorted(set(tabs) | {r[0] for r in r2})
    say(f"[1] 예방점검·타발 관련으로 보이는 테이블 {len(tabs)}개: {', '.join(tabs[:30])}{' ...' if len(tabs) > 30 else ''}")

    # 2) 기준 금형의 줄 (금형 컬럼이 있는 후보 테이블마다)
    rows3 = []
    for t in tabs:
        _, cols = q(cur, f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE OWNER = {own} AND TABLE_NAME = :t ORDER BY COLUMN_ID", o=OWNER, t=t)
        cols = [c[0] for c in cols]
        mc = next((c for c in cols if c == "MOLD_CODE"), None) or next((c for c in cols if "MOLD" in c and "CODE" in c), None)
        if not mc:
            continue
        nm, rr = q(cur, f"SELECT * FROM (SELECT * FROM {MES.qualify(t)} WHERE TRIM({mc}) = :m) WHERE ROWNUM <= 50", m=MOLD)
        if nm == ["오류"] or not rr:
            continue
        rows3.append([f"■ {t}  ({len(rr)}줄)"])
        rows3.append(nm)
        rows3 += [[text(v) for v in r] for r in rr]
        rows3.append([])
    sheets["3_기준금형_자료"] = ([f"금형 {MOLD} 가 들어 있는 줄 (후보 테이블별)"], rows3)
    say(f"[2] 금형 {MOLD} 자료가 있는 후보 테이블: {sum(1 for r in rows3 if r and str(r[0]).startswith('■'))}개 (3번 시트)")

    # 3) 이 테이블들을 쓰는 트리거·프로그램 (알림/누적타발을 만드는 곳)
    trows = []
    for t in tabs:
        _, tr = q(cur, f"SELECT TRIGGER_NAME, TRIGGERING_EVENT, STATUS, TABLE_NAME FROM ALL_TRIGGERS WHERE TABLE_OWNER = {own} AND TABLE_NAME = :t",
                  o=OWNER, t=t)
        trows += [list(r) for r in tr if len(r) == 4]
    sheets["4_트리거"] = (["트리거", "동작", "상태(ENABLED/DISABLED)", "테이블"], trows)
    dis = [r[0] for r in trows if str(r[2]).upper() != "ENABLED"]
    say(f"[3] 후보 테이블의 트리거 {len(trows)}개" + (f" / ※ 꺼져 있음(DISABLED): {', '.join(dis)}" if dis else ""))

    prow = []
    for t in tabs[:40]:
        _, pr = q(cur, f"""SELECT NAME, TYPE, MIN(LINE) FROM ALL_SOURCE WHERE OWNER = {own}
                            AND UPPER(TEXT) LIKE '%' || :t || '%' GROUP BY NAME, TYPE""", o=OWNER, t=t)
        prow += [[t] + list(r) for r in pr if len(r) == 3]
    sheets["5_사용프로그램"] = (["테이블", "프로그램", "종류", "처음 나오는 줄"], prow)
    n6, r6 = q(cur, f"""SELECT O.OBJECT_NAME, O.OBJECT_TYPE, O.STATUS, O.LAST_DDL_TIME FROM ALL_OBJECTS O
                        WHERE O.OWNER = {own} AND O.STATUS <> 'VALID' ORDER BY 2, 1""", o=OWNER)
    sheets["6_고장난프로그램"] = (n6, [[text(v) for v in r] for r in r6])
    say(f"[4] 후보 테이블을 쓰는 프로그램 {len({r[1] for r in prow})}개 / 컴파일 오류(INVALID) 객체 {len(r6) if n6 != ['오류'] else '?'}개")

    # 4) 자동 실행 작업 (정해진 시간마다 누적타발·알림을 계산하는 JOB 이 멈췄는지)
    jobs = []
    n7, r7 = q(cur, "SELECT JOB, WHAT, LAST_DATE, NEXT_DATE, BROKEN, FAILURES, INTERVAL FROM USER_JOBS")
    if n7 != ["오류"]:
        jobs.append(["[USER_JOBS]"]); jobs.append(n7); jobs += [[text(v) for v in r] for r in r7]
    n8, r8 = q(cur, f"SELECT JOB, WHAT, LAST_DATE, NEXT_DATE, BROKEN, FAILURES, INTERVAL FROM ALL_JOBS WHERE SCHEMA_USER = {own}", o=OWNER)
    if n8 != ["오류"]:
        jobs.append([]); jobs.append(["[ALL_JOBS]"]); jobs.append(n8); jobs += [[text(v) for v in r] for r in r8]
    n9, r9 = q(cur, f"""SELECT JOB_NAME, ENABLED, STATE, LAST_START_DATE, NEXT_RUN_DATE, FAILURE_COUNT, REPEAT_INTERVAL, JOB_ACTION
                        FROM ALL_SCHEDULER_JOBS WHERE OWNER = {own}""", o=OWNER)
    if n9 != ["오류"]:
        jobs.append([]); jobs.append(["[ALL_SCHEDULER_JOBS]"]); jobs.append(n9); jobs += [[text(v) for v in r] for r in r9]
    sheets["7_자동실행JOB"] = (["DB 자동 실행 작업"], jobs)
    broken = [r for r in (r7 if n7 != ["오류"] else []) + (r8 if n8 != ["오류"] else []) if str(r[4]).upper() == "Y" or (r[5] or 0) > 0]
    njob = (len(r7) if n7 != ["오류"] else 0) + (len(r9) if n9 != ["오류"] else 0)
    say(f"[5] DB 자동 실행 작업(JOB): {njob}개"
        + (f" / ※ 멈춤·실패: {len(broken)}개 → {', '.join(text(r[1])[:40] for r in broken[:5])}" if broken else ""))

    # 5) 실제 생산 타발수 (최근 기록) - 누적타발수가 쌓이고 있는지 비교용
    n10, r10 = q(cur, f"""SELECT TABLE_NAME, COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE OWNER = {own}
                          AND (COLUMN_NAME LIKE '%SHOT%' OR COLUMN_NAME LIKE '%HIT%' OR COLUMN_NAME LIKE '%STROKE%')
                          ORDER BY 1, 2""", o=OWNER)
    shot = []
    for t, c in (r10 if n10 != ["오류"] else [])[:60]:
        _, cols = q(cur, f"SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS WHERE OWNER = {own} AND TABLE_NAME = :t", o=OWNER, t=t)
        cols = [x[0] for x in cols]
        mc = "MOLD_CODE" if "MOLD_CODE" in cols else None
        if not mc:
            shot.append([t, c, "(금형코드 컬럼 없음)", "", ""])
            continue
        _, s = q(cur, f"SELECT COUNT(*), SUM({c}), MAX({c}) FROM {MES.qualify(t)} WHERE TRIM({mc}) = :m", m=MOLD)
        if s and len(s[0]) == 3:
            shot.append([t, c, text(s[0][0]), text(s[0][1]), text(s[0][2])])
        else:
            shot.append([t, c, text(s[0][0]) if s else "", "", ""])
    sheets["8_타발수_비교"] = (["테이블", "타발 컬럼", f"금형 {MOLD} 줄수", "합계", "최대값"], shot)
    say(f"[6] 타발(SHOT) 컬럼이 있는 테이블 {len(shot)}개 - 금형 {MOLD} 의 합계·최대값 (8번 시트)")
    say("    → 화면의 '누적타발수'(지금 2) 가 실제 생산 타발수만큼 올라가는 곳이 있는지 보는 용도입니다.")


try:
    APP.run_db(work)
except Exception as e:
    say(f"※ DB 조사 중 오류: {e}")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "요약"
for r in notes:
    ws.append(r)
ws.column_dimensions["A"].width = 140
for name, (head, rows) in sheets.items():
    w = wb.create_sheet(name[:31])
    w.append(head)
    for r in rows:
        w.append([text(v) for v in r])
out = os.path.join(HERE, f"예방점검_조사_{dt.datetime.now():%Y%m%d_%H%M}.xlsx")
wb.save(out)
print("\n결과 파일:", out)
print("이 파일을 보내 주세요. (MES는 조회만 했고 아무것도 바꾸지 않았습니다)")
if os.name == "nt":
    os.system("pause")
