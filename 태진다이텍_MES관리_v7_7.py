# -*- coding: utf-8 -*-
"""
태진다이텍 MES 관리 도구 v7.7
- 금형·품번 신규등록 (엑셀 -> Oracle MES) / 기존 금형·품번 보기·수정·삭제
- 품목 단위중량, 작업지시 확인·새로 만들기·수정 (간편 입력창)
- 소재 관리: 입고 조회·새 입고·수정·삭제 (달력 날짜 선택)
- 작업지시 구조 조사 / 전후 비교 (읽기 전용)
필요: pip install oracledb openpyxl
"""
import os
import re
import sys
import csv
import copy
import json
import datetime as dt
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# exe(PyInstaller)로 만들 때 oracledb가 안에서 쓰는 기본 모듈이 빠지지 않도록 미리 불러둠
import getpass, secrets, uuid, hashlib, hmac, ssl, socket, select, asyncio, decimal, struct  # noqa: F401,E401
import queue, ipaddress, platform, locale, zlib, random, string, base64, threading, array  # noqa: F401,E401
import pickle, contextvars, numbers, collections, functools, enum, glob, getpass  # noqa: F401,E401

IMPORT_ERR = []
try:
    import oracledb
except Exception as _e:
    oracledb = None
    IMPORT_ERR.append(f"oracledb: {_e!r}")
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter
except Exception as _e:
    openpyxl = None
    IMPORT_ERR.append(f"openpyxl: {_e!r}")

APP_TITLE = "태진다이텍 MES 관리 도구 v7.7"
BASE = os.path.dirname(os.path.abspath(sys.argv[0]))
CFG_PATH = os.path.join(BASE, "mold_config.json")
LOG_PATH = os.path.join(BASE, "mold_register_log.csv")
BACKUP_PATH = os.path.join(BASE, "mold_deleted_backup.jsonl")   # 삭제한 줄 전체 백업

PARTS = ("mold", "item")
PART_NAME = {"mold": "금형", "item": "품번"}
SHEET = {"mold": "금형", "item": "품번"}
LIST_SHEET = "목록"
TEMPLATE_ROWS = 500
PREVIEW_ROWS = 500
EXACT_COUNT_LIMIT = 200000   # 통계 행수가 이보다 적은 테이블만 실제 COUNT

C_ERR, C_WARN, C_OK, C_INFO = "FFC7CE", "FFEB9C", "C6EFCE", "DDEBF7"

CODE_NAMES = ("MOLD_CODE", "MOLD_CD", "MOLD_NO")
ITEM_NAMES = ("ITEM_CODE", "ITEM_CD", "ITEM_NO")
COPY_NAMES = {"ORGANIZATION_ID", "ORG_ID", "COMPANY_CODE", "COMPANY_ID", "CORP_CODE",
              "PLANT_CODE", "PLANT_ID", "SITE_ID", "SITE_CODE", "FACTORY_CODE", "BIZ_CODE"}
RE_AUTO_DATE = re.compile(r"(ENTER|CREATE|CREATED|CREATION|REG|REGIST|INSERT|INS|LAST_UPDATE|"
                          r"UPDATE|UPDATED|UPD|MODIFY|MOD)_?(DATE|DT|TIME|DTM)$")
RE_AUTO_USER = re.compile(r"(ENTER|CREATE|CREATED|REG|REGIST|INSERT|INS|LAST_UPDATE|"
                          r"UPDATE|UPDATED|UPD|MODIFY|MOD)_?(BY|USER|USER_ID|EMP|EMP_NO)$")
RE_UPD_DATE = re.compile(r"(LAST_UPDATE|UPDATE|UPDATED|UPD|MODIFY|MODIFIED|MOD)_?(DATE|DT|TIME|DTM)$")
RE_UPD_USER = re.compile(r"(LAST_UPDATE|UPDATE|UPDATED|UPD|MODIFY|MODIFIED|MOD)_?(BY|USER|USER_ID|EMP|EMP_NO)$")
ORA_HINT = [
    ("ORA-01400", "필수 컬럼이 비어 있습니다 (아래 괄호 안 마지막 이름이 컬럼)"),
    ("ORA-00001", "같은 값이 이미 있습니다 (중복)"),
    ("ORA-02291", "연결된 기준 데이터(부모 키)가 없습니다"),
    ("ORA-02292", "다른 테이블에서 쓰고 있는 값이라 바꿀 수 없습니다"),
    ("ORA-12899", "값이 너무 깁니다"),
    ("ORA-01722", "숫자 칸에 글자가 들어갔습니다"),
    ("ORA-01031", "이 계정에 입력/수정 권한이 없습니다 (DB 관리자에게 권한 요청)"),
    ("ORA-00942", "테이블이 없거나 볼 권한이 없습니다"),
    ("ORA-01861", "날짜 형식이 맞지 않습니다"),
    ("ORA-00904", "없는 컬럼 이름입니다 (컬럼매핑 확인)"),
    ("ORA-01438", "숫자가 허용 자릿수보다 큽니다"),
]


def ora_hint(e):
    s = str(e)
    for code, msg in ORA_HINT:
        if code in s:
            return f"▶ {msg}\n\n{s}"
    return s


def mk(use, col, name, typ="TEXT", req=False, key=False, maxlen=0, default="", ref="", sim=False, note=""):
    return {"use": use, "col": col, "name": name, "type": typ, "req": req, "key": key,
            "maxlen": maxlen, "default": default, "ref": ref, "sim": sim, "note": note}


DEFAULT_CFG = {
    "oracle": {"user": "", "dsn": "192.168.1.250:1521/XE", "lib_dir": r"C:\instantclient_19_30"},
    "schema": "INFINITY21_PIMMES",
    "reg_user": os.environ.get("USERNAME", ""),
    "ref_mold": "",
    "allow_item_to_existing": False,      # 기존 금형에 품번 줄 추가 허용
    "kor_bytes": 3,
    "limit": 5000,
    "item_master_table": "ICOM_ITEM_MASTER",
    "wo_table": "IPLN_WORK_ORDER_MASTER",
    "bom_table": "ICOM_ITEM_CHILD",               # 품번별 원소재 단위중량 (UNIT_PER_QTY)
    "keywords": "WORK_ORDER, WO_, WORKORDER, PLAN, PROD, BOARD",   # 구조 조사: 찾을 단어
    "small_limit": 30000,                                         # 전후 비교: 줄 단위 비교 한도
    "show_dev_tabs": False,
    "show_reg_tabs": False,
    "show_item_tab": False,                                       # 품목 관리 탭 보이기                                       # 금형 등록 탭(간편등록·엑셀 대량등록) 보이기
    "rcv_weight_unit": "auto",                                    # 소재 입고 순중량 단위 (auto/kg/g)                                       # 고급 탭(구조조사·전후비교) 보이기
    "mold": {
        "table": "ICOM_MOLD", "order_by": "", "code_col": "MOLD_CODE",
        "columns": [
            mk(True, "MOLD_CODE", "금형코드", req=True, key=True, maxlen=30, note="예시 - [DB컬럼 가져오기] 하세요"),
            mk(True, "MOLD_NAME", "금형명", maxlen=100),
        ],
    },
    "item": {
        "enabled": True, "table": "ICOM_MOLD_ITEM", "link_col": "MOLD_CODE",
        "columns": [
            mk(True, "MOLD_CODE", "금형코드", req=True, key=True, maxlen=30, note="예시 - [DB컬럼 가져오기] 하세요"),
            mk(True, "ITEM_CODE", "품번", req=True, key=True, maxlen=30),
        ],
    },
}

# (키, 제목, 폭, 종류)
MAP_FIELDS = [
    ("use", "사용", 45, "yn"), ("col", "DB컬럼", 150, "s"), ("name", "표시이름", 120, "s"),
    ("type", "형식", 55, "type"), ("req", "필수", 45, "yn"), ("key", "중복검사키", 75, "yn"),
    ("maxlen", "최대길이", 65, "int"), ("default", "기본값", 100, "s"),
    ("ref", "존재확인(테이블.컬럼)", 170, "s"), ("sim", "유사경고", 65, "yn"), ("note", "비고", 170, "s"),
]

ID_RE = re.compile(r"^[A-Z_][A-Z0-9_$#]*(\.[A-Z_][A-Z0-9_$#]*){0,2}$")
ORDER_RE = re.compile(r"^[A-Z0-9_$#., ]+$")
CTRL_RE = re.compile(r"[\x00-\x1f\x7f]")
DB_USER = "infinity21_pimmes"          # MES 접속 계정 (고정)
DB_PASS = "infinity21_pimmes"


# ==============================================================
# 공통 도구
# ==============================================================
def ident(s, what="이름"):
    s = (s or "").strip().upper()
    if not ID_RE.match(s):
        raise ValueError(f"{what} 이름이 올바르지 않습니다: '{s}'")
    return s


def blen(s, kb):
    return sum(1 if ord(ch) < 128 else kb for ch in s)


def show(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def cell_text(v, limit=300):
    """화면 표시용 안전 변환 (LOB, 바이트, 제어문자, 너무 긴 글자)"""
    try:
        if v is None:
            return ""
        if hasattr(v, "read"):
            v = v.read(1, limit)
        if isinstance(v, (bytes, bytearray)):
            v = v[:limit // 2].hex().upper()
        s = show(v)
    except Exception as e:
        s = f"<읽기오류 {type(e).__name__}>"
    s = CTRL_RE.sub(" ", s)
    return s if len(s) <= limit else s[:limit] + "…"


def conv(c, v):
    """엑셀 값 -> DB 값. 빈칸은 None, 형식 오류는 ValueError"""
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        if v == "":
            return None
    t = c["type"]
    if t == "NUM":
        if isinstance(v, bool):
            raise ValueError("숫자 아님")
        if isinstance(v, (int, float)):
            return int(v) if float(v).is_integer() else float(v)
        try:
            f = float(str(v).replace(",", ""))
        except ValueError:
            raise ValueError("숫자 아님")
        return int(f) if f.is_integer() else f
    if t == "DATE":
        if isinstance(v, dt.datetime):
            return v
        if isinstance(v, dt.date):
            return dt.datetime(v.year, v.month, v.day)
        s = str(v).strip().replace(".", "-").replace("/", "-").rstrip("-")
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y%m%d"):
            try:
                return dt.datetime.strptime(s, fmt)
            except ValueError:
                pass
        raise ValueError("날짜 아님")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (dt.datetime, dt.date)):
        v = v.strftime("%Y-%m-%d")
    return str(v).strip()


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def load_cfg():
    cfg = copy.deepcopy(DEFAULT_CFG)
    if os.path.exists(CFG_PATH):
        try:
            with open(CFG_PATH, encoding="utf-8") as f:
                saved = json.load(f)
            if "columns" in saved and "mold" not in saved:          # v1 설정 옮기기
                saved["mold"] = {"table": saved.pop("table", ""), "order_by": saved.pop("order_by", ""),
                                 "code_col": "", "columns": saved.pop("columns")}
            for k, v in saved.items():
                if k in ("oracle", "mold", "item") and isinstance(v, dict):
                    cfg[k].update(v)
                elif k in cfg:
                    cfg[k] = v
        except Exception as e:
            messagebox.showwarning(APP_TITLE, f"설정 파일을 읽지 못해 기본값으로 시작합니다.\n{e}")
    cfg["mold"].update(table="ICOM_MOLD", code_col="MOLD_CODE")                 # 고정
    cfg["item"].update(table="ICOM_MOLD_ITEM", link_col="MOLD_CODE", enabled=True)
    return cfg


def save_cfg(cfg):
    with open(CFG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def open_file(path):
    try:
        os.startfile(path)
    except Exception:
        pass


def new_result(rno):
    return {"rno": rno, "vals": {}, "errs": [], "warns": [], "bad": set(), "warncols": set(),
            "status": "", "msg": ""}


def finalize(r):
    r["status"] = "오류" if r["errs"] else ("경고" if r["warns"] else "OK")
    r["msg"] = "OK" if r["status"] == "OK" else f"{r['status']}: " + ", ".join(r["errs"] or r["warns"])


# ==============================================================
# MES (Oracle)
# ==============================================================
class Mes:
    def __init__(self, app):
        self.app = app
        self.pwd = None
        self.client_ready = False
        self._meta = {}
        self._uniq = {}
        self.last_auto = {}

    @property
    def cfg(self):
        return self.app.cfg

    # ---------- 매핑 ----------
    def part(self, p):
        return self.cfg[p]

    def enabled(self, p):
        return p == "mold" or bool(self.cfg["item"].get("enabled"))

    def used(self, p):
        return [c for c in self.part(p)["columns"] if c["use"]]

    def insertable(self, p):
        return [c for c in self.part(p)["columns"] if c["use"] or c["default"]]

    def keys(self, p):
        return [c for c in self.used(p) if c["key"]]

    def code_col(self):
        return (self.cfg["mold"].get("code_col") or "").strip().upper()

    def link_col(self):
        return (self.cfg["item"].get("link_col") or "").strip().upper()

    def qualify(self, name):
        tb = ident(name, "테이블")
        s = (self.cfg["schema"] or "").strip()
        return tb if (not s or "." in tb) else ident(s, "스키마") + "." + tb

    def table(self, p):
        t = (self.part(p).get("table") or "").strip()
        if not t:
            raise ValueError(f"{PART_NAME[p]} 테이블명을 [엑셀칸 설정] 탭에 입력하세요.")
        return self.qualify(t)

    def ref_table(self, ref):
        ref = ident(ref, "존재확인")
        if "." not in ref:
            raise ValueError(f"존재확인은 '테이블.컬럼' 형식으로 쓰세요: {ref}")
        tb, col = ref.rsplit(".", 1)
        return self.qualify(tb), col

    def validate_map(self):
        cc = self.code_col()
        if not cc or cc not in [c["col"] for c in self.used("mold")]:
            raise ValueError(f"금형 매핑에서 금형코드 컬럼({cc or '미지정'})이 사용 Y여야 합니다.")
        if self.enabled("item"):
            lc = self.link_col()
            if not lc or lc not in [c["col"] for c in self.used("item")]:
                raise ValueError(f"품번 매핑에서 금형코드 연결 컬럼({lc or '미지정'})이 사용 Y여야 합니다.")

    # ---------- 접속 ----------
    @staticmethod
    def client_dirs(lib):
        """Instant Client 찾을 곳: 설정값 → 프로그램 폴더 안 → C드라이브 oracle 폴더 → C드라이브 → PATH의 oci.dll 폴더"""
        import glob
        cands = [lib] if lib else []
        for pat in (os.path.join(BASE, "instantclient*"), os.path.join(BASE, "*", "instantclient*"),
                    r"C:\oracle\instantclient*", r"C:\instantclient*", r"D:\instantclient*"):
            cands += sorted(glob.glob(pat), reverse=True)
        for d in os.environ.get("PATH", "").split(os.pathsep):
            if d and os.path.isfile(os.path.join(d, "oci.dll")):
                cands.append(d)
        out = []
        for d in cands:
            if d and os.path.isdir(d) and d not in out:
                out.append(d)
        return out

    def init_client(self, lib):
        """11g 접속에 필요한 Instant Client를 찾아서 켜기 (못 찾으면 어디를 봤는지 알려줌)"""
        if not oracledb.is_thin_mode():
            return
        tried = []
        for d in self.client_dirs(lib):
            try:
                oracledb.init_oracle_client(lib_dir=d)
                if d != lib:
                    self.cfg["oracle"]["lib_dir"] = d
                    try:
                        save_cfg(self.cfg)
                    except Exception:
                        pass
                return
            except Exception as e:
                tried.append(f"  {d}\n    → {str(e).splitlines()[0][:150]}")
        raise RuntimeError(
            "Oracle Instant Client를 찾지 못했습니다 (11g 접속에 꼭 필요).\n\n"
            "해결: instantclient_19_xx 폴더를 이 프로그램과 같은 폴더에 넣거나 C:\\ 에 두세요.\n"
            "그래도 안 되면 Visual C++ 재배포 패키지(2017 이상, x64)를 설치하세요.\n\n"
            + ("찾아본 곳:\n" + "\n".join(tried) if tried else
               f"찾아본 곳: 설정값({lib or '없음'}), {BASE}\\instantclient*, C:\\instantclient*, C:\\oracle\\instantclient*"))

    def connect(self):
        if oracledb is None:
            raise RuntimeError("oracledb가 없습니다.\npip install oracledb")
        o = self.cfg["oracle"]
        if not getattr(self, "client_ready", getattr(self, "ready", False)):
            self.init_client((o.get("lib_dir") or "").strip())
            self.client_ready = self.ready = True
        try:
            return oracledb.connect(user=DB_USER, password=DB_PASS, dsn=o["dsn"])
        except Exception as e:
            raise RuntimeError(f"MES 접속 실패\n{e}")

    # ---------- 조회 ----------
    def desc_defs(self, p, description):
        """조회 결과 컬럼 정보 -> 컬럼 정의 (매핑에 있으면 매핑 내용 우선)"""
        by = {c["col"]: c for c in self.part(p)["columns"]}
        defs = {}
        for d in description:
            name, t, isize, null_ok = d[0], d[1], d[3], d[6]
            tn = str(getattr(t, "name", t)).upper()
            if "LOB" in tn or "RAW" in tn or "LONG" in tn or "ROWID" in tn:
                typ = None
            elif "CHAR" in tn:
                typ = "TEXT"
            elif any(w in tn for w in ("NUMBER", "BINARY_", "FLOAT", "INTEGER", "DECIMAL")):
                typ = "NUM"
            elif "DATE" in tn or "TIMESTAMP" in tn:
                typ = "DATE"
            else:
                typ = None
            m = by.get(name)
            c = mk(True, name, (m["name"] if m else name), typ or "TEXT",
                   req=(null_ok is False) or bool(m and m["req"]),
                   key=bool(m and m["key"] and m["use"]),
                   maxlen=(m["maxlen"] if m and m["maxlen"] else (int(isize or 0) if typ == "TEXT" else 0)),
                   ref=(m["ref"] if m else ""))
            c["editable"] = typ is not None
            c["dbtype"] = tn.replace("DB_TYPE_", "")
            defs[name] = c
        return defs

    @staticmethod
    def reorder(cols, rows, first):
        order = []
        for c in first:
            if c in cols and cols.index(c) not in order:
                order.append(cols.index(c))
        order += [i for i in range(len(cols)) if i not in order]
        return [cols[i] for i in order], [[r[i] for i in order] for r in rows]

    def fetch_old(self, conn):
        sql = f"SELECT * FROM {self.table('mold')}"
        ob = (self.cfg["mold"].get("order_by") or "").strip().upper()
        if ob:
            if not ORDER_RE.match(ob):
                raise ValueError(f"조회 정렬 값이 올바르지 않습니다: {ob}")
            sql += " ORDER BY " + ob
        cur = conn.cursor()
        cur.execute(f"SELECT * FROM ({sql}) WHERE ROWNUM <= :lim", lim=int(self.cfg.get("limit") or 5000))
        rows = cur.fetchall()
        defs = self.desc_defs("mold", cur.description)
        first = [self.code_col()] + [c["col"] for c in self.used("mold")]
        cols, rows = self.reorder([d[0] for d in cur.description], rows, first)
        return cols, rows, defs

    def fetch_items_of(self, conn, code):
        cur = conn.cursor()
        cur.execute(f"SELECT * FROM {self.table('item')} WHERE {ident(self.link_col(), '연결컬럼')} = :v", v=code)
        rows = cur.fetchall()
        defs = self.desc_defs("item", cur.description)
        first = [self.link_col()] + [c["col"] for c in self.used("item")]
        cols, rows = self.reorder([d[0] for d in cur.description], rows, first)
        return cols, rows, defs

    def fetch_columns(self, conn, table, p):
        sql = """
        SELECT c.COLUMN_NAME, c.DATA_TYPE, c.DATA_LENGTH, c.NULLABLE,
          (SELECT 'Y' FROM ALL_CONS_COLUMNS cc JOIN ALL_CONSTRAINTS k
              ON k.OWNER = cc.OWNER AND k.CONSTRAINT_NAME = cc.CONSTRAINT_NAME
            WHERE k.CONSTRAINT_TYPE = 'P' AND cc.OWNER = c.OWNER AND cc.TABLE_NAME = c.TABLE_NAME
              AND cc.COLUMN_NAME = c.COLUMN_NAME AND ROWNUM = 1) AS PK,
          (SELECT m.COMMENTS FROM ALL_COL_COMMENTS m WHERE m.OWNER = c.OWNER
              AND m.TABLE_NAME = c.TABLE_NAME AND m.COLUMN_NAME = c.COLUMN_NAME) AS CMT
        FROM ALL_TAB_COLUMNS c
        WHERE c.OWNER = NVL(:o, USER) AND c.TABLE_NAME = :t
        ORDER BY c.COLUMN_ID"""
        cur = conn.cursor()
        cur.execute(sql, o=(self.cfg["schema"] or "").strip().upper() or None, t=ident(table, "테이블"))
        out = []
        for name, dtype, dlen, nullable, pk, cmt in cur.fetchall():
            d = (dtype or "").upper()
            if "CHAR" in d:
                typ = "TEXT"
            elif d in ("NUMBER", "FLOAT") or d.startswith("BINARY"):
                typ = "NUM"
            elif d == "DATE" or d.startswith("TIMESTAMP"):
                typ = "DATE"
            else:
                typ = "TEXT"
            c = mk(True, name, (cmt or name)[:30], typ, req=(nullable == "N"), key=(pk == "Y"),
                   maxlen=int(dlen or 0) if typ == "TEXT" else 0,
                   note=f"{d}({dlen}){' PK' if pk == 'Y' else ''}{' NOT NULL' if nullable == 'N' else ''}")
            # 자동 추천: 등록일/등록자/회사코드는 입력칸에서 빼고 기본값으로
            if typ == "DATE" and RE_AUTO_DATE.search(name):
                c.update(use=False, default="SYSDATE")
            elif RE_AUTO_USER.search(name):
                c.update(use=False, default="{COPY}" if typ == "NUM" else "{USER}")
            elif name in COPY_NAMES:
                c.update(use=False, default="{COPY}")
            out.append(c)

        names = [c["col"] for c in out]
        code = next((n for n in CODE_NAMES if n in names), "")
        item = next((n for n in ITEM_NAMES if n in names), "") if p == "item" else ""
        for c in out:
            if c["col"] in (code, item) and c["col"]:
                c.update(use=True, req=True, key=True, default="")
        return out, code

    def find_mold_tables(self, conn):
        sql = """
        SELECT t.TABLE_NAME, t.NUM_ROWS,
          (SELECT c.COMMENTS FROM ALL_TAB_COMMENTS c
            WHERE c.OWNER = t.OWNER AND c.TABLE_NAME = t.TABLE_NAME) AS CMT,
          (SELECT LISTAGG(cc.COLUMN_NAME, ',') WITHIN GROUP (ORDER BY cc.POSITION)
             FROM ALL_CONSTRAINTS k JOIN ALL_CONS_COLUMNS cc
               ON cc.OWNER = k.OWNER AND cc.CONSTRAINT_NAME = k.CONSTRAINT_NAME
            WHERE k.OWNER = t.OWNER AND k.TABLE_NAME = t.TABLE_NAME AND k.CONSTRAINT_TYPE = 'P') AS PK,
          (SELECT COUNT(*) FROM ALL_CONSTRAINTS f JOIN ALL_CONSTRAINTS p
               ON p.OWNER = f.R_OWNER AND p.CONSTRAINT_NAME = f.R_CONSTRAINT_NAME
            WHERE f.CONSTRAINT_TYPE = 'R' AND p.OWNER = t.OWNER AND p.TABLE_NAME = t.TABLE_NAME) AS REFCNT,
          (SELECT COUNT(*) FROM ALL_TAB_COLUMNS c
            WHERE c.OWNER = t.OWNER AND c.TABLE_NAME = t.TABLE_NAME) AS COLCNT
        FROM ALL_TABLES t
        WHERE t.OWNER = NVL(:o, USER)
          AND (t.TABLE_NAME LIKE '%MOLD%' OR t.TABLE_NAME LIKE '%DIE%'
               OR EXISTS (SELECT 1 FROM ALL_TAB_COLUMNS c
                           WHERE c.OWNER = t.OWNER AND c.TABLE_NAME = t.TABLE_NAME
                             AND c.COLUMN_NAME IN ('MOLD_CODE', 'MOLD_CD', 'MOLD_NO')))"""
        cur = conn.cursor()
        cur.execute(sql, o=(self.cfg["schema"] or "").strip().upper() or None)
        out = []
        for name, nrows, cmt, pk, refcnt, colcnt in cur.fetchall():
            n = name.upper()
            pkcols = [x for x in (pk or "").split(",") if x]
            score = 0
            if len(pkcols) == 1 and pkcols[0] in CODE_NAMES:
                score += 50
            if "MOLD" in n:
                score += 10
                if any(w in n for w in ("MST", "MASTER", "INFO", "BASE")):
                    score += 30
            if any(w in n for w in ("HIST", "DTL", "LOG", "DETAIL", "BAK", "TMP", "TEMP")):
                score -= 25
            score += min(30, 5 * int(refcnt or 0))
            exact = False
            if nrows is None or nrows < EXACT_COUNT_LIMIT:
                try:
                    c2 = conn.cursor()
                    c2.execute(f"SELECT COUNT(*) FROM {self.qualify(name)}")
                    nrows, exact = c2.fetchone()[0], True
                except Exception:
                    pass
            if nrows is not None and 100 <= nrows <= 3000:
                score += 10
            out.append({"table": name, "rows": nrows, "exact": exact, "cmt": cmt or "", "pk": pk or "",
                        "refcnt": int(refcnt or 0), "colcnt": int(colcnt or 0), "score": score})
        out.sort(key=lambda x: (-x["score"], x["table"]))
        return out

    def preview(self, conn, table, n=PREVIEW_ROWS):
        cur = conn.cursor()
        cur.execute(f"SELECT * FROM {self.qualify(table)} WHERE ROWNUM <= {int(n)}")
        hdr = [d[0] for d in cur.description]
        return hdr, [[cell_text(v) for v in r] for r in cur.fetchall()]

    def ref_values(self, conn, ref):
        tb, col = self.ref_table(ref)
        cur = conn.cursor()
        cur.execute(f"SELECT DISTINCT {col} FROM {tb} WHERE {col} IS NOT NULL ORDER BY 1")
        return [r[0] for r in cur.fetchmany(50000)]

    # ---------- 기준 금형 값 복사 ({COPY}) ----------
    def table_meta(self, conn, p):
        """실제 테이블 컬럼 정보 (NOT NULL / DB 기본값)"""
        t = (self.part(p).get("table") or "").strip().upper()
        owner = (self.cfg["schema"] or "").strip().upper() or None
        if "." in t:
            owner, t = t.rsplit(".", 1)
        k = (owner, t)
        if k not in self._meta:
            cur = conn.cursor()
            cur.execute("SELECT COLUMN_NAME, DATA_TYPE, NULLABLE, DATA_DEFAULT FROM ALL_TAB_COLUMNS "
                        "WHERE OWNER = NVL(:o, USER) AND TABLE_NAME = :t ORDER BY COLUMN_ID", o=owner, t=t)
            meta = {}
            for name, dtype, nullable, dflt in cur.fetchall():
                d = (dtype or "").upper()
                typ = "TEXT" if "CHAR" in d else ("NUM" if d in ("NUMBER", "FLOAT") else
                                                  ("DATE" if d == "DATE" or d.startswith("TIMESTAMP") else "OTHER"))
                raw = str(dflt or "").strip()
                dv = raw.upper()
                meta[name] = {"type": typ, "notnull": nullable == "N", "hasdef": dv not in ("", "NULL"),
                              "default": "" if dv in ("", "NULL") else raw}
            self._meta[k] = meta
        return self._meta[k]

    def auto_cols(self, conn, p):
        """NOT NULL 인데 입력도 기본값도 없는 컬럼 -> 자동으로 채울 방법"""
        meta = self.table_meta(conn, p)
        ins = {c["col"] for c in self.insertable(p)}
        reg_user = (self.cfg.get("reg_user") or "").strip()
        out = {}
        for col, m in meta.items():
            if not m["notnull"] or m["hasdef"] or col in ins:
                continue
            if m["type"] == "DATE" and RE_AUTO_DATE.search(col):
                out[col] = "SYSDATE"
            elif m["type"] == "TEXT" and RE_AUTO_USER.search(col) and reg_user:
                out[col] = "{USER}"
            else:
                out[col] = "{COPY}"
        return out

    def copy_row(self, conn, p, extra=()):
        need = [c["col"] for c in self.insertable(p) if (c["default"] or "").strip().upper() == "{COPY}"]
        need += [c for c in extra if c not in need]
        if not need:
            return {}
        ref = (self.cfg.get("ref_mold") or "").strip()
        if not ref:
            raise ValueError(f"{PART_NAME[p]} 테이블의 아래 컬럼을 채울 값이 없습니다:\n  {', '.join(need)}\n\n"
                             f"[환경설정]의 '기준 금형코드'에 작업지시가 잘 되는 금형코드를 넣고 [저장]하면\n"
                             f"그 금형의 값을 복사해서 채웁니다.")
        col = self.code_col() if p == "mold" else self.link_col()
        cur = conn.cursor()
        cur.execute(f"SELECT {', '.join(ident(c, '컬럼') for c in need)} FROM {self.table(p)} "
                    f"WHERE {ident(col, '컬럼')} = :v AND ROWNUM = 1", v=ref)
        row = cur.fetchone()
        if not row:
            raise ValueError(f"기준 금형 '{ref}'의 {PART_NAME[p]} 정보가 {self.table(p)}에 없습니다.\n"
                             f"{PART_NAME[p]}이(가) 등록되어 있는 금형을 기준 금형으로 지정하세요.")
        return dict(zip(need, row))

    def prepare(self, conn):
        """자동 채움 컬럼 계산 + 기준 금형 값 읽기"""
        self.last_auto, copies = {}, {}
        for p in PARTS:
            if not self.enabled(p):
                continue
            auto = self.auto_cols(conn, p)
            self.last_auto[p] = auto
            copies[p] = self.copy_row(conn, p, [c for c, how in auto.items() if how == "{COPY}"])
        return copies

    # ---------- 검사 ----------
    def check_part(self, cur, p, rows, ref_cache, meta=None):
        meta = meta or {}
        kb = int(self.cfg.get("kor_bytes") or 3)
        used, keys, tbl = self.used(p), self.keys(p), self.table(p)
        id_cols = [ident(c["col"], "컬럼") for c in keys] or [ident(used[0]["col"], "컬럼")]
        seen, out = {}, []
        for rno, row in rows:
            r = new_result(rno)
            errs, warns, bad, vals = r["errs"], r["warns"], r["bad"], r["vals"]
            for c in used:
                col = c["col"]
                try:
                    v = conv(c, row.get(col))
                except ValueError as e:
                    errs.append(f"{c['name']} {e}")
                    bad.add(col)
                    continue
                vals[col] = v
                if v is None:
                    m = meta.get(col, {})
                    if (c["req"] or (m.get("notnull") and not m.get("hasdef"))) and not c["default"]:
                        errs.append(f"{c['name']} 비어있음")
                        bad.add(col)
                    continue
                if c["type"] == "TEXT" and c["maxlen"]:
                    n = blen(v, kb)
                    if n > int(c["maxlen"]):
                        errs.append(f"{c['name']} 너무 김({n}/{c['maxlen']}byte)")
                        bad.add(col)
                        continue
                if c["ref"]:
                    k = (c["ref"], str(v))
                    if k not in ref_cache:
                        rtb, rc = self.ref_table(c["ref"])
                        cur.execute(f"SELECT COUNT(*) FROM {rtb} WHERE {rc} = :v", v=str(v))
                        ref_cache[k] = cur.fetchone()[0] > 0
                    if not ref_cache[k]:
                        errs.append(f"{c['name']} MES에 없음({v})")
                        bad.add(col)
                        continue
                if c["sim"] and not c["key"]:
                    cur.execute(f"SELECT {', '.join(id_cols)} FROM {tbl} "
                                f"WHERE {ident(col, '컬럼')} = :v AND ROWNUM <= 6", v=v)
                    found = ["/".join(show(x) for x in rr) for rr in cur.fetchall()]
                    if found:
                        warns.append(f"같은 {c['name']} 기존: {', '.join(found[:5])}{' 외' if len(found) > 5 else ''}")
                        r["warncols"].add(col)

            if keys and all(vals.get(k["col"]) is not None for k in keys):
                ks = tuple(str(vals[k["col"]]).upper() for k in keys)
                if ks in seen:
                    errs.append(f"시트 안 중복({seen[ks]}행과 같음)")
                    bad.update(k["col"] for k in keys)
                else:
                    seen[ks] = rno
                    if not errs:
                        where = " AND ".join(f"{ident(k['col'], '컬럼')} = :{i + 1}" for i, k in enumerate(keys))
                        cur.execute(f"SELECT COUNT(*) FROM {tbl} WHERE {where}", [vals[k["col"]] for k in keys])
                        if cur.fetchone()[0] > 0:
                            errs.append(f"이미 MES에 등록됨")
                            bad.update(k["col"] for k in keys)
            out.append(r)
        return out

    def check(self, conn, mold_rows, item_rows):
        self.validate_map()
        copies = self.prepare(conn)             # 필수 컬럼 자동 채움 / 기준 금형 확인
        cur = conn.cursor()
        ref_cache = {}
        mres = self.check_part(cur, "mold", mold_rows, ref_cache, self.table_meta(conn, "mold"))
        ires = (self.check_part(cur, "item", item_rows, ref_cache, self.table_meta(conn, "item"))
                if self.enabled("item") else [])

        if self.enabled("item"):
            code, link = self.code_col(), self.link_col()
            new = {}
            for r in mres:
                v = r["vals"].get(code)
                if v is not None:
                    new.setdefault(str(v).upper(), r)
            linked, exist_cache = set(), {}
            for r in ires:
                v = r["vals"].get(link)
                if v is None:
                    continue
                k = str(v).upper()
                linked.add(k)
                if k in new:
                    if new[k]["errs"]:
                        r["errs"].append(f"금형 시트 {new[k]['rno']}행에 오류")
                        r["bad"].add(link)
                else:
                    if k not in exist_cache:
                        cur.execute(f"SELECT COUNT(*) FROM {self.table('mold')} "
                                    f"WHERE {ident(code, '컬럼')} = :v", v=v)
                        exist_cache[k] = cur.fetchone()[0] > 0
                    if not exist_cache[k]:
                        r["errs"].append(f"금형코드 {v}가 금형 시트와 MES 어디에도 없음")
                        r["bad"].add(link)
                    else:
                        if self.cfg.get("allow_item_to_existing"):
                            r["warns"].append(f"기존 금형 {v}에 품번 줄 추가 (기존 줄은 그대로)")
                        else:
                            r["errs"].append(f"{v}는 이미 있는 금형입니다. 신규 금형만 등록 가능 "
                                             f"(기존 금형에 품번 추가는 '기존 금형 품번추가 허용' 체크)")
                            r["bad"].add(link)
            for k, r in new.items():
                if k not in linked:
                    r["warns"].append("품번 연결 없음 (품번 시트에 추가해야 작업지시 가능)")
                    r["warncols"].add(code)
        # DB의 '중복 불가' 규칙(유니크 인덱스) 전부 검사 - 자동으로 채울 값까지 포함
        for p, res in (("mold", mres), ("item", ires)):
            if res:
                self.check_unique(conn, p, res, copies.get(p, {}), self.last_auto.get(p, {}))
        for r in mres + ires:
            finalize(r)
        return mres, ires

    def unique_sets(self, conn, p):
        """유니크 인덱스 -> {이름: [식...]}  (함수 인덱스 UPPER(..), NVL(..) 도 포함)"""
        t = (self.part(p).get("table") or "").strip().upper()
        owner = (self.cfg["schema"] or "").strip().upper() or None
        if "." in t:
            owner, t = t.rsplit(".", 1)
        k = (owner, t)
        if k not in self._uniq:
            cur = conn.cursor()
            cur.execute("SELECT c.INDEX_NAME, c.COLUMN_POSITION, c.COLUMN_NAME FROM ALL_INDEXES i "
                        "JOIN ALL_IND_COLUMNS c ON c.INDEX_OWNER = i.OWNER AND c.INDEX_NAME = i.INDEX_NAME "
                        "WHERE i.TABLE_OWNER = NVL(:o, USER) AND i.TABLE_NAME = :t AND i.UNIQUENESS = 'UNIQUE' "
                        "ORDER BY c.INDEX_NAME, c.COLUMN_POSITION", o=owner, t=t)
            cols = cur.fetchall()
            exprs = {}
            try:
                cur.execute("SELECT INDEX_NAME, COLUMN_POSITION, COLUMN_EXPRESSION FROM ALL_IND_EXPRESSIONS "
                            "WHERE TABLE_OWNER = NVL(:o, USER) AND TABLE_NAME = :t", o=owner, t=t)
                for n, pos, ex in cur.fetchall():
                    exprs[(n, pos)] = str(ex or "").strip()
            except Exception:
                pass
            sets = {}
            for n, pos, col in cols:
                sets.setdefault(n, []).append(exprs.get((n, pos)) or f'"{col}"')
            self._uniq[k] = sets
        return self._uniq[k]

    @staticmethod
    def _refs(expr, meta):
        names = set(re.findall(r'"([A-Z0-9_$#]+)"', expr)) | set(re.findall(r"\b([A-Z_][A-Z0-9_$#]*)\b", expr))
        return [n for n in names if n in meta]

    @staticmethod
    def _alias(expr, alias, meta):
        def sub(m):
            name = m.group(1) or m.group(2)
            return f'{alias}."{name}"' if name in meta else m.group(0)
        return re.sub(r'"([A-Z0-9_$#]+)"|\b([A-Z_][A-Z0-9_$#]*)\b', sub, expr)

    def full_values(self, p, vals, copyrow, auto):
        """실제로 INSERT될 값 (입력 + 기본값 + 자동 채움). SYSDATE는 제외"""
        reg = (self.cfg.get("reg_user") or "").strip()
        out = {}
        for c in self.insertable(p):
            v = vals.get(c["col"]) if c["use"] else None
            if v is None:
                d = (c["default"] or "").strip()
                du = d.upper()
                if not d or du == "SYSDATE":
                    continue
                v = copyrow.get(c["col"]) if du == "{COPY}" else (reg if d == "{USER}" else conv(c, d))
            out[c["col"]] = v
        for col, how in (auto or {}).items():
            if how != "SYSDATE":
                out[col] = reg if how == "{USER}" else copyrow.get(col)
        return out

    def check_unique(self, conn, p, results, copyrow, auto):
        """새 줄이 DB의 '중복 불가' 규칙에 걸리는지 실제 DB 식 그대로 확인"""
        sets = self.unique_sets(conn, p)
        if not sets:
            return
        meta = self.table_meta(conn, p)
        by = {c["col"]: c for c in self.part(p)["columns"]}
        idcol = self.code_col() if p == "mold" else self.link_col()
        tbl = self.table(p)
        cur = conn.cursor()
        seen = {n: {} for n in sets}
        for r in results:
            if r["errs"]:
                continue
            full = self.full_values(p, r["vals"], copyrow, auto)
            for name, exprs in sets.items():
                refs = sorted({c for e in exprs for c in self._refs(e, meta)})
                if not refs:
                    continue
                sel, binds, shown = [], [], []
                for c in refs:
                    v = full.get(c)
                    dflt = meta[c].get("default", "")
                    if v is not None:
                        binds.append(v)
                        sel.append(f':{len(binds)} "{c}"')
                        shown.append(f"{by[c]['name'] if c in by else c}={show(v)}")
                    elif dflt and "NEXTVAL" not in dflt.upper():
                        sel.append(f'{dflt} "{c}"')
                        shown.append(f"{c}=기본값({dflt})")
                    else:
                        sel.append(f'NULL "{c}"')
                        shown.append(f"{c}=빈값")
                desc = " + ".join(shown)
                key = tuple(str(full.get(c)).upper().strip() for c in refs)
                if key in seen[name]:
                    r["errs"].append(f"시트 안 중복({seen[name][key]}행) [{name}: {desc}]")
                    r["bad"].update(c for c in refs if c in r["vals"])
                    continue
                seen[name][key] = r["rno"]
                cond = " AND ".join(f"DECODE({self._alias(e, 'T', meta)}, {self._alias(e, 'N', meta)}, 1, 0) = 1"
                                    for e in exprs)
                sql = (f'SELECT T."{idcol}" FROM {tbl} T, (SELECT {", ".join(sel)} FROM DUAL) N '
                       f"WHERE {cond} AND ROWNUM <= 3")
                try:
                    cur.execute(sql, binds)
                    found = [show(x[0]).strip() for x in cur.fetchall()]
                except Exception as e:
                    r["warns"].append(f"중복규칙 {name} 확인 못 함({str(e)[:60]})")
                    continue
                if found:
                    rule = " + ".join(exprs).replace('"', "")
                    who = f"이미 금형 {', '.join(found)}에 같은 값이 있음" if p == "item" else f"기존 {', '.join(found)}와 겹침"
                    r["errs"].append(f"중복 불가 [{name}: {rule}] {desc} → {who}")
                    r["bad"].update(c for c in refs if c in r["vals"])

    def build_insert(self, p, vals, copyrow, auto=None):
        reg_user = (self.cfg.get("reg_user") or "").strip()
        cols, ph, binds, log = [], [], [], []
        for c in self.insertable(p):
            v = vals.get(c["col"]) if c["use"] else None
            if v is None:
                d = (c["default"] or "").strip()
                if not d:
                    continue
                du = d.upper()
                if du == "SYSDATE":
                    cols.append(ident(c["col"], "컬럼"))
                    ph.append("SYSDATE")
                    log.append(f"{c['name']}=SYSDATE")
                    continue
                if du == "{COPY}":
                    v = copyrow.get(c["col"])
                    if v is None:
                        continue
                elif d == "{USER}":
                    v = reg_user
                else:
                    v = conv(c, d)
            binds.append(v)
            cols.append(ident(c["col"], "컬럼"))
            ph.append(f":{len(binds)}")
            log.append(f"{c['name']}={show(v)}")
        for col, how in (auto or {}).items():
            cols.append(ident(col, "컬럼"))
            if how == "SYSDATE":
                ph.append("SYSDATE")
                log.append(f"{col}=SYSDATE(자동)")
                continue
            v = reg_user if how == "{USER}" else copyrow.get(col)
            binds.append(v)
            ph.append(f":{len(binds)}")
            log.append(f"{col}={show(v)}(자동)")
        if not cols:
            raise ValueError("넣을 값이 없습니다.")
        tbl = self.table(p)
        keys = [c for c in self.keys(p) if vals.get(c["col"]) is not None]
        if keys:
            # 같은 키가 이미 있으면 아무것도 넣지 않음 (기존 줄은 절대 건드리지 않음)
            conds = []
            for k in keys:
                binds.append(vals[k["col"]])
                conds.append(f"{ident(k['col'], '컬럼')} = :{len(binds)}")
            sql = (f"INSERT INTO {tbl} ({', '.join(cols)}) SELECT {', '.join(ph)} FROM DUAL "
                   f"WHERE NOT EXISTS (SELECT 1 FROM {tbl} WHERE {' AND '.join(conds)})")
        else:
            sql = f"INSERT INTO {tbl} ({', '.join(cols)}) VALUES ({', '.join(ph)})"
        return sql, binds, "; ".join(log)

    def triggers(self, conn):
        """등록 대상 테이블에 걸린 자동 동작(트리거)"""
        out = {}
        cur = conn.cursor()
        for p in PARTS:
            if not self.enabled(p):
                continue
            t = (self.part(p).get("table") or "").strip().upper()
            owner = (self.cfg["schema"] or "").strip().upper() or None
            if "." in t:
                owner, t = t.rsplit(".", 1)
            try:
                cur.execute("SELECT TRIGGER_NAME, TRIGGERING_EVENT FROM ALL_TRIGGERS "
                            "WHERE TABLE_OWNER = NVL(:o, USER) AND TABLE_NAME = :t AND STATUS = 'ENABLED'", o=owner, t=t)
                out[p] = [f"{a} ({b.strip()})" for a, b in cur.fetchall()]
            except Exception:
                out[p] = []
        return out

    def usage_tables(self, conn, cols, skip=()):
        """col 이름을 가진 테이블 목록 (통계상 행수 작은 것부터)"""
        owner = (self.cfg["schema"] or "").strip().upper() or None
        cur = conn.cursor()
        cur.execute("SELECT c.TABLE_NAME, c.COLUMN_NAME, NVL(t.NUM_ROWS, 0) FROM ALL_TAB_COLUMNS c JOIN ALL_TABLES t "
                    "ON t.OWNER = c.OWNER AND t.TABLE_NAME = c.TABLE_NAME "
                    f"WHERE c.OWNER = NVL(:o, USER) AND c.COLUMN_NAME IN ({', '.join(repr(c) for c in cols)}) "
                    "ORDER BY 3, 1", o=owner)
        return [(t, c, n) for t, c, n in cur.fetchall() if t not in skip]

    def usage_of_code(self, conn, code, progress=None, cancel=None, timeout_ms=15000):
        """conn 대신 {"c": conn} 을 주면, 시간초과로 연결이 끊겨도 새로 접속해서 계속함"""
        holder = conn if isinstance(conn, dict) else {"c": conn}
        conn = holder["c"]
        """금형코드를 쓰는 다른 테이블 (작업지시·샷수 이력 등). 금형/품번 테이블은 제외.
        표마다 1건만 찾으면 멈추고, 오래 걸리는 표는 timeout_ms 뒤 '확인 못함'으로 넘어감"""
        skip = {(self.part(p).get("table") or "").strip().upper().split(".")[-1] for p in PARTS}
        tables = self.usage_tables(conn, ("MOLD_CODE", "MOLD_CD", "MOLD_NO"), skip)
        used, failed = [], []
        old_to = getattr(conn, "call_timeout", 0)
        for k, (tname, col, nrows) in enumerate(tables, 1):
            if cancel and cancel():
                failed.append("(취소됨)")
                break
            if progress:
                progress(f"{code}: {k}/{len(tables)}  {tname} ({nrows:,}행) 확인 중...")
            try:
                conn.call_timeout = timeout_ms
                c2 = conn.cursor()
                c2.execute(f"SELECT COUNT(*) FROM {self.qualify(tname)} "
                           f"WHERE {ident(col, '컬럼')} = :v AND ROWNUM = 1", v=code)
                if c2.fetchone()[0]:
                    used.append((tname, col, 1))
            except Exception as e:
                failed.append(f"{tname}({'시간초과' if 'timeout' in str(e).lower() or 'DPI-1067' in str(e) or 'DPY-4024' in str(e) else '오류'})")
                try:
                    conn.ping()
                except Exception:
                    try:
                        conn.close()
                    except Exception:
                        pass
                    conn = holder["c"] = self.connect()      # 끊겼으면 다시 접속해서 계속
                    old_to = 0
            finally:
                try:
                    conn.call_timeout = old_to
                except Exception:
                    pass
        return used, failed

    def usage_of_value(self, conn, col, value, skip=(), progress=None, cancel=None, timeout_ms=15000):
        """col 값(예: 품번, 작업지시번호)을 쓰는 다른 테이블. 표마다 1건만 찾으면 멈추고, 오래 걸리면 넘어감"""
        tables = self.usage_tables(conn, (col,), set(skip))
        used, failed = [], []
        for k, (tname, cname, nrows) in enumerate(tables, 1):
            if cancel and cancel():
                failed.append("(취소됨)")
                break
            if progress:
                progress(f"{value}: {k}/{len(tables)}  {tname} ({nrows:,}행) 확인 중...")
            try:
                conn.call_timeout = timeout_ms
                c2 = conn.cursor()
                c2.execute(f"SELECT COUNT(*) FROM {self.qualify(tname)} "
                           f"WHERE {ident(cname, '컬럼')} = :v AND ROWNUM = 1", v=value)
                if c2.fetchone()[0]:
                    used.append((tname, cname, 1))
            except Exception:
                failed.append(tname)
                conn.ping()
            finally:
                try:
                    conn.call_timeout = 0
                except Exception:
                    pass
        return used, failed

    def col_meta(self, conn, tbl):
        owner, t = tbl.split(".") if "." in tbl else ((self.cfg["schema"] or "").strip().upper() or None, tbl)
        cur = conn.cursor()
        cur.execute("SELECT COLUMN_NAME, NULLABLE, DATA_DEFAULT FROM ALL_TAB_COLUMNS "
                    "WHERE OWNER = NVL(:o, USER) AND TABLE_NAME = :t", o=owner, t=t)
        out = {}
        for n, nl, d in cur.fetchall():
            dv = str(d or "").strip().upper()
            out[n] = {"notnull": nl == "N", "hasdef": dv not in ("", "NULL")}
        return out

    def insert_many(self, conn, steps):
        """steps: [(tbl, vals, auto)] -> 전부 INSERT 후 한 번에 커밋 (하나라도 실패하면 전체 취소)"""
        cur = conn.cursor()
        try:
            for k, (tbl, vals, auto) in enumerate(steps, 1):
                cols, ph, binds = [], [], []
                for c, v in vals.items():
                    binds.append(v)
                    cols.append(ident(c, "컬럼"))
                    ph.append(f":{len(binds)}")
                for c, how in auto.items():
                    cols.append(ident(c, "컬럼"))
                    if how == "SYSDATE":
                        ph.append("SYSDATE")
                    else:
                        binds.append(how)
                        ph.append(f":{len(binds)}")
                sql = f"INSERT INTO {tbl} ({', '.join(cols)}) VALUES ({', '.join(ph)})"
                if not sql.upper().startswith("INSERT INTO "):
                    raise RuntimeError("INSERT 이외 명령은 실행하지 않습니다.")
                try:
                    cur.execute(sql, binds)
                except Exception as ex:
                    raise RuntimeError(f"{k}번째 ({tbl.split('.')[-1]}) 등록 실패\n{ora_hint(ex)}")
            conn.commit()
        except Exception as ex:
            conn.rollback()
            raise RuntimeError(f"등록 중 오류 -> 전체 취소 (MES에 아무것도 안 들어감)\n\n{ex}")

    def insert_row(self, conn, tbl, vals, auto, derive=None):
        """새 줄 1개 INSERT (auto: {컬럼: 'SYSDATE' 또는 등록자})"""
        vals = dict(vals)
        if derive:
            try:
                vals.update(derive(conn, vals))
            except Exception:
                pass
        cols, ph, binds = [], [], []
        for c, v in vals.items():
            binds.append(v)
            cols.append(ident(c, "컬럼"))
            ph.append(f":{len(binds)}")
        for c, how in auto.items():
            cols.append(ident(c, "컬럼"))
            if how == "SYSDATE":
                ph.append("SYSDATE")
            else:
                binds.append(how)
                ph.append(f":{len(binds)}")
        cur = conn.cursor()
        try:
            cur.execute(f"INSERT INTO {tbl} ({', '.join(cols)}) VALUES ({', '.join(ph)})", binds)
            if cur.rowcount != 1:
                raise RuntimeError("등록되지 않았습니다.")
            conn.commit()
        except Exception as ex:
            conn.rollback()
            raise RuntimeError(f"등록 중 오류 -> 취소 (MES는 그대로)\n\n{ora_hint(ex)}")

    def delete_by_rowid(self, conn, tbl, targets):
        """targets: [(rowid, label, row)] -> 한 묶음으로 DELETE"""
        cur = conn.cursor()
        n = 0
        try:
            for rid, label, row in targets:
                cur.execute(f"DELETE FROM {tbl} WHERE ROWID = CHARTOROWID(:1)", [rid])
                if cur.rowcount != 1:
                    raise RuntimeError(f"{label}: 이미 없어진 줄입니다. 다시 조회하세요.")
                n += 1
            conn.commit()
        except Exception as ex:
            conn.rollback()
            raise RuntimeError(f"삭제 중 오류 -> 전체 취소 (MES는 그대로)\n\n{ora_hint(ex)}")
        return n

    def update_by_rowid(self, conn, tbl, edits, defs):
        """edits: [(rowid, {"changes": {col: (old, new)}, "label"})] -> 한 묶음으로 UPDATE"""
        reg = (self.cfg.get("reg_user") or "").strip()
        a_date = [c for c, d in defs.items() if d["type"] == "DATE" and d.get("editable") and RE_UPD_DATE.search(c)]
        a_user = [c for c, d in defs.items() if d["type"] == "TEXT" and d.get("editable") and RE_UPD_USER.search(c)] \
            if reg else []
        cur = conn.cursor()
        n = 0
        try:
            for rid, e in edits:
                sets, binds, where = [], [], []
                for col, (old, new) in e["changes"].items():
                    binds.append(new)
                    sets.append(f"{ident(col, '컬럼')} = :{len(binds)}")
                for c in a_date:
                    if c not in e["changes"]:
                        sets.append(f"{ident(c, '컬럼')} = SYSDATE")
                for c in a_user:
                    if c not in e["changes"]:
                        binds.append(reg)
                        sets.append(f"{ident(c, '컬럼')} = :{len(binds)}")
                binds.append(rid)
                where.append(f"ROWID = CHARTOROWID(:{len(binds)})")
                for col, (old, new) in e["changes"].items():
                    if old is None:
                        where.append(f"{ident(col, '컬럼')} IS NULL")
                    else:
                        binds.append(old)
                        where.append(f"{ident(col, '컬럼')} = :{len(binds)}")
                cur.execute(f"UPDATE {tbl} SET {', '.join(sets)} WHERE {' AND '.join(where)}", binds)
                if cur.rowcount != 1:
                    raise RuntimeError(f"{e['label']}: 그 사이 다른 곳에서 바뀌었거나 없어진 줄입니다. 다시 조회하세요.")
                n += 1
            conn.commit()
        except Exception as ex:
            conn.rollback()
            raise RuntimeError(f"수정 중 오류 -> 전체 취소 (MES는 그대로)\n\n{ora_hint(ex)}")
        return n

    def delete_rows(self, conn, targets):
        """targets: [{"p", "key": {col: 값}, "row": {col: 값}, "label"}] -> 한 묶음으로 DELETE"""
        cur = conn.cursor()
        done = []
        try:
            for t in targets:
                where, binds = [], []
                for col, v in t["key"].items():
                    if v is None:
                        where.append(f"{ident(col, '컬럼')} IS NULL")
                    else:
                        binds.append(v)
                        where.append(f"{ident(col, '컬럼')} = :{len(binds)}")
                cur.execute(f"DELETE FROM {self.table(t['p'])} WHERE {' AND '.join(where)}", binds)
                if cur.rowcount != 1:
                    raise RuntimeError(f"{t['label']}: 해당 줄이 {cur.rowcount}개라 삭제하지 않았습니다.")
                done.append(t)
            conn.commit()
        except Exception as ex:
            conn.rollback()
            raise RuntimeError(f"삭제 중 오류 -> 전체 취소 (MES는 그대로)\n\n{ora_hint(ex)}")
        return done

    def apply_updates(self, conn, edits):
        """edits: [{"p", "key": {col: 원래값}, "changes": {col: (원래값, 새값)}}] -> 한 묶음으로 UPDATE"""
        cur = conn.cursor()
        reg_user = (self.cfg.get("reg_user") or "").strip()
        ref_cache = {}
        # 1) 저장 전 검사 (존재확인 / 바뀐 키 중복)
        for e in edits:
            p = e["p"]
            by = {c["col"]: c for c in self.part(p)["columns"]}
            for col, (old, new) in e["changes"].items():
                c = by.get(col)
                if c and c["ref"] and new is not None:
                    k = (c["ref"], str(new))
                    if k not in ref_cache:
                        rtb, rc = self.ref_table(c["ref"])
                        cur.execute(f"SELECT COUNT(*) FROM {rtb} WHERE {rc} = :v", v=str(new))
                        ref_cache[k] = cur.fetchone()[0] > 0
                    if not ref_cache[k]:
                        raise ValueError(f"{e['label']} / {c['name']}: '{new}' 는 MES에 없는 값입니다.")
            if any(col in e["key"] for col in e["changes"]):
                newkey = {k: e["changes"][k][1] if k in e["changes"] else v for k, v in e["key"].items()}
                where = " AND ".join(f"{ident(k, '컬럼')} = :{i + 1}" for i, k in enumerate(newkey))
                cur.execute(f"SELECT COUNT(*) FROM {self.table(p)} WHERE {where}", list(newkey.values()))
                if cur.fetchone()[0] > 0:
                    raise ValueError(f"{e['label']}: 바꾼 값 {'/'.join(show(v) for v in newkey.values())} 은(는) 이미 있습니다.")
        # 2) UPDATE
        done = []
        try:
            for e in edits:
                p = e["p"]
                sets, binds, where = [], [], []
                for col, (old, new) in e["changes"].items():
                    binds.append(new)
                    sets.append(f"{ident(col, '컬럼')} = :{len(binds)}")
                for c in self.part(p)["columns"]:
                    n = c["col"]
                    if n in e["changes"]:
                        continue
                    if c["type"] == "DATE" and RE_UPD_DATE.search(n):
                        sets.append(f"{ident(n, '컬럼')} = SYSDATE")
                    elif c["type"] == "TEXT" and RE_UPD_USER.search(n) and reg_user:
                        binds.append(reg_user)
                        sets.append(f"{ident(n, '컬럼')} = :{len(binds)}")
                conds = dict(e["key"])
                conds.update({col: old for col, (old, new) in e["changes"].items()})   # 그 사이 바뀌었는지 확인
                for col, v in conds.items():
                    if v is None:
                        where.append(f"{ident(col, '컬럼')} IS NULL")
                    else:
                        binds.append(v)
                        where.append(f"{ident(col, '컬럼')} = :{len(binds)}")
                cur.execute(f"UPDATE {self.table(p)} SET {', '.join(sets)} WHERE {' AND '.join(where)}", binds)
                if cur.rowcount != 1:
                    raise RuntimeError(f"{e['label']}: 해당 줄이 {cur.rowcount}개입니다.\n"
                                       f"(그 사이 다른 사람이 바꿨거나, 같은 키가 여러 줄)")
                done.append(e)
            conn.commit()
        except Exception as ex:
            conn.rollback()
            raise RuntimeError(f"수정 중 오류 -> 전체 취소 (MES는 그대로)\n\n{ora_hint(ex)}")
        return done

    def verify(self, conn, done):
        """등록 직후 MES에서 다시 읽어서 실제로 들어갔는지 확인"""
        cur = conn.cursor()
        ok, miss = 0, []
        for p, r, _ in done:
            keys = [c for c in self.keys(p) if r["vals"].get(c["col"]) is not None]
            if not keys:
                ok += 1
                continue
            where = " AND ".join(f"{ident(k['col'], '컬럼')} = :{i + 1}" for i, k in enumerate(keys))
            cur.execute(f"SELECT COUNT(*) FROM {self.table(p)} WHERE {where}", [r["vals"][k["col"]] for k in keys])
            if cur.fetchone()[0] > 0:
                ok += 1
            else:
                miss.append(f"{PART_NAME[p]} " + "/".join(show(r["vals"][k["col"]]) for k in keys))
        return ok, miss

    def find_items(self, conn, text):
        """품번 테이블 직접 검색 (금형코드·품번 일부, 공백/대소문자 무시)"""
        like = f"%{text.strip().upper()}%"
        conds = [f"UPPER(TRIM({ident(self.link_col(), '컬럼')})) LIKE :q"]
        for c in self.keys("item"):
            if c["col"] != self.link_col():
                conds.append(f"UPPER(TRIM({ident(c['col'], '컬럼')})) LIKE :q")
        cur = conn.cursor()
        cur.execute(f"SELECT * FROM {self.table('item')} WHERE ({' OR '.join(conds)}) AND ROWNUM <= 500", q=like)
        rows = cur.fetchall()
        defs = self.desc_defs("item", cur.description)
        first = [self.link_col()] + [c["col"] for c in self.used("item")]
        cols, rows = self.reorder([d[0] for d in cur.description], rows, first)
        return cols, rows, defs

    def register(self, conn, mres, ires):
        copies = self.prepare(conn)
        cur = conn.cursor()
        done, where = [], ("mold", 0)
        todo = {"mold": mres, "item": ires}
        before = {}
        for p, res in todo.items():
            if res:
                cur.execute(f"SELECT COUNT(*) FROM {self.table(p)}")
                before[p] = cur.fetchone()[0]
        try:
            for p, res in todo.items():
                for r in res:
                    where = (p, r["rno"])
                    sql, binds, log = self.build_insert(p, r["vals"], copies.get(p, {}), self.last_auto.get(p))
                    if not sql.upper().startswith("INSERT INTO "):          # 신규등록은 INSERT만
                        raise RuntimeError("INSERT 이외 명령은 실행하지 않습니다.")
                    cur.execute(sql, binds)
                    if cur.rowcount != 1:
                        raise RuntimeError("같은 키가 방금 MES에 생겼습니다. 기존 줄은 건드리지 않고 취소합니다.")
                    done.append((p, r, log))
            for p, n0 in before.items():                                   # 기존 줄 수 그대로인지 확인
                cur.execute(f"SELECT COUNT(*) FROM {self.table(p)}")
                n1 = cur.fetchone()[0]
                added = sum(1 for d in done if d[0] == p)
                if n1 != n0 + added:
                    raise RuntimeError(f"{self.table(p)} 줄 수가 예상과 다릅니다 (전 {n0} + 신규 {added} ≠ 후 {n1}).\n"
                                       f"다른 작업과 겹쳤을 수 있어 안전하게 취소합니다. 잠시 후 다시 하세요.")
            conn.commit()
        except Exception as e:
            conn.rollback()
            extra = ""
            mm = re.search(r"\(([^.()]+)\.([^.()]+)\) violated", str(e))
            if "ORA-00001" in str(e) and mm:
                try:
                    rule = self.unique_sets(conn, where[0]).get(mm.group(2), [])
                    if rule:
                        extra = f"\n\n이 규칙은 [{' + '.join(rule).replace(chr(34), '')}] 조합이 같으면 막습니다."
                except Exception:
                    pass
            raise RuntimeError(f"{PART_NAME[where[0]]} 시트 {where[1]}행 등록 중 오류\n"
                               f"-> 전체 취소 (MES에 아무것도 안 들어감)\n\n{ora_hint(e)}{extra}")
        self.last_verify = self.verify(conn, done)
        return done


# ==============================================================
# 엑셀
# ==============================================================
class ExcelIO:
    def __init__(self, mes):
        self.mes = mes

    def _sheet(self, wb, p, first, prefill):
        used = self.mes.used(p)
        kb = int(self.mes.cfg.get("kor_bytes") or 3)
        ws = wb.active if first else wb.create_sheet()
        ws.title = SHEET[p]
        center = Alignment(horizontal="center", vertical="center")
        for j, c in enumerate(used, 1):
            ws.cell(1, j, c["col"]).font = Font(size=8, color="969696")
            h = ws.cell(2, j, c["name"] + (" *" if c["req"] and not c["default"] else ""))
            h.font = Font(bold=True)
            h.alignment = center
            h.fill = fill("F8CBAD" if c["key"] else ("BDD7EE" if c["req"] else "D9D9D9"))
            w = max(10, min(30, blen(c["name"], 2) + 4))
            if c["type"] == "DATE":
                w = max(w, 12)
            ws.column_dimensions[get_column_letter(j)].width = w
            fmt = "@" if c["type"] == "TEXT" else ("yyyy-mm-dd" if c["type"] == "DATE" else "General")
            for r in range(3, 3 + TEMPLATE_ROWS):
                ws.cell(r, j).number_format = fmt
        rc = len(used) + 1
        h = ws.cell(2, rc, "검사결과")
        h.font = Font(bold=True)
        h.alignment = center
        h.fill = fill("FFF2CC")
        ws.column_dimensions[get_column_letter(rc)].width = 60
        ws.freeze_panes = "A3"

        blank = {c["col"] for c in used if c["key"]} if p == "mold" else {self.mes.link_col()}
        for i, row in enumerate(prefill or []):
            for j, c in enumerate(used, 1):
                if c["col"] in blank:
                    continue
                v = row.get(c["col"])
                if v is not None:
                    ws.cell(3 + i, j, show(v) if c["type"] == "TEXT" else v)
            ws.cell(3 + i, rc, "복사됨 - 주황 칸(금형코드) 입력 후 검사하세요").fill = fill(C_INFO)
        return ws

    def make_template(self, path, prefill=None, conn=None):
        prefill = prefill or {}
        wb = openpyxl.Workbook()
        sheets = {}
        for i, p in enumerate(p for p in PARTS if self.mes.enabled(p)):
            sheets[p] = self._sheet(wb, p, i == 0, prefill.get(p))
        if conn is not None:
            wl, k = None, 0
            for p, ws in sheets.items():
                for j, c in enumerate(self.mes.used(p), 1):
                    if not c["ref"]:
                        continue
                    vals = self.mes.ref_values(conn, c["ref"])
                    if not vals:
                        continue
                    if wl is None:
                        wl = wb.create_sheet(LIST_SHEET)
                        wl.sheet_state = "hidden"
                    k += 1
                    L = get_column_letter(k)
                    wl.cell(1, k, f"{SHEET[p]}.{c['col']}")
                    for n, v in enumerate(vals, 2):
                        wl.cell(n, k, show(v)).number_format = "@"
                    dv = DataValidation(type="list", formula1=f"'{LIST_SHEET}'!${L}$2:${L}${len(vals) + 1}",
                                        allow_blank=True, showErrorMessage=True, errorStyle="warning",
                                        errorTitle=c["name"], error=f"MES에 없는 {c['name']} 입니다.")
                    ws.add_data_validation(dv)
                    CL = get_column_letter(j)
                    dv.add(f"{CL}3:{CL}{2 + TEMPLATE_ROWS}")
        wb.save(path)

    def _load_sheet(self, ws, p, path):
        used = self.mes.used(p)
        dbset = {c["col"] for c in used}
        name2col = {c["name"]: c["col"] for c in used}
        first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        row1 = ["" if x is None else str(x).strip() for x in first]
        if any(v.upper() in dbset for v in row1):
            colmap = {i: v.upper() for i, v in enumerate(row1) if v.upper() in dbset}
            start = 3
        else:
            colmap = {i: name2col[v.replace("*", "").strip()] for i, v in enumerate(row1)
                      if v.replace("*", "").strip() in name2col}
            start = 2
        missing = [c["name"] for c in used if c["col"] not in colmap.values()]
        if missing:
            raise ValueError(f"[{ws.title}] 시트에 없는 컬럼: " + ", ".join(missing) +
                             "\n매핑을 바꿨다면 [엑셀 양식 만들기]로 새 양식을 만드세요.")
        rows = []
        for rno, vals in enumerate(ws.iter_rows(min_row=start, values_only=True), start):
            d = {col: (vals[i] if i < len(vals) else None) for i, col in colmap.items()}
            if all(v is None or str(v).strip() == "" for v in d.values()):
                continue
            rows.append((rno, d))
        return rows, {"sheet": ws.title, "hdr": start - 1, "rescol": max(colmap) + 2, "colmap": colmap}

    def load(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        out = {}
        for p in PARTS:
            if not self.mes.enabled(p):
                out[p] = ([], None)
                continue
            if SHEET[p] in wb.sheetnames:
                ws = wb[SHEET[p]]
            elif p == "mold":
                ws = wb.worksheets[0]
            else:
                out[p] = ([], None)
                continue
            out[p] = self._load_sheet(ws, p, path)
        return out

    def write_results(self, path, parts, done_text=None):
        """parts: [(info, results), ...]"""
        wb = openpyxl.load_workbook(path)
        none = PatternFill(fill_type=None)
        for info, results in parts:
            if not info:
                continue
            ws = wb[info["sheet"]]
            col_of = {v: i + 1 for i, v in info["colmap"].items()}
            rc = info["rescol"]
            ws.cell(info["hdr"], rc, "검사결과").font = Font(bold=True)
            ws.column_dimensions[get_column_letter(rc)].width = 60
            for r in results:
                rno = r["rno"]
                for c in col_of.values():
                    ws.cell(rno, c).fill = none
                if done_text:
                    ws.cell(rno, rc, done_text).fill = fill(C_OK)
                    continue
                for col in r["warncols"]:
                    if col in col_of:
                        ws.cell(rno, col_of[col]).fill = fill(C_WARN)
                for col in r["bad"]:
                    if col in col_of:
                        ws.cell(rno, col_of[col]).fill = fill(C_ERR)
                ws.cell(rno, rc, r["msg"]).fill = fill({"오류": C_ERR, "경고": C_WARN}.get(r["status"], C_OK))
        try:
            wb.save(path)
            return path
        except PermissionError:
            base, ext = os.path.splitext(path)
            alt = f"{base}_결과_{dt.datetime.now():%H%M%S}{ext}"
            wb.save(alt)
            return alt


# ==============================================================
# 화면
# ==============================================================
# ==============================================================
# 품목 / 작업지시 탭 (표에서 바로 보고 고치기 - ROWID 기준)
# ==============================================================
def _date_arg(s, what):
    s = (s or "").strip().replace(".", "-").replace("/", "-")
    try:
        return dt.datetime.strptime(s, "%Y-%m-%d")
    except ValueError:
        raise ValueError(f"{what} 날짜 형식이 틀렸습니다 (예: 2026-09-30): {s}")


def item_where(g):
    f, conds, b = g.fvars, [], {}
    q = f["q"].get().strip().upper()
    if q:
        conds.append("(UPPER(T.ITEM_CODE) LIKE :q OR UPPER(T.ITEM_NAME) LIKE :q OR UPPER(T.ITEM_SPEC) LIKE :q)")
        b["q"] = f"%{q}%"
    return conds, b


def item_tag(g, row):
    i = g.cols.index("ITEM_CODE") if "ITEM_CODE" in g.cols else None
    if i is not None and re.match(r"^\d{4}-\d{2}-\d{2}", show(row[i]).strip()):
        return "err"                     # 품번 자리에 날짜가 들어간 오타
    return ""


def wo_where(g):
    f, conds, b = g.fvars, [], {}
    d1 = _date_arg(f["d1"].get(), "시작")
    d2 = _date_arg(f["d2"].get(), "끝")
    conds.append("T.WORK_ORDER_DATE >= :d1 AND T.WORK_ORDER_DATE < :d2")
    b["d1"], b["d2"] = d1, d2 + dt.timedelta(days=1)
    mc = f["mc"].get().strip().upper()
    if mc:
        conds.append("UPPER(T.MACHINE_CODE) LIKE :mc")
        b["mc"] = f"%{mc}%"
    it = f["item"].get().strip().upper()
    if it:
        conds.append("(UPPER(T.ITEM_CODE) LIKE :it OR UPPER(T.ITEM_NAME) LIKE :it)")
        b["it"] = f"%{it}%"
    if f["nomold"].get():
        conds.append("NVL(TRIM(T.MOLD_CODE), '*') = '*'")
    return conds, b


def wo_extra(g):
    item_tbl = g.app.mes.qualify(g.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")
    return [(f"(SELECT MAX(X.MOLD_CODE) FROM {item_tbl} X WHERE X.ORGANIZATION_ID = T.ORGANIZATION_ID "
             f"AND X.ITEM_CODE = T.ITEM_CODE)", "LINK_MOLD__")]


def wo_prepare(g):
    """같은 날·설비·품번 여러 건 표시용"""
    g.dups = {}
    need = ("WORK_ORDER_DATE", "MACHINE_CODE", "ITEM_CODE")
    if all(c in g.cols for c in need):
        idx = [g.cols.index(c) for c in need]
        for r in g.rows:
            k = tuple(show(r[i]) for i in idx)
            g.dups[k] = g.dups.get(k, 0) + 1


def wo_tag(g, row):
    c = g.cols
    mold = show(row[c.index("MOLD_CODE")]).strip() if "MOLD_CODE" in c else ""
    link = show(row[c.index("LINK_MOLD__")]).strip() if "LINK_MOLD__" in c else ""
    if mold in ("", "*"):
        return "err"
    if link and mold != link:
        return "warn"
    need = ("WORK_ORDER_DATE", "MACHINE_CODE", "ITEM_CODE")
    if all(x in c for x in need) and g.dups.get(tuple(show(row[c.index(x)]) for x in need), 0) > 1:
        return "dup"
    return ""


def wo_summary(g):
    tags = [wo_tag(g, r) for r in g.rows]
    parts = []
    if tags.count("err"):
        parts.append(f"금형 없음(*) {tags.count('err')}건")
    if tags.count("warn"):
        parts.append(f"금형이 현재 연결과 다름 {tags.count('warn')}건")
    if tags.count("dup"):
        parts.append(f"같은 날·설비·품번 여러 건 {tags.count('dup')}건")
    return "   ※ " + " / ".join(parts) if parts else ""


def item_summary(g):
    n = sum(1 for r in g.rows if item_tag(g, r))
    return f"   ※ 품번이 날짜 모양인 오타 의심 {n}건 (빨강)" if n else ""


def item_new_defaults(g, base, meta):
    v = dict(base)
    v["ITEM_CODE"] = None
    v["ITEM_ID"] = None
    return v


def item_new_check(g, conn, tbl, vals):
    code = vals.get("ITEM_CODE")
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM {tbl} WHERE ITEM_CODE = :c AND NVL(ORGANIZATION_ID, -1) = NVL(:o, -1)",
                c=code, o=vals.get("ORGANIZATION_ID"))
    if cur.fetchone()[0]:
        return f"품번 {code}는 이미 있습니다."
    return None


def item_delete_block(g, row):
    return None


WO_RESET = ("ACTUAL_QTY", "PASS_QTY", "BAD_QTY", "INPUT_QTY", "OUT_RST_QTY", "IN_RST_QTY", "OUT_PREVIOUS_QTY",
            "IN_PREVIOUS_QTY", "AVG_TACT_TIME", "LASTEST_TACT_TIME", "EXTRA_IN_QTY", "EXTRA_OUT_QTY", "INTERFACE_DATE")


def wo_next_no(conn, tbl, date):
    """다음 작업지시번호 제안: 같은 날 번호의 앞부분 + 전체 최대 일련번호 + 1"""
    cur = conn.cursor()
    cur.execute(f"SELECT WORK_ORDER_NO, WORK_ORDER_DATE FROM {tbl} WHERE WORK_ORDER_DATE >= :d",
                d=dt.datetime.now() - dt.timedelta(days=60))
    seq, prefix = 0, None
    for no, d in cur.fetchall():
        m = WO_RE.match(str(no or "").strip())
        if not m:
            continue
        seq = max(seq, int(m.group(2)))
        if d and date and d.date() == date.date():
            prefix = "WO" + m.group(1)
    if prefix is None and date:
        rule, _note = wo_learn_rule(conn, tbl)          # 과거 번호에서 배운 규칙 (10월 이후 월 표기 포함)
        prefix = "WO" + rule(date)
    return f"{prefix}P{seq + 1}" if seq else None


def wo_new_defaults(g, base, meta):
    v = dict(base)
    today = dt.datetime.combine(dt.date.today(), dt.time())
    v["WORK_ORDER_DATE"] = today
    for c in WO_RESET:
        if c in v:
            v[c] = None
    tbl = g.table()
    v["WORK_ORDER_NO"] = g.app.run_db(lambda conn: wo_next_no(conn, tbl, today))
    return v


def wo_new_check(g, conn, tbl, vals):
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM {tbl} WHERE WORK_ORDER_NO = :n", n=vals.get("WORK_ORDER_NO"))
    if cur.fetchone()[0]:
        return f"작업지시번호 {vals.get('WORK_ORDER_NO')}는 이미 있습니다. 번호를 바꾸세요."
    if vals.get("ITEM_CODE"):
        item_tbl = g.app.mes.qualify(g.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")
        cur.execute(f"SELECT MAX(MOLD_CODE) FROM {item_tbl} WHERE ITEM_CODE = :i", i=vals["ITEM_CODE"])
        if not cur.fetchone()[0]:
            if not messagebox.askyesno(APP_TITLE, f"품번 {vals['ITEM_CODE']}에 연결된 금형이 없어 금형이 '*'로 들어갑니다.\n"
                                                  f"그래도 등록할까요?"):
                return "취소했습니다."
    return None


def wo_derive(conn, vals):
    """설비명·품명·규격은 MES 함수로 다시 채움"""
    out = {}
    cur = conn.cursor()
    org, mc, it = vals.get("ORGANIZATION_ID"), vals.get("MACHINE_CODE"), vals.get("ITEM_CODE")
    if mc is not None:
        try:
            cur.execute("SELECT F_GET_MACHINE_NAME(:1, :2) FROM DUAL", [mc, org])
            r = cur.fetchone()[0]
            if r:
                out["MACHINE_NAME"] = r
        except Exception:
            pass
    if it is not None:
        try:
            cur.execute("SELECT F_GET_ITEM_NAME(:o, :i), F_GET_ITEM_SPEC(:o, :i) FROM DUAL", o=org, i=it)
            a, b = cur.fetchone()
            if a:
                out["ITEM_NAME"] = a
            if b:
                out["ITEM_SPEC"] = b
        except Exception:
            pass
    return out


def wo_delete_block(g, row):
    for c in ("ACTUAL_QTY", "PASS_QTY", "BAD_QTY", "INPUT_QTY"):
        if c in g.cols:
            v = row[g.cols.index(c)]
            if v not in (None, 0, 0.0):
                return f"실적이 있습니다 ({c}={show(v)}). 실적 있는 작업지시는 지우지 않습니다."
    return None


def bom_where(g):
    f, conds, b = g.fvars, [], {}
    q = f["q"].get().strip().upper()
    if q:
        conds.append("(UPPER(T.ITEM_CODE) LIKE :q OR UPPER(T.CHILD_ITEM_CODE) LIKE :q)")
        b["q"] = f"%{q}%"
    if f["zero"].get():
        conds.append("NVL(T.UNIT_PER_QTY, 0) = 0")
    return conds, b


def bom_extra(g):
    mst = g.app.mes.qualify(g.app.cfg.get("item_master_table") or "ICOM_ITEM_MASTER")
    return [(f"(SELECT MAX(M.ITEM_NAME) FROM {mst} M WHERE M.ORGANIZATION_ID = T.ORGANIZATION_ID "
             f"AND M.ITEM_CODE = T.ITEM_CODE)", "ITEM_NAME__")]


def bom_tag(g, row):
    i = g.cols.index("UNIT_PER_QTY") if "UNIT_PER_QTY" in g.cols else None
    return "warn" if i is not None and not row[i] else ""


def bom_summary(g):
    n = sum(1 for r in g.rows if bom_tag(g, r))
    return f"   ※ 단위중량 0/없음 {n}건 (노랑) - 이 조합으로 생산하면 소재 사용량이 0으로 계산됩니다" if n else ""


def bom_new_check(g, conn, tbl, vals):
    keys = [c for c in ("ORGANIZATION_ID", "ITEM_CODE", "CHILD_ITEM_CODE", "WORKSTAGE_CODE", "CHILD_WORKSTAGE_CODE")
            if c in g.cols]
    conds, b = [], {}
    for k in keys:
        if vals.get(k) is None:
            conds.append(f"{k} IS NULL")
        else:
            conds.append(f"{k} = :{k.lower()}")
            b[k.lower()] = vals[k]
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM {tbl} WHERE {' AND '.join(conds)}", b)
    if cur.fetchone()[0]:
        return f"제품 {vals.get('ITEM_CODE')} + 원소재 {vals.get('CHILD_ITEM_CODE')} 조합이 이미 있습니다. 기존 줄의 단위중량을 고치세요."
    return None


def bom_empty(g):
    """단위중량 조회 0건일 때: 품목 마스터·금형 연결·최근 생산 여부를 알려주고 새로 등록 안내"""
    q = g.fvars["q"].get().strip().upper()
    if not q:
        return
    link = g.app.mes.qualify(g.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")
    mst = g.app.mes.qualify(g.app.cfg.get("item_master_table") or "ICOM_ITEM_MASTER")
    dtl = g.app.mes.qualify("IMTL_RAW_SF_USE_CASE_DTL")

    def work(conn):
        cur = conn.cursor()
        out = {}
        for k, sql in (("mst", f"SELECT COUNT(*) FROM {mst} WHERE UPPER(ITEM_CODE) LIKE :q"),
                       ("link", f"SELECT COUNT(*) FROM {link} WHERE UPPER(ITEM_CODE) LIKE :q"),
                       ("prod", f"SELECT COUNT(*) FROM {dtl} WHERE UPPER(ITEM_CODE) LIKE :q AND ACTUCAL_DATE >= TRUNC(SYSDATE) - 60")):
            try:
                cur.execute(sql, q=f"%{q}%")
                out[k] = cur.fetchone()[0]
            except Exception:
                out[k] = None
        return out
    try:
        r = g.app.run_db(work)
    except Exception:
        r = {}
    yn = lambda v, a, b: "?" if v is None else (a if v else b)  # noqa: E731
    msg = (f"'{q}' 가 들어간 제품의 단위중량(피치당 소재중량) 줄이 MES에 없습니다.\n\n"
           f"  · 금형-품번 연결: {yn(r.get('link'), '있음', '없음')}\n"
           f"  · 품목 마스터: {yn(r.get('mst'), '있음', '없음')}\n"
           f"  · 최근 60일 생산 기록: {yn(r.get('prod'), str(r.get('prod')) + '건 (소재 사용량이 0으로 쌓이는 중)', '없음')}\n\n"
           f"넣으려면: [수정 모드] 체크 → [새로 등록] → 제품 품번, 원소재(코일) 코드, 피치당 소재중량(g) 입력.\n"
           f"(비슷한 제품을 먼저 조회해서 그 줄을 선택하고 [새로 등록]하면 나머지 칸이 복사됩니다)")
    messagebox.showinfo(APP_TITLE, msg)


def bom_recent_zero(g):
    """최근 60일 단위중량 0으로 생산된 제품+원소재 조합"""
    dtl = g.app.mes.qualify("IMTL_RAW_SF_USE_CASE_DTL")

    def work(conn):
        cur = conn.cursor()
        cur.execute(f"SELECT ITEM_CODE, RAW_ITEM_CODE, COUNT(*), SUM(SHOT_QTY), MAX(ACTUCAL_DATE) FROM {dtl} "
                    f"WHERE ACTUCAL_DATE >= TRUNC(SYSDATE) - 60 AND NVL(UNIT_PER_QTY, 0) = 0 "
                    f"GROUP BY ITEM_CODE, RAW_ITEM_CODE ORDER BY MAX(ACTUCAL_DATE) DESC")
        return cur.fetchall()
    try:
        rows = g.app.run_db(work)
    except Exception as e:
        g.app.err(e)
        return
    if not rows:
        messagebox.showinfo(APP_TITLE, "최근 60일 동안 단위중량 0으로 생산된 기록이 없습니다.")
        return
    w = tk.Toplevel(g.app.root)
    w.title("최근 60일 - 단위중량 0으로 생산된 조합")
    w.geometry("820x480")
    ttk.Label(w, text="이 조합들은 단위중량이 없어 소재 사용량이 0으로 쌓였습니다. [새로 등록]으로 단위중량을 넣으세요.\n"
                      "(넣은 뒤부터 계산됩니다. 이미 지난 기록은 바뀌지 않습니다)",
              foreground="#C00000", justify="left").pack(fill="x", padx=8, pady=8)
    fr, tv = g.app.make_tree(w, ["제품 품번", "원소재 코드", "기록 수", "샷 합계", "마지막 생산일"], [180, 220, 80, 100, 150])
    fr.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    for r in rows:
        tv.insert("", "end", values=[cell_text(v) for v in r])


ERP_MSG = "ERP 연동이 매일 이 칸을 덮어씁니다. 여기서 고쳐도 다음날 ERP 값으로 돌아갈 수 있습니다."

GRID_SPECS = {
    "bom": {
        "title": "단위중량", "table_cfg": "bom_table", "order": "T.ITEM_CODE, T.CHILD_ITEM_CODE",
        "first": ["ITEM_CODE", "ITEM_NAME__", "CHILD_ITEM_CODE", "UNIT_PER_QTY", "CAVITY_QTY",
                  "WORKSTAGE_CODE", "CHILD_WORKSTAGE_CODE"],
        "names": {"ITEM_CODE": "제품 품번", "ITEM_NAME__": "품명(참고)", "CHILD_ITEM_CODE": "원소재(코일) 코드",
                  "UNIT_PER_QTY": "피치당 소재중량(g)", "CAVITY_QTY": "캐비티",
                  "WORKSTAGE_CODE": "공정", "CHILD_WORKSTAGE_CODE": "원소재 공정"},
        "label_cols": ["ITEM_CODE", "CHILD_ITEM_CODE"],
        "locked": {"ORGANIZATION_ID": "회사코드는 바꿀 수 없습니다.",
                   "ITEM_CODE": "제품 품번은 바꿀 수 없습니다. 잘못된 줄은 삭제 후 [새로 등록]하세요.",
                   "CHILD_ITEM_CODE": "원소재 코드는 바꿀 수 없습니다. 잘못된 줄은 삭제 후 [새로 등록]하세요.",
                   "ENTER_DATE": "등록일은 바꾸지 않습니다.", "ENTER_BY": "등록자는 바꾸지 않습니다."},
        "sensitive": {"CAVITY_QTY": "캐비티는 샷수 × 캐비티로 생산수량·소재 사용량 계산에 쓰입니다."},
        "filters": [("q", "제품/원소재", "entry", "", 22), ("zero", ("모두", "단위중량 0/없음만"), "choice", False, 0)],
        "auto_fetch": True,
        "where": bom_where, "extra": bom_extra, "tag": bom_tag, "summary": bom_summary,
        "match_cols": ["ITEM_CODE", "CHILD_ITEM_CODE"],
        "new_check": bom_new_check, "new_locked": {"ITEM_NAME__": "-"},
        "tools": [("최근 60일 단위중량 0 생산 점검", bom_recent_zero)],
        "empty": bom_empty,
        "new_hint": "비슷한 줄을 선택하고 누르면 복사해서 시작합니다. 제품 품번·원소재 코드·단위중량을 확인하세요.",
        "hint": "피치당 소재중량(g) = 두께×폭×피치×7.85÷1000 (스크랩 포함, 제품 순수무게 아님).  캐비티 2 이상이면 ÷캐비티 값.  노랑 = 0/없음",
    },
    "itemmst": {
        "title": "품목", "table_cfg": "item_master_table", "order": "T.ITEM_CODE",
        "first": ["ITEM_CODE", "ITEM_NAME", "ITEM_SPEC", "ITEM_UNIT", "BUYER_PARTNO", "MODEL_CODE", "PACKING_QTY"],
        "names": {"ITEM_CODE": "품번", "ITEM_NAME": "품명", "ITEM_SPEC": "규격",
                  "ITEM_WEIGHT": "중량(미사용)", "GUIDE_WEIGHT": "게이트중량(미사용)"},
        "label_cols": ["ITEM_CODE"],
        "locked": {"ORGANIZATION_ID": "회사코드는 바꿀 수 없습니다.",
                   "ITEM_ID": "ERP 품목 번호라 바꿀 수 없습니다.",
                   "ITEM_CODE": "품번은 작업지시·재고 등 모든 곳에 연결되어 있어 바꿀 수 없습니다.",
                   "ENTER_DATE": "등록일은 바꾸지 않습니다.", "ENTER_BY": "등록자는 바꾸지 않습니다."},
        "sensitive": {c: ERP_MSG for c in ("PACKING_QTY", "BUYER_PARTNO", "MODEL_CODE", "ATTRIBUTE01",
                                            "ATTRIBUTE02", "ITEM_LVL1", "ITEM_LVL2", "ITEM_LVL3")},
        "filters": [("q", "품번/품명/규격", "entry", "", 22)],
        "where": item_where, "tag": item_tag, "summary": item_summary,
        "match_col": "ITEM_CODE",
        "new_defaults": item_new_defaults, "new_check": item_new_check, "delete_block": item_delete_block,
        "usage_col": "ITEM_CODE", "usage_block": True,
        "new_hint": "품번은 새로 입력하세요. 다른 칸은 선택한 품목에서 복사됩니다.",
        "hint": "품목 기본정보.  빨강 = 품번이 날짜 모양(오타 의심).  단위중량은 [단위중량 관리] 탭에서 관리합니다",
    },
    "wo": {
        "title": "작업지시", "table_cfg": "wo_table",
        "order": "T.WORK_ORDER_DATE, T.MACHINE_CODE, T.WORK_ORDER_NO",
        "first": ["WORK_ORDER_DATE", "MACHINE_CODE", "MACHINE_NAME", "WORK_ORDER_NO", "ITEM_CODE", "ITEM_NAME",
                  "MOLD_CODE", "LINK_MOLD__", "PLAN_QTY", "ACTUAL_QTY", "PASS_QTY", "BAD_QTY",
                  "WORK_ORDER_STATUS", "PLAN_STATUS", "WORK_SHIFT", "START_DATE", "END_DATE", "ST_VALUE"],
        "names": {"WORK_ORDER_DATE": "지시일", "MACHINE_CODE": "설비", "MACHINE_NAME": "설비명",
                  "WORK_ORDER_NO": "작업지시번호", "ITEM_CODE": "품번", "ITEM_NAME": "품명",
                  "MOLD_CODE": "금형", "LINK_MOLD__": "현재 연결 금형(참고)", "PLAN_QTY": "계획수량",
                  "ACTUAL_QTY": "실적수량", "PASS_QTY": "양품", "BAD_QTY": "불량",
                  "WORK_ORDER_STATUS": "상태", "PLAN_STATUS": "계획상태", "WORK_SHIFT": "근무조",
                  "START_DATE": "시작", "END_DATE": "종료", "ST_VALUE": "ST(초)"},
        "label_cols": ["WORK_ORDER_NO", "WORK_ORDER_DATE", "MACHINE_CODE"],
        "locked": {"ORGANIZATION_ID": "회사코드는 바꿀 수 없습니다.",
                   "WORK_ORDER_NO": "작업지시번호는 실적·자재출고와 연결되어 있어 바꿀 수 없습니다.",
                   "WIP_ENTITY_ID": "ERP 작업번호라 바꿀 수 없습니다.",
                   "MOLD_CODE": "금형은 품번-금형 연결로 자동으로 정해집니다 (MES 트리거).\n"
                                "금형을 바꾸려면 [금형·품번 관리] 탭에서 품번 연결을 고치세요.",
                   "MACHINE_NAME": "설비명은 설비를 바꾸면 자동으로 바뀝니다.",
                   "ENTER_DATE": "등록일은 바꾸지 않습니다.", "ENTER_BY": "등록자는 바꾸지 않습니다."},
        "sensitive": {
            "ACTUAL_QTY": "실적수량은 현장 실적(생산현황판)과 연결된 값입니다.",
            "PASS_QTY": "양품수량은 현장 실적과 연결된 값입니다.",
            "BAD_QTY": "불량수량은 현장 실적과 연결된 값입니다.",
            "INPUT_QTY": "투입수량은 현장 실적과 연결된 값입니다.",
            "WORK_ORDER_STATUS": "상태를 바꾸면 현장 화면에서 작업지시가 보이거나 사라질 수 있습니다.",
            "PLAN_STATUS": "상태를 바꾸면 현장 화면에서 작업지시가 보이거나 사라질 수 있습니다.",
            "MACHINE_CODE": "설비를 바꾸면 다른 설비로 작업지시가 옮겨갑니다 (설비명은 자동).",
            "ITEM_CODE": "품번을 바꾸면 금형이 새 품번의 연결 금형으로 자동으로 바뀝니다.",
            "WORK_ORDER_DATE": "지시일을 바꾸면 일자별 계획·실적 집계가 달라집니다.",
        },
        "filters": [("d1", "지시일", "entry", "", 11), ("d2", "~", "entry", "", 11),
                    ("mc", "설비", "entry", "", 8), ("item", "품번/품명", "entry", "", 14),
                    ("nomold", "금형 없음(*)만", "check", False, 0)],
        "new_defaults": wo_new_defaults, "new_check": wo_new_check, "derive": wo_derive,
        "new_locked": {"MOLD_CODE": "품번 연결 금형", "LINK_MOLD__": "-"},
        "delete_block": wo_delete_block, "usage_col": "WORK_ORDER_NO", "usage_block": False,
        "new_hint": "선택한 작업지시를 복사해서 오늘 날짜·다음 번호로 채웠습니다 (번호는 제안값 - 확인하세요).\n"
                    "금형은 품번 연결로, 설비명은 설비코드로 자동으로 들어갑니다.",
        "where": wo_where, "extra": wo_extra, "prepare": wo_prepare, "tag": wo_tag, "summary": wo_summary,
        "hint": "빨강 = 금형 없음(*)   노랑 = 금형이 지금 품번 연결과 다름   하늘 = 같은 날·설비·품번 여러 건   "
                "|   줄 더블클릭 = 고치기 창   (표에서 칸을 직접 고치려면 [수정 모드])",
        "actions": [("최근 작업지시 불러오기", lambda g: wo_open_recent(g)),
                    ("여러 건 일괄등록", lambda g: WoBatch(g)),
                    ("선택 작업지시 고치기", lambda g: wo_open_edit(g))],
        "hide_new": True,
        "dbl": lambda g, i: wo_open_edit(g, i),
    },
}


class GridTab:
    def __init__(self, app, parent, key):
        self.app, self.key, self.spec = app, key, GRID_SPECS[key]
        self.cols, self.rows, self.rids, self.defs, self.dups = [], [], [], {}, {}
        self.edits, self._ed, self.fetched, self.comments = {}, None, None, {}
        sp = self.spec
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(6, 2))
        self.fvars = {}
        for fk, label, kind, default, width in sp["filters"]:
            if kind == "check":
                v = tk.BooleanVar(value=default)
                ttk.Checkbutton(bar, text=label, variable=v).pack(side="left", padx=6)
            elif kind == "choice":                    # 둘 중 하나 고르기 (기본 = 첫 번째)
                v = tk.BooleanVar(value=default)
                for j, txt in enumerate(label):
                    ttk.Radiobutton(bar, text=txt, variable=v, value=bool(j),
                                    command=self.fetch).pack(side="left", padx=(8 if j == 0 else 2, 2))
            elif fk in ("d1", "d2"):
                ttk.Label(bar, text=label).pack(side="left", padx=(6, 2))
                de = DateEntry(bar)
                de.pack(side="left")
                v = de.var
            else:
                ttk.Label(bar, text=label).pack(side="left", padx=(6, 2))
                v = tk.StringVar(value=default)
                e = ttk.Entry(bar, textvariable=v, width=width)
                e.pack(side="left")
                e.bind("<Return>", lambda _e: self.fetch())
            self.fvars[fk] = v
        ttk.Button(bar, text="MES에서 조회", style="Big.TButton", command=self.fetch).pack(side="left", padx=8)
        ttk.Label(bar, text="표에서 찾기").pack(side="left", padx=(6, 2))
        self.var_q = tk.StringVar()
        ttk.Entry(bar, textvariable=self.var_q, width=16).pack(side="left")
        self.var_q.trace_add("write", lambda *_: self.show())
        if sp.get("match_col") or sp.get("match_cols"):
            ttk.Button(bar, text="엑셀로 일괄 수정", command=self.import_excel).pack(side="right", padx=3)
        ttk.Button(bar, text="엑셀로 저장", command=self.export_excel).pack(side="right", padx=3)
        for label, fn in sp.get("tools", []):
            ttk.Button(bar, text=label, command=lambda f=fn: f(self)).pack(side="right", padx=3)

        bar2 = ttk.Frame(parent)
        bar2.pack(fill="x", pady=(2, 2))
        self.var_mode = tk.BooleanVar(value=False)
        W = 16                                       # 버튼 너비 통일
        for label, fn in sp.get("actions", []):
            ttk.Button(bar2, text=label, width=W + 4, style="Big.TButton",
                       command=lambda f=fn: f(self)).pack(side="left", padx=3)
        if sp.get("actions"):
            ttk.Separator(bar2, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Checkbutton(bar2, text="수정 모드", variable=self.var_mode, command=self.on_mode).pack(side="left", padx=(4, 10))
        ttk.Label(parent, text=sp["hint"], foreground="#555").pack(fill="x", padx=6)   # 안내는 버튼 아래 한 줄
        ttk.Button(bar2, text="수정 취소", width=W - 4, style="Big.TButton", command=self.cancel).pack(side="right", padx=3)
        self.btn_save = ttk.Button(bar2, text="수정 저장 (0)", width=W - 2, style="Big.TButton", command=self.save)
        self.btn_save.pack(side="right", padx=3)
        ttk.Button(bar2, text="선택 줄 세로로 보기", width=W, style="Big.TButton",
                   command=lambda: self.detail()).pack(side="right", padx=3)
        ttk.Button(bar2, text="선택 줄 삭제", width=W - 4, style="Big.TButton", command=self.delete_rows).pack(side="right", padx=3)
        if not sp.get("hide_new"):
            ttk.Button(bar2, text="새로 등록", width=W - 4, style="Big.TButton", command=self.new_row).pack(side="right", padx=3)

        fr, self.tv = self.app.make_tree(parent, ["(조회 전)"])
        fr.pack(fill="both", expand=True, pady=4)
        self.tv.tag_configure("dup", background="#DDEBF7")
        self.tv.bind("<Double-1>", self.begin_edit)
        self.lbl = ttk.Label(parent, text="")
        self.lbl.pack(fill="x", padx=4, pady=(0, 4))

    # ---------- 조회 ----------
    def table(self):
        t = (self.app.cfg.get(self.spec["table_cfg"]) or "").strip()
        if not t:
            raise ValueError(f"[환경설정] 탭에 {self.spec['title']} 테이블명을 넣으세요.")
        return self.app.mes.qualify(t)

    def fetch(self, ask=True):
        self.end_edit(True)
        if ask and self.edits and not messagebox.askyesno(APP_TITLE, f"저장 안 한 수정 {len(self.edits)}줄이 있습니다.\n버리고 다시 조회할까요?"):
            return
        try:
            conds, binds = self.spec["where"](self)
            tbl = self.table()
            extra = self.spec.get("extra", lambda g: [])(self)
            sel = ", ".join(["ROWIDTOCHAR(T.ROWID) AS RID__", "T.*"] + [f"{e} AS {a}" for e, a in extra])
            sql = f"SELECT {sel} FROM {tbl} T"
            if conds:
                sql += " WHERE " + " AND ".join(conds)
            sql += f" ORDER BY {self.spec['order']}"
            sql = f"SELECT * FROM ({sql}) WHERE ROWNUM <= {int(self.app.cfg.get('limit') or 5000)}"
            owner, tname = tbl.split(".") if "." in tbl else (None, tbl)

            def work(conn):
                cur = conn.cursor()
                cur.execute(sql, binds)
                rows, desc = cur.fetchall(), cur.description
                if not self.comments:
                    try:
                        c2 = conn.cursor()
                        c2.execute("SELECT COLUMN_NAME, COMMENTS FROM ALL_COL_COMMENTS WHERE OWNER = NVL(:o, USER) "
                                   "AND TABLE_NAME = :t", o=owner, t=tname)
                        self.comments = {a: (b or "").strip() for a, b in c2.fetchall()}
                    except Exception:
                        pass
                return desc, rows
            desc, rows = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return
        names = [d[0] for d in desc]
        self.defs = self.make_defs(desc[1:])
        cols, data = Mes.reorder(names[1:], [list(r[1:]) for r in rows], self.spec["first"])
        self.cols, self.rows, self.rids = cols, data, [r[0] for r in rows]
        self.edits.clear()
        self.update_btn()
        self.fetched = dt.datetime.now()
        if self.spec.get("prepare"):
            self.spec["prepare"](self)
        self.show()
        if not rows and self.spec.get("empty"):
            self.spec["empty"](self)

    def make_defs(self, desc):
        defs = {}
        for d in desc:
            name, t, isize, null_ok = d[0], d[1], d[3], d[6]
            tn = str(getattr(t, "name", t)).upper()
            if any(w in tn for w in ("LOB", "RAW", "LONG", "ROWID")):
                typ = None
            elif "CHAR" in tn:
                typ = "TEXT"
            elif any(w in tn for w in ("NUMBER", "BINARY_", "FLOAT", "INTEGER", "DECIMAL")):
                typ = "NUM"
            elif "DATE" in tn or "TIMESTAMP" in tn:
                typ = "DATE"
            else:
                typ = None
            c = mk(True, name, self.head(name), typ or "TEXT", req=(null_ok is False),
                   maxlen=int(isize or 0) if typ == "TEXT" else 0)
            c["editable"] = typ is not None and not name.endswith("__")
            c["dbtype"] = tn.replace("DB_TYPE_", "")
            defs[name] = c
        return defs

    def head(self, col):
        nm = self.spec["names"].get(col) or self.comments.get(col) or ""
        return f"{nm} ({col})" if nm and not col.endswith("__") else (nm or col)

    def disp(self, i):
        vals = [cell_text(v) for v in self.rows[i]]
        e = self.edits.get(self.rids[i])
        if e:
            for col, (o, n) in e["changes"].items():
                vals[self.cols.index(col)] = cell_text(n)
        return vals, bool(e)

    def show(self):
        if self.fetched is None:
            return
        names = [self.head(c) for c in self.cols]
        disp = [self.disp(i) for i in range(len(self.rows))]
        self.app.reset_tree(self.tv, names, self.app.fit_widths(names, [v for v, _ in disp]))
        q = self.var_q.get().strip().lower()
        n = 0
        for i, (vals, edited) in enumerate(disp):
            if q and not any(q in v.lower() for v in vals):
                continue
            tag = "edit" if edited else self.spec["tag"](self, self.rows[i])
            self.tv.insert("", "end", iid=str(i), values=vals, tags=(tag,) if tag else ())
            n += 1
        lim = int(self.app.cfg.get("limit") or 5000)
        self.lbl.config(text=f"조회 {self.fetched:%H:%M:%S}  {len(self.rows):,}건"
                             f"{' (최대건수 도달 - 조건을 좁히세요)' if len(self.rows) >= lim else ''}"
                             f"{f'  / 표에서 찾기 {n:,}건' if q else ''}" + self.spec["summary"](self))

    def refresh_row(self, i):
        if self.tv.exists(str(i)):
            vals, edited = self.disp(i)
            tag = "edit" if edited else self.spec["tag"](self, self.rows[i])
            self.tv.item(str(i), values=vals, tags=(tag,) if tag else ())

    # ---------- 수정 ----------
    def on_mode(self):
        if self.var_mode.get():
            if not messagebox.askyesno(APP_TITLE, f"수정 모드를 켜면 MES의 {self.spec['title']} 내용을 바꿀 수 있습니다.\n켤까요?"):
                self.var_mode.set(False)
        else:
            self.end_edit(False)

    def label(self, i):
        return " / ".join(show(self.rows[i][self.cols.index(c)]).strip()
                          for c in self.spec["label_cols"] if c in self.cols)

    def block(self, col):
        if not self.var_mode.get():
            return "지금은 보기 전용입니다. 고치려면 [수정 모드]를 체크하세요."
        if col in self.spec["locked"]:
            return self.spec["locked"][col]
        c = self.defs.get(col)
        if not c or not c.get("editable"):
            return "이 칸은 여기서 고칠 수 없습니다." if col.endswith("__") else \
                f"{col} ({c.get('dbtype', '') if c else ''}) 형식은 고칠 수 없습니다."
        return None

    def confirm(self, col):
        msg = self.spec["sensitive"].get(col)
        return (not msg) or messagebox.askyesno(APP_TITLE, f"[{self.head(col)}]\n\n{msg}\n\n그래도 고칠까요?")

    def set_edit(self, i, col, text):
        c = self.defs[col]
        old = self.rows[i][self.cols.index(col)]
        try:
            new = conv(c, text)
            if new is None and c["req"]:
                raise ValueError("필수 칸이라 비울 수 없음")
            if c["type"] == "TEXT" and c["maxlen"] and new is not None:
                n = blen(new, int(self.app.cfg.get("kor_bytes") or 3))
                if n > int(c["maxlen"]):
                    raise ValueError(f"너무 김 ({n}/{c['maxlen']}byte)")
            if c["type"] == "NUM" and new is not None and isinstance(new, (int, float)) and new < 0:
                raise ValueError("음수는 넣을 수 없습니다")
        except ValueError as e:
            return f"{self.head(col)}: {e}\n수정하지 않았습니다."
        rid = self.rids[i]
        same = (cell_text(new) == cell_text(old) and (new is None) == (old is None)) or \
               (isinstance(old, str) and isinstance(new, str) and old.rstrip() == new)
        if same:
            if rid in self.edits:
                self.edits[rid]["changes"].pop(col, None)
                if not self.edits[rid]["changes"]:
                    del self.edits[rid]
        else:
            e = self.edits.setdefault(rid, {"changes": {}, "label": self.label(i), "idx": i})
            e["changes"][col] = (old, new)
        self.update_btn()
        return None

    def update_btn(self):
        self.btn_save.config(text=f"수정 저장 ({sum(len(e['changes']) for e in self.edits.values())})")

    def begin_edit(self, event):
        self.end_edit(True)
        tv = self.tv
        if tv.identify_region(event.x, event.y) != "cell":
            return
        iid, colid = tv.identify_row(event.y), tv.identify_column(event.x)
        if not iid:
            return
        if not self.var_mode.get():
            tv.selection_set(iid)
            if self.spec.get("dbl"):
                self.spec["dbl"](self, int(iid))
            else:
                self.detail()
            return
        ci = int(colid[1:]) - 1
        if ci < 0 or ci >= len(self.cols):
            return
        col = self.cols[ci]
        why = self.block(col)
        if why:
            messagebox.showwarning(APP_TITLE, why)
            return
        if not self.confirm(col):
            return
        bbox = tv.bbox(iid, colid)
        if not bbox:
            return
        x, y, w, h = bbox
        ent = ttk.Entry(tv)
        ent.place(x=x, y=y, width=max(w, 150), height=h)
        ent.insert(0, tv.item(iid, "values")[ci])
        ent.select_range(0, "end")
        ent.focus_set()
        self._ed = {"ent": ent, "i": int(iid), "col": col}
        ent.bind("<Return>", lambda _e: self.end_edit(True))
        ent.bind("<KP_Enter>", lambda _e: self.end_edit(True))
        ent.bind("<Escape>", lambda _e: self.end_edit(False))
        ent.bind("<FocusOut>", lambda _e: self.end_edit(True))
        tv.bind("<MouseWheel>", lambda _e: self.end_edit(True), add="+")

    def end_edit(self, commit):
        ed = self._ed
        if not ed:
            return
        self._ed = None
        text = ed["ent"].get()
        ed["ent"].destroy()
        if commit:
            msg = self.set_edit(ed["i"], ed["col"], text)
            if msg:
                messagebox.showerror(APP_TITLE, msg)
            self.refresh_row(ed["i"])

    def detail(self, i=None):
        self.end_edit(True)
        if i is None:
            sel = self.tv.selection()
            if len(sel) != 1:
                messagebox.showinfo(APP_TITLE, "한 줄을 선택하세요.")
                return
            i = int(sel[0])
        win = tk.Toplevel(self.app.root)
        win.title(f"{self.spec['title']} - {self.label(i)}")
        win.geometry("760x720")
        win.transient(self.app.root)
        ttk.Label(win, text=("값을 더블클릭해서 수정 → 주황 = 저장 대기. [수정 저장]을 눌러야 MES에 반영됩니다."
                             if self.var_mode.get() else "보기 전용 (고치려면 [수정 모드] 체크)"),
                  foreground="#555").pack(fill="x", padx=8, pady=(8, 2))
        qv = tk.StringVar()
        top = ttk.Frame(win)
        top.pack(fill="x", padx=8)
        ttk.Label(top, text="컬럼 찾기").pack(side="left")
        ttk.Entry(top, textvariable=qv, width=24).pack(side="left", padx=4)
        fr, dtv = self.app.make_tree(win, ["DB컬럼", "이름", "값", "형식"], [190, 170, 280, 90])
        fr.pack(fill="both", expand=True, padx=8, pady=6)

        def refill(*_):
            dtv.delete(*dtv.get_children())
            vals, _ = self.disp(i)
            e = self.edits.get(self.rids[i])
            q = qv.get().strip().lower()
            for j, col in enumerate(self.cols):
                nm = self.spec["names"].get(col) or self.comments.get(col, "")
                if q and q not in col.lower() and q not in nm.lower() and q not in vals[j].lower():
                    continue
                tag = ("edit",) if (e and col in e["changes"]) else (("off",) if self.block(col) else ())
                c = self.defs.get(col, {})
                dtv.insert("", "end", iid=str(j), tags=tag,
                           values=[col, nm, vals[j], c.get("dbtype", "") + (" 필수" if c.get("req") else "")])

        def edit(_e=None):
            s = dtv.selection()
            if not s:
                return
            j = int(s[0])
            col = self.cols[j]
            why = self.block(col)
            if why:
                messagebox.showwarning(APP_TITLE, why, parent=win)
                return
            if not self.confirm(col):
                return
            v = simpledialog.askstring(col, f"{self.head(col)}\n새 값을 입력하세요:",
                                       initialvalue=dtv.item(s[0], "values")[2], parent=win)
            if v is None:
                return
            msg = self.set_edit(i, col, v)
            if msg:
                messagebox.showerror(APP_TITLE, msg, parent=win)
            refill()
            self.refresh_row(i)
            dtv.selection_set(str(j))
            dtv.see(str(j))

        dtv.bind("<Double-1>", edit)
        qv.trace_add("write", refill)
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(bar, text="수정 저장", style="Big.TButton",
                   command=lambda: self.save() and win.destroy()).pack(side="left", padx=3)
        ttk.Button(bar, text="닫기", command=win.destroy).pack(side="right", padx=3)
        refill()

    def cancel(self):
        self.end_edit(False)
        if self.edits and messagebox.askyesno(APP_TITLE, f"저장 안 한 수정 {len(self.edits)}줄을 되돌릴까요?"):
            self.edits.clear()
            self.update_btn()
            self.show()

    def save(self):
        self.end_edit(True)
        if not self.edits:
            messagebox.showinfo(APP_TITLE, "수정한 내용이 없습니다.")
            return False
        lines = []
        for e in self.edits.values():
            for col, (o, n) in e["changes"].items():
                lines.append(f"{e['label']} · {self.head(col)}: '{cell_text(o)}' → '{cell_text(n)}'")
        more = f"\n... 외 {len(lines) - 25}건" if len(lines) > 25 else ""
        if not messagebox.askyesno("MES 수정", f"MES {self.spec['title']}에 바로 반영합니다.\n\n" +
                                   "\n".join(lines[:25]) + more + "\n\n계속할까요?"):
            return False
        edits = list(self.edits.items())
        try:
            tbl = self.table()
            n = self.app.run_db(lambda c: self.app.mes.update_by_rowid(c, tbl, edits, self.defs))
        except Exception as ex:
            self.app.err(ex)
            return False
        now = dt.datetime.now()
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", self.app.cfg.get("reg_user", ""),
                              f"{self.spec['title']}수정", tbl, e["label"],
                              "; ".join(f"{c}: {cell_text(o)} -> {cell_text(nv)}" for c, (o, nv) in e["changes"].items())]
                             for _, e in edits])
        self.app.load_log()
        messagebox.showinfo(APP_TITLE, f"{n}줄 수정 완료. [작업기록] 탭에 기록했습니다.")
        self.edits.clear()
        self.update_btn()
        self.fetch(ask=False)
        return True

    # ---------- 새로 등록 ----------
    def new_row(self):
        self.end_edit(True)
        if not self.var_mode.get():
            messagebox.showinfo(APP_TITLE, "새로 등록은 [수정 모드]를 체크해야 할 수 있습니다.")
            return
        if self.fetched is None:
            messagebox.showinfo(APP_TITLE, "먼저 [MES에서 조회]를 하세요. (표 구조를 읽어야 합니다)")
            return
        sel = self.tv.selection()
        base = dict(zip(self.cols, self.rows[int(sel[0])])) if len(sel) == 1 else {}
        if not base and not messagebox.askyesno(APP_TITLE, "선택한 줄이 없습니다. 빈 양식으로 시작할까요?\n"
                                                           "(비슷한 줄을 선택하고 누르면 그 값을 복사해서 시작합니다)"):
            return
        try:
            tbl = self.table()
            meta = self.app.run_db(lambda c: self.app.mes.col_meta(c, tbl))
            vals = self.spec["new_defaults"](self, base, meta) if self.spec.get("new_defaults") else dict(base)
        except Exception as e:
            self.app.err(e)
            return
        cols = [c for c in self.cols if not c.endswith("__")]
        reg = (self.app.cfg.get("reg_user") or "").strip()
        auto = {}
        for c in cols:
            d = self.defs.get(c, {})
            if d.get("type") == "DATE" and (RE_AUTO_DATE.search(c) or RE_UPD_DATE.search(c)):
                auto[c] = "SYSDATE"
            elif d.get("type") == "TEXT" and (RE_AUTO_USER.search(c) or RE_UPD_USER.search(c)) and reg:
                auto[c] = reg
        for c in auto:
            vals.pop(c, None)
        locked_new = set(self.spec.get("new_locked", {}))

        win = tk.Toplevel(self.app.root)
        win.title(f"{self.spec['title']} 새로 등록" + (" (선택 줄 복사)" if base else ""))
        win.geometry("820x760")
        win.transient(self.app.root)
        ttk.Label(win, text=self.spec.get("new_hint", "") + "\n값 칸을 더블클릭해서 입력.  빨강 = 반드시 입력.  "
                            "회색 = 자동(등록일·등록자 등).  주황 = 복사한 값과 다르게 바꾼 칸",
                  foreground="#555", justify="left").pack(fill="x", padx=8, pady=(8, 2))
        fr, dtv = self.app.make_tree(win, ["DB컬럼", "이름", "값", "형식"], [190, 170, 300, 110])
        fr.pack(fill="both", expand=True, padx=8, pady=6)
        orig = dict(vals)

        def need(c):
            m = meta.get(c, {})
            return m.get("notnull") and not m.get("hasdef") and c not in auto

        def refill():
            dtv.delete(*dtv.get_children())
            for j, c in enumerate(cols):
                d = self.defs.get(c, {})
                if c in auto:
                    v, tag = ("현재시각" if auto[c] == "SYSDATE" else auto[c]) + " (자동)", "off"
                elif c in locked_new:
                    v, tag = f"(자동: {self.spec['new_locked'][c]})", "off"
                else:
                    v = cell_text(vals.get(c))
                    tag = "err" if (need(c) and vals.get(c) is None) else \
                        ("edit" if base and cell_text(vals.get(c)) != cell_text(orig.get(c)) else "")
                typ = d.get("dbtype", "") + (" 필수" if need(c) else "")
                dtv.insert("", "end", iid=str(j), tags=(tag,) if tag else (),
                           values=[c, self.spec["names"].get(c) or self.comments.get(c, ""), v, typ])

        def edit(_e=None):
            s = dtv.selection()
            if not s:
                return
            c = cols[int(s[0])]
            if c in auto or c in locked_new:
                messagebox.showinfo(APP_TITLE, "자동으로 채워지는 칸입니다.", parent=win)
                return
            d = self.defs.get(c)
            if not d or not d.get("editable"):
                messagebox.showinfo(APP_TITLE, "이 형식은 입력할 수 없습니다.", parent=win)
                return
            v = simpledialog.askstring(c, f"{self.head(c)}\n값을 입력하세요 (비우면 빈값):",
                                       initialvalue=cell_text(vals.get(c)), parent=win)
            if v is None:
                return
            try:
                nv = conv(d, v)
                if d["type"] == "TEXT" and d["maxlen"] and nv is not None:
                    n = blen(nv, int(self.app.cfg.get("kor_bytes") or 3))
                    if n > int(d["maxlen"]):
                        raise ValueError(f"너무 김 ({n}/{d['maxlen']}byte)")
            except ValueError as e:
                messagebox.showerror(APP_TITLE, f"{c}: {e}", parent=win)
                return
            vals[c] = nv
            refill()
            dtv.selection_set(str(cols.index(c)))
            dtv.see(str(cols.index(c)))

        def register():
            miss = [c for c in cols if need(c) and c not in locked_new and vals.get(c) is None]
            if miss:
                messagebox.showerror(APP_TITLE, "반드시 입력할 칸이 비었습니다:\n  " + ", ".join(miss), parent=win)
                return
            chk = self.spec.get("new_check")
            label = " / ".join(show(vals.get(c)) for c in self.spec["label_cols"])
            try:
                if chk:
                    msg = self.app.run_db(lambda conn: chk(self, conn, tbl, vals))
                    if msg:
                        messagebox.showerror(APP_TITLE, msg, parent=win)
                        return
                if not messagebox.askyesno("새로 등록", f"MES {self.spec['title']}에 새 줄을 추가합니다.\n\n{label}\n\n"
                                                       "기존 줄은 바꾸지 않습니다. 계속할까요?", parent=win):
                    return
                ins = {c: v for c, v in vals.items() if v is not None and c in cols and c not in locked_new}
                self.app.run_db(lambda conn: self.app.mes.insert_row(conn, tbl, ins, auto,
                                                                     self.spec.get("derive")))
            except Exception as e:
                messagebox.showerror(APP_TITLE, ora_hint(e), parent=win)
                return
            now = dt.datetime.now()
            self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", reg, f"{self.spec['title']}등록", tbl, label,
                                  "; ".join(f"{c}={cell_text(v)}" for c, v in ins.items())]])
            self.app.load_log()
            win.destroy()
            messagebox.showinfo(APP_TITLE, f"새로 등록했습니다.\n{label}")
            self.fetch(ask=False)

        dtv.bind("<Double-1>", edit)
        dtv.bind("<Return>", edit)
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(bar, text="MES에 등록", style="Big.TButton", command=register).pack(side="left", padx=3)
        ttk.Button(bar, text="취소", command=win.destroy).pack(side="right", padx=3)
        refill()

    # ---------- 줄 삭제 ----------
    def delete_rows(self):
        self.end_edit(True)
        if not self.var_mode.get():
            messagebox.showinfo(APP_TITLE, "삭제는 [수정 모드]를 체크해야 할 수 있습니다.")
            return
        if self.edits:
            messagebox.showinfo(APP_TITLE, "저장 안 한 수정이 있습니다. 먼저 [수정 저장] 또는 [수정 취소]를 하세요.")
            return
        sel = [int(i) for i in self.tv.selection()]
        if not sel:
            messagebox.showinfo(APP_TITLE, "삭제할 줄을 선택하세요. (여러 줄: Ctrl/Shift+클릭)")
            return
        blocked = []
        for i in sel:
            why = self.spec.get("delete_block", lambda g, r: None)(self, self.rows[i])
            if why:
                blocked.append(f"{self.label(i)}: {why}")
        if blocked:
            messagebox.showwarning(APP_TITLE, "아래 줄은 삭제할 수 없습니다:\n\n" + "\n".join(blocked[:15]))
            return
        uc = self.spec.get("usage_col")
        tbl = self.table()
        used = []
        if uc and uc in self.cols:
            ci = self.cols.index(uc)
            vals = sorted({show(self.rows[i][ci]).strip() for i in sel})

            def work(conn, progress, cancelled):
                out = []
                for v in vals:
                    u, _ = self.app.mes.usage_of_value(conn, uc, v, {tbl.split(".")[-1]}, progress, cancelled)
                    out += [(v, t, n) for t, c, n in u]
                return out
            try:
                used = self.app.run_bg(work, "사용 이력 확인 중")
            except Exception as e:
                self.app.err(e)
                return
        lines = "\n".join(f"  {self.label(i)}" for i in sel[:20]) + (f"\n  ... 외 {len(sel) - 20}줄" if len(sel) > 20 else "")
        if used:
            ul = "\n".join(f"  {v} → {t}: 기록 있음" for v, t, n in used[:15])
            if self.spec.get("usage_block"):
                messagebox.showwarning(APP_TITLE, f"아래처럼 다른 곳에서 쓰이고 있어 삭제할 수 없습니다:\n\n{ul}")
                return
            typed = simpledialog.askstring("삭제 확인", f"삭제할 줄:\n{lines}\n\n※ 다른 테이블에 연결된 기록이 있습니다:\n{ul}\n\n"
                                                    f"그래도 지우려면 '삭제'라고 입력하세요:", parent=self.app.root)
            if (typed or "").strip() != "삭제":
                return
        elif not messagebox.askyesno("줄 삭제", f"MES {self.spec['title']}에서 아래 {len(sel)}줄을 삭제합니다.\n\n{lines}\n\n"
                                            f"삭제한 내용은 백업 파일에 저장됩니다. 계속할까요?", icon="warning"):
            return
        targets = [(self.rids[i], self.label(i), dict(zip(self.cols, self.rows[i]))) for i in sel]
        now = dt.datetime.now()
        with open(BACKUP_PATH, "a", encoding="utf-8") as f:
            for rid, label, row in targets:
                f.write(json.dumps({"time": f"{now:%Y-%m-%d %H:%M:%S}", "user": self.app.cfg.get("reg_user", ""),
                                    "table": tbl, "row": {k: show(v) for k, v in row.items() if not k.endswith("__")}},
                                   ensure_ascii=False) + "\n")
        try:
            n = self.app.run_db(lambda c: self.app.mes.delete_by_rowid(c, tbl, targets))
        except Exception as e:
            self.app.err(e)
            return
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", self.app.cfg.get("reg_user", ""), f"{self.spec['title']}삭제",
                              tbl, label, "; ".join(f"{k}={cell_text(v)}" for k, v in row.items()
                                                    if v is not None and not k.endswith("__"))]
                             for rid, label, row in targets])
        self.app.load_log()
        messagebox.showinfo(APP_TITLE, f"{n}줄 삭제 완료.\n백업: {BACKUP_PATH}")
        self.fetch(ask=False)

    # ---------- 엑셀 ----------
    def export_excel(self):
        if self.fetched is None:
            messagebox.showinfo(APP_TITLE, "먼저 [MES에서 조회]를 하세요.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialfile=f"{self.spec['title']}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = self.spec["title"]
        cols = [c for c in self.cols if not c.endswith("__")]
        ws.append(cols)
        ws.append([self.spec["names"].get(c) or self.comments.get(c, "") for c in cols])
        for c in ws[1]:
            c.font = Font(bold=True)
        for c in ws[2]:
            c.font = Font(size=9, color="808080")
        for i in range(len(self.rows)):
            vals = [self.rows[i][self.cols.index(c)] for c in cols]
            ws.append([cell_text(v, 2000) if not isinstance(v, (int, float, dt.datetime)) else v for v in vals])
        for j, c in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(j)].width = max(10, min(30, len(c) + 4))
        ws.freeze_panes = "B3"
        wb.save(path)
        open_file(path)
        if self.spec.get("match_col") or self.spec.get("match_cols"):
            messagebox.showinfo(APP_TITLE, "저장했습니다.\n\n이 파일에서 값을 고치고 저장한 뒤 [엑셀로 일괄 수정]으로 불러오면\n"
                                           "바뀐 칸만 주황색으로 표시됩니다. 확인 후 [수정 저장]을 누르세요.\n"
                                           "(1행 컬럼이름은 지우지 마세요. 필요 없는 열은 지워도 됩니다)")

    def import_excel(self):
        mcs = self.spec.get("match_cols") or [self.spec["match_col"]]
        if self.fetched is None or any(m not in self.cols for m in mcs):
            messagebox.showinfo(APP_TITLE, "먼저 [MES에서 조회]로 고칠 줄들을 불러오세요.")
            return
        if not self.var_mode.get():
            messagebox.showinfo(APP_TITLE, "[수정 모드]를 체크한 뒤 하세요.")
            return
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
        except Exception as e:
            self.app.err(e)
            return
        if not rows:
            return
        hdr = [str(h or "").strip().upper() for h in rows[0]]
        miss = [m for m in mcs if m not in hdr]
        if miss:
            messagebox.showerror(APP_TITLE, f"1행에 {', '.join(miss)} 컬럼이 있어야 합니다.")
            return
        kis = [hdr.index(m) for m in mcs]
        mcis = [self.cols.index(m) for m in mcs]
        idx = {}
        for i, r in enumerate(self.rows):
            idx.setdefault(tuple(show(r[c]).strip().upper() for c in mcis), []).append(i)
        mc = " + ".join(mcs)
        targets = [(j, h) for j, h in enumerate(hdr) if h in self.cols and h not in mcs]
        start = 2 if len(rows) > 1 and all((str(v or "").strip() in ("", self.spec["names"].get(hdr[j], ""),
                                                                      self.comments.get(hdr[j], "")))
                                           for j, v in enumerate(rows[1]) if j < len(hdr)) else 1
        nchg, notfound, errs, skipped = 0, [], [], set()
        for r in rows[start:]:
            if any(k >= len(r) or r[k] is None or str(r[k]).strip() == "" for k in kis):
                continue
            keyt = tuple(show(r[k]).strip().upper() for k in kis)
            key = "/".join(keyt)
            hits = idx.get(keyt)
            if not hits:
                notfound.append(key)
                continue
            for i in hits:
                for j, col in targets:
                    if j >= len(r):
                        continue
                    if self.block(col):
                        skipped.add(col)
                        continue
                    v = r[j]
                    old = self.rows[i][self.cols.index(col)]
                    newtxt = "" if v is None else (show(v) if not isinstance(v, str) else v)
                    if cell_text(conv_safe(self.defs[col], newtxt)) == cell_text(old):
                        continue
                    before = len(self.edits.get(self.rids[i], {}).get("changes", {}))
                    msg = self.set_edit(i, col, newtxt)
                    if msg:
                        errs.append(f"{key} {col}: {msg.splitlines()[0]}")
                    elif len(self.edits.get(self.rids[i], {}).get("changes", {})) > before:
                        nchg += 1
        self.show()
        msg = f"엑셀에서 {nchg}칸이 바뀐 것으로 표시했습니다 (주황).\n확인 후 [수정 저장]을 누르세요."
        if notfound:
            msg += f"\n\n조회 목록에 없는 {mc} {len(notfound)}개: {', '.join(notfound[:10])}" + \
                   (" ..." if len(notfound) > 10 else "") + "\n(조회 조건을 넓혀서 다시 조회 후 불러오세요)"
        if skipped:
            msg += f"\n\n고칠 수 없는 칸이라 건너뜀: {', '.join(sorted(skipped))}"
        if errs:
            msg += f"\n\n형식 오류 {len(errs)}건:\n" + "\n".join(errs[:10])
        messagebox.showinfo(APP_TITLE, msg)


def conv_safe(c, text):
    try:
        return conv(c, text)
    except ValueError:
        return text


# ==============================================================
# 구조 조사 / 전후 비교 (읽기 전용)
# ==============================================================
HASHABLE = ("CHAR", "VARCHAR2", "NCHAR", "NVARCHAR2", "NUMBER", "FLOAT", "DATE", "RAW",
            "BINARY_FLOAT", "BINARY_DOUBLE")
RE_TIME_COL = re.compile(r"(ENTER|CREATE|CREATED|CREATION|REG|REGIST|INSERT|INS|LAST_UPDATE|UPDATE|"
                         r"UPDATED|UPD|MODIFY|MOD)_?(DATE|DT|TIME|DTM)$")


def cell(v, limit=300):
    try:
        if v is None:
            return ""
        if hasattr(v, "read"):
            v = v.read(1, limit)
        if isinstance(v, (bytes, bytearray)):
            v = v[:limit // 2].hex().upper()
        if isinstance(v, dt.datetime):
            s = v.strftime("%Y-%m-%d %H:%M:%S")
        elif isinstance(v, float) and v.is_integer():
            s = str(int(v))
        else:
            s = str(v)
    except Exception as e:
        s = f"<읽기오류 {type(e).__name__}>"
    s = CTRL_RE.sub(" ", s)
    return s if len(s) <= limit else s[:limit] + "…"


class Inspector:
    """읽기 전용 조사 (SELECT만)"""
    def __init__(self, app):
        self.app = app
        self.pwd = None
        self.ready = False

    @property
    def cfg(self):
        return self.app.cfg

    @property
    def owner(self):
        return (self.cfg["schema"] or "").strip().upper() or None

    def connect(self):
        return self.app.mes.connect()

    @staticmethod
    def q(conn, sql, binds=None, many=None):
        """읽기 전용 안전장치: SELECT / WITH 외에는 실행하지 않음"""
        head = sql.lstrip().upper()
        if not (head.startswith("SELECT") or head.startswith("WITH")):
            raise RuntimeError("이 도구는 조회(SELECT)만 합니다.")
        cur = conn.cursor()
        cur.execute(sql, binds or {})
        rows = cur.fetchmany(many) if many else cur.fetchall()
        return [d[0] for d in cur.description], rows

    def tq(self, name):
        n = ident(name)
        return f"{ident(self.owner)}.{n}" if self.owner else n

    def kws(self):
        return [k.strip().upper() for k in (self.cfg.get("keywords") or "").split(",") if k.strip()]

    # ---------- 구조 조사 ----------
    def related_tables(self, conn):
        kws = self.kws()
        like = " OR ".join(f"t.TABLE_NAME LIKE :k{i}" for i in range(len(kws))) or "1=0"
        binds = {f"k{i}": f"%{k}%" for i, k in enumerate(kws)}
        binds["o"] = self.owner
        _, rows = self.q(conn, f"""
            SELECT t.TABLE_NAME, t.NUM_ROWS,
              (SELECT c.COMMENTS FROM ALL_TAB_COMMENTS c WHERE c.OWNER = t.OWNER AND c.TABLE_NAME = t.TABLE_NAME),
              (SELECT COUNT(*) FROM ALL_TAB_COLUMNS c WHERE c.OWNER = t.OWNER AND c.TABLE_NAME = t.TABLE_NAME)
            FROM ALL_TABLES t
            WHERE t.OWNER = NVL(:o, USER)
              AND (({like}) OR EXISTS (SELECT 1 FROM ALL_TAB_COLUMNS c WHERE c.OWNER = t.OWNER
                     AND c.TABLE_NAME = t.TABLE_NAME AND c.COLUMN_NAME = 'WORK_ORDER_NO'))
            ORDER BY t.TABLE_NAME""", binds)
        out = []
        for name, nrows, cmt, ncol in rows:
            exact = False
            if nrows is None or nrows < 200000:
                try:
                    _, r = self.q(conn, f"SELECT COUNT(*) FROM {self.tq(name)}")
                    nrows, exact = r[0][0], True
                except Exception:
                    pass
            out.append((name, cmt or "", ncol, nrows, exact))
        return out

    def programs(self, conn):
        kws = self.kws() + ["WORK_ORDER", "WORKORDER", "WO"]
        like = " OR ".join(f"OBJECT_NAME LIKE :k{i}" for i in range(len(kws)))
        binds = {f"k{i}": f"%{k}%" for i, k in enumerate(kws)}
        binds["o"] = self.owner
        _, a = self.q(conn, f"""
            SELECT OBJECT_NAME, OBJECT_TYPE, STATUS, LAST_DDL_TIME, '이름' AS WHY
            FROM ALL_OBJECTS
            WHERE OWNER = NVL(:o, USER)
              AND OBJECT_TYPE IN ('PROCEDURE', 'FUNCTION', 'PACKAGE', 'PACKAGE BODY', 'TRIGGER')
              AND ({like})""", binds)
        _, b = self.q(conn, """
            SELECT DISTINCT s.NAME, s.TYPE, o.STATUS, o.LAST_DDL_TIME, '내용' AS WHY
            FROM ALL_SOURCE s JOIN ALL_OBJECTS o
              ON o.OWNER = s.OWNER AND o.OBJECT_NAME = s.NAME AND o.OBJECT_TYPE = s.TYPE
            WHERE s.OWNER = NVL(:o, USER)
              AND UPPER(s.TEXT) LIKE '%WORK_ORDER%'
              AND (UPPER(s.TEXT) LIKE '%INSERT%' OR UPPER(s.TEXT) LIKE '%UPDATE%')""", {"o": self.owner})
        seen, out = {}, []
        for name, typ, st, ddl, why in a + b:
            k = (name, typ)
            if k in seen:
                seen[k][4] = "이름+내용"
                continue
            row = [name, typ, st, ddl, why]
            seen[k] = row
            out.append(row)
        out.sort(key=lambda r: (0 if "내용" in r[4] else 1, r[0]))
        return out

    def source(self, conn, name, typ):
        _, lines = self.q(conn, "SELECT TEXT FROM ALL_SOURCE WHERE OWNER = NVL(:o, USER) AND NAME = :n "
                                "AND TYPE = :t ORDER BY LINE", {"o": self.owner, "n": name, "t": typ})
        text = "".join(str(l[0] or "") for l in lines)
        args = ""
        if typ in ("PROCEDURE", "FUNCTION", "PACKAGE"):
            try:
                _, ar = self.q(conn, """
                    SELECT NVL(PACKAGE_NAME, '-'), OBJECT_NAME, ARGUMENT_NAME, IN_OUT, DATA_TYPE, POSITION
                    FROM ALL_ARGUMENTS WHERE OWNER = NVL(:o, USER) AND (OBJECT_NAME = :n OR PACKAGE_NAME = :n)
                    ORDER BY PACKAGE_NAME, OBJECT_NAME, SUBPROGRAM_ID, POSITION""", {"o": self.owner, "n": name})
                cur_obj, parts = None, []
                for pkg, obj, arg, io, dtp, pos in ar:
                    if obj != cur_obj:
                        parts.append(f"\n  {obj}(")
                        cur_obj = obj
                    if arg:
                        parts.append(f"      {arg} {io} {dtp}")
                args = "\n".join(parts)
            except Exception:
                pass
        if typ == "TRIGGER" and not text:
            try:
                _, t = self.q(conn, "SELECT TABLE_NAME, TRIGGERING_EVENT, TRIGGER_TYPE, TRIGGER_BODY FROM ALL_TRIGGERS "
                                    "WHERE OWNER = NVL(:o, USER) AND TRIGGER_NAME = :n", {"o": self.owner, "n": name})
                if t:
                    text = f"-- 테이블: {t[0][0]} / {t[0][1]} / {t[0][2]}\n{t[0][3]}"
            except Exception:
                pass
        return text, args

    def sequences(self, conn):
        _, rows = self.q(conn, "SELECT SEQUENCE_NAME, LAST_NUMBER, INCREMENT_BY, MIN_VALUE, MAX_VALUE "
                               "FROM ALL_SEQUENCES WHERE SEQUENCE_OWNER = NVL(:o, USER) ORDER BY SEQUENCE_NAME",
                         {"o": self.owner})
        kws = self.kws() + ["WO", "ORDER"]
        return [r for r in rows if any(k in r[0] for k in kws)] or rows

    def triggers(self, conn, tables):
        out = []
        for t in tables:
            _, rows = self.q(conn, "SELECT TRIGGER_NAME, TABLE_NAME, TRIGGERING_EVENT, TRIGGER_TYPE, STATUS "
                                   "FROM ALL_TRIGGERS WHERE TABLE_OWNER = NVL(:o, USER) AND TABLE_NAME = :t",
                             {"o": self.owner, "t": t})
            out += rows
        return out

    def search_source(self, conn, text):
        _, rows = self.q(conn, "SELECT NAME, TYPE, LINE, TEXT FROM ALL_SOURCE WHERE OWNER = NVL(:o, USER) "
                               "AND UPPER(TEXT) LIKE :k ORDER BY NAME, TYPE, LINE",
                         {"o": self.owner, "k": f"%{text.strip().upper()}%"}, many=1000)
        return rows

    def columns(self, conn, table):
        _, rows = self.q(conn, "SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE, DATA_DEFAULT, "
                               "(SELECT COMMENTS FROM ALL_COL_COMMENTS m WHERE m.OWNER = c.OWNER "
                               " AND m.TABLE_NAME = c.TABLE_NAME AND m.COLUMN_NAME = c.COLUMN_NAME) "
                               "FROM ALL_TAB_COLUMNS c WHERE OWNER = NVL(:o, USER) AND TABLE_NAME = :t "
                               "ORDER BY COLUMN_ID", {"o": self.owner, "t": table})
        return rows

    # ---------- 전후 비교 ----------
    def all_tables(self, conn, only_kw):
        _, rows = self.q(conn, "SELECT TABLE_NAME FROM ALL_TABLES WHERE OWNER = NVL(:o, USER) ORDER BY TABLE_NAME",
                         {"o": self.owner})
        names = [r[0] for r in rows if ID_RE.match(r[0])]
        if only_kw:
            rel = {r[0] for r in self.related_tables(conn)}
            names = [n for n in names if n in rel]
        return names

    def hash_cols(self, conn, table):
        _, rows = self.q(conn, "SELECT COLUMN_NAME, DATA_TYPE FROM ALL_TAB_COLUMNS WHERE OWNER = NVL(:o, USER) "
                               "AND TABLE_NAME = :t ORDER BY COLUMN_ID", {"o": self.owner, "t": table})
        out = []
        for name, dtype in rows:
            d = (dtype or "").upper()
            if ID_RE.match(name) and (d in HASHABLE or d.startswith("TIMESTAMP")):
                out.append((name, d))
        return out

    def snap_table(self, conn, table, small_limit):
        cols = self.hash_cols(conn, table)
        if not cols:
            _, r = self.q(conn, f"SELECT COUNT(*) FROM {self.tq(table)}")
            return {"count": r[0][0], "sums": {}, "fp": None, "cols": []}
        hs = [f'NVL(ORA_HASH("{c}", 4294967295, {i}), 0)' for i, (c, _) in enumerate(cols)]
        _, r = self.q(conn, f"SELECT COUNT(*), {', '.join(f'SUM({h})' for h in hs)} FROM {self.tq(table)}")
        count = r[0][0]
        sums = {c: r[0][i + 1] for i, (c, _) in enumerate(cols)}
        fp = None
        if count <= small_limit:
            _, rr = self.q(conn, f"SELECT ROWIDTOCHAR(ROWID), {' + '.join(hs)} FROM {self.tq(table)}")
            fp = {a: b for a, b in rr}
        return {"count": count, "sums": sums, "fp": fp, "cols": [c for c, _ in cols]}

    def rows_by_rowid(self, conn, table, rowids):
        out, hdr = [], None
        rowids = list(rowids)
        for i in range(0, len(rowids), 200):
            part = rowids[i:i + 200]
            binds = {f"r{j}": r for j, r in enumerate(part)}
            ph = ", ".join(f"CHARTOROWID(:r{j})" for j in range(len(part)))
            hdr, rows = self.q(conn, f"SELECT * FROM {self.tq(table)} WHERE ROWID IN ({ph})", binds)
            out += rows
        return hdr, out

    def rows_since(self, conn, table, since):
        cols = self.columns(conn, table)
        tcols = [c[0] for c in cols if (c[1] or "").upper() == "DATE" and RE_TIME_COL.search(c[0])]
        if not tcols:
            return None, [], []
        cond = " OR ".join(f'"{c}" >= :t' for c in tcols)
        hdr, rows = self.q(conn, f"SELECT * FROM {self.tq(table)} WHERE {cond} AND ROWNUM <= 500",
                           {"t": since - dt.timedelta(minutes=1)})
        return hdr, rows, tcols


# ==============================================================
# 화면
# ==============================================================

class InspectTabs:
    """구조 조사 / 전후 비교 탭 (작업지시 조사 도구 통합, 읽기 전용)"""
    def __init__(self, app, tab_find, tab_diff):
        self.app, self.root = app, app.root
        self.db = Inspector(app)
        self.snap, self.snap_time, self.diff, self.diff_view = {}, None, [], []
        self.tabs = {"find": tab_find, "diff": tab_diff}
        self.build_find()
        self.build_diff()

    @property
    def cfg(self):
        return self.app.cfg

    def busy(self, on, text=None):
        if text is not None:
            self.status.config(text=text)
        self.app.busy(on)

    def run_db(self, fn):
        return self.app.run_db(fn)

    def err(self, e):
        self.app.err(e)

    def make_tree(self, parent, cols, widths=None):
        fr, tv = self.app.make_tree(parent, cols, widths)
        tv.tag_configure("hot", background="#FFEB9C")
        tv.tag_configure("new", background="#C6EFCE")
        tv.tag_configure("upd", background="#FCE4D6")
        tv.tag_configure("del", background="#FFC7CE")
        return fr, tv

    @staticmethod
    def fill_tree(tv, cols, rows, widths=None, tags=None):
        tv.delete(*tv.get_children())
        ids = [f"c{i}" for i in range(len(cols))]
        tv["columns"] = ids
        if widths is None:
            widths = []
            for i, c in enumerate(cols):
                w = len(str(c))
                for r in rows[:200]:
                    if i < len(r):
                        w = max(w, len(str(r[i])[:40]))
                widths.append(max(60, min(360, w * 8 + 20)))
        for i, (cid, c) in enumerate(zip(ids, cols)):
            tv.heading(cid, text=c)
            tv.column(cid, width=widths[i], anchor="w", stretch=False)
        for k, r in enumerate(rows):
            tv.insert("", "end", iid=str(k), values=[cell(v) for v in r],
                      tags=(tags[k],) if tags and tags[k] else ())

    def text_window(self, title, text, head=""):
        w = tk.Toplevel(self.root)
        w.title(title)
        w.geometry("1000x700")
        bar = ttk.Frame(w)
        bar.pack(fill="x", padx=8, pady=(8, 0))
        ttk.Label(bar, text="찾기").pack(side="left")
        qv = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=qv, width=30)
        ent.pack(side="left", padx=4)
        info = ttk.Label(bar, text="")
        info.pack(side="left", padx=8)
        full = (head + "\n\n" if head else "") + text

        def copy_all():
            w.clipboard_clear()
            w.clipboard_append(full)
            info.config(text="전체 복사됨 - 채팅에 붙여넣기 하세요")
        ttk.Button(bar, text="전체 복사", command=copy_all).pack(side="right")
        fr = ttk.Frame(w)
        fr.pack(fill="both", expand=True, padx=8, pady=8)
        t = tk.Text(fr, wrap="none", font=("Consolas", 10))
        ys = ttk.Scrollbar(fr, orient="vertical", command=t.yview)
        xs = ttk.Scrollbar(fr, orient="horizontal", command=t.xview)
        t.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        t.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        fr.rowconfigure(0, weight=1)
        fr.columnconfigure(0, weight=1)
        t.tag_configure("hit", background="#FFEB9C")
        t.tag_configure("dml", foreground="#C00000", font=("Consolas", 10, "bold"))
        t.insert("1.0", (head + "\n\n" if head else "") + text)
        start = "1.0"
        while True:                                            # INSERT/UPDATE/DELETE 강조
            pos = t.search(r"\m(INSERT|UPDATE|DELETE|MERGE)\M", start, "end", regexp=True, nocase=True)
            if not pos:
                break
            end = t.index(f"{pos} lineend")
            t.tag_add("dml", pos, end)
            start = end
        t.configure(state="disabled")

        def find(*_):
            t.tag_remove("hit", "1.0", "end")
            q = qv.get().strip()
            if not q:
                info.config(text="")
                return
            n, pos = 0, "1.0"
            first = None
            while True:
                pos = t.search(q, pos, "end", nocase=True)
                if not pos:
                    break
                end = f"{pos}+{len(q)}c"
                t.tag_add("hit", pos, end)
                first = first or pos
                n += 1
                pos = end
            info.config(text=f"{n}곳")
            if first:
                t.see(first)
        qv.trace_add("write", find)
        ent.focus_set()

    # ==========================================================
    # 구조 조사
    # ==========================================================
    def build_find(self):
        t = self.tabs["find"]
        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=(6, 2))
        ttk.Button(bar, text="① 관련 테이블", style="Big.TButton", command=self.on_tables).pack(side="left", padx=3)
        ttk.Button(bar, text="② 작업지시 만드는 프로그램", style="Big.TButton", command=self.on_programs).pack(side="left", padx=3)
        ttk.Button(bar, text="③ 번호 생성기", style="Big.TButton", command=self.on_sequences).pack(side="left", padx=3)
        ttk.Button(bar, text="④ 트리거", style="Big.TButton", command=self.on_triggers).pack(side="left", padx=3)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Label(bar, text="프로그램 내용 검색").pack(side="left")
        self.var_src = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.var_src, width=24)
        e.pack(side="left", padx=4)
        e.bind("<Return>", lambda _e: self.on_search_src())
        ttk.Button(bar, text="검색", command=self.on_search_src).pack(side="left")
        ttk.Button(bar, text="엑셀로 저장", command=self.on_export_find).pack(side="right", padx=3)
        ttk.Button(bar, text="선택 항목 내용 txt 저장", command=self.on_export_src).pack(side="right", padx=3)
        self.lbl_find = ttk.Label(t, text="버튼을 누르면 결과가 나옵니다. 줄을 더블클릭하면 자세히 봅니다.", foreground="#555")
        self.lbl_find.pack(fill="x", padx=4)
        fr, self.tv_find = self.make_tree(t, ["-"])
        fr.pack(fill="both", expand=True, pady=4)
        self.tv_find.bind("<Double-1>", self.on_find_detail)
        self.find_kind, self.find_cols, self.find_rows = None, [], []
        self.status = ttk.Label(t, text="")
        self.status.pack(fill="x", padx=4, pady=(0, 4))

    def show_find(self, kind, cols, rows, msg, tags=None):
        self.find_kind, self.find_cols, self.find_rows = kind, cols, rows
        self.fill_tree(self.tv_find, cols, rows, tags=tags)
        self.lbl_find.config(text=msg)

    def on_tables(self):
        try:
            rows = self.run_db(self.db.related_tables)
        except Exception as e:
            self.err(e)
            return
        data = [[n, c, k, ("" if r is None else f"{r:,}") + ("" if ex else " (추정)")] for n, c, k, r, ex in rows]
        self.show_find("table", ["테이블", "설명", "컬럼수", "행수"], data,
                       f"작업지시 관련 테이블 {len(data)}개 (이름에 키워드가 있거나 WORK_ORDER_NO 컬럼이 있는 것). "
                       f"더블클릭 = 컬럼 구조")

    def on_programs(self):
        try:
            rows = self.run_db(self.db.programs)
        except Exception as e:
            self.err(e)
            return
        tags = ["hot" if "내용" in r[4] else "" for r in rows]
        self.show_find("prog", ["이름", "종류", "상태", "마지막 변경", "찾은 이유"], rows,
                       f"프로그램 {len(rows)}개.  노란 줄 = 안에서 작업지시 테이블에 INSERT/UPDATE 함 (가장 중요).  "
                       f"더블클릭 = 내용 보기", tags)

    def on_sequences(self):
        try:
            rows = self.run_db(self.db.sequences)
        except Exception as e:
            self.err(e)
            return
        self.show_find("seq", ["번호생성기", "다음 번호 근처", "증가", "최소", "최대"], rows,
                       f"번호 생성기(시퀀스) {len(rows)}개. 작업지시 번호가 여기서 나오는지 확인하는 용도입니다.")

    def on_triggers(self):
        def work(conn):
            names = [r[0] for r in self.db.related_tables(conn)]
            return self.db.triggers(conn, names)
        try:
            rows = self.run_db(work)
        except Exception as e:
            self.err(e)
            return
        self.show_find("trg", ["트리거", "테이블", "언제", "방식", "상태"], rows,
                       f"관련 테이블에 걸린 자동동작(트리거) {len(rows)}개. 더블클릭 = 내용 보기")

    def on_search_src(self):
        q = self.var_src.get().strip()
        if len(q) < 3:
            messagebox.showinfo(APP_TITLE, "3글자 이상 입력하세요. 예: IPRD_WORK_ORDER_MASTER")
            return
        try:
            rows = self.run_db(lambda c: self.db.search_source(c, q))
        except Exception as e:
            self.err(e)
            return
        tags = ["hot" if re.search(r"\b(INSERT|UPDATE|DELETE|MERGE)\b", str(r[3]).upper()) else "" for r in rows]
        self.show_find("src", ["프로그램", "종류", "줄", "내용"], rows,
                       f"'{q}' 가 들어간 곳 {len(rows)}줄{' (1000줄까지)' if len(rows) >= 1000 else ''}.  "
                       f"노란 줄 = 쓰기(INSERT/UPDATE/DELETE).  더블클릭 = 프로그램 전체 보기", tags)

    def on_find_detail(self, _e=None):
        sel = self.tv_find.selection()
        if not sel:
            return
        r = self.find_rows[int(sel[0])]
        kind = self.find_kind
        try:
            if kind == "table":
                cols = self.run_db(lambda c: self.db.columns(c, r[0]))
                w = tk.Toplevel(self.root)
                w.title(f"{r[0]} 컬럼 구조")
                w.geometry("900x600")
                fr, tv = self.make_tree(w, ["컬럼"])
                fr.pack(fill="both", expand=True, padx=8, pady=8)
                self.fill_tree(tv, ["컬럼", "형식", "길이", "빈값허용", "DB기본값", "설명"],
                               [[a, b, c, d, str(e or "").strip(), f or ""] for a, b, c, d, e, f in cols])
                return
            if kind in ("prog", "trg", "src"):
                name = r[0]
                typ = {"prog": r[1], "trg": "TRIGGER", "src": r[1]}[kind]
                text, args = self.run_db(lambda c: self.db.source(c, name, typ))
                head = f"-- {typ} {name}" + (f"\n-- 입력값(파라미터):{args}" if args else "")
                if not text:
                    text = "(내용을 볼 권한이 없거나 암호화된 프로그램입니다)"
                self.text_window(f"{typ} {name}", text, head)
        except Exception as e:
            self.err(e)

    def on_export_src(self):
        """선택한 트리거/프로그램 내용을 한 txt 파일로"""
        if self.find_kind not in ("prog", "trg", "src"):
            messagebox.showinfo(APP_TITLE, "② 프로그램, ④ 트리거, 내용 검색 결과에서 쓸 수 있습니다.")
            return
        sel = self.tv_find.selection() or self.tv_find.get_children()
        items, seen = [], set()
        for i in sel:
            r = self.find_rows[int(i)]
            typ = {"prog": r[1], "trg": "TRIGGER", "src": r[1]}[self.find_kind]
            if (r[0], typ) not in seen:
                seen.add((r[0], typ))
                items.append((r[0], typ))
        if len(items) > 50 and not messagebox.askyesno(APP_TITLE, f"{len(items)}개입니다. 모두 저장할까요?"):
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt",
                                            initialfile=f"작업지시_{self.find_kind}_내용_{dt.datetime.now():%Y%m%d_%H%M}.txt",
                                            filetypes=[("Text", "*.txt")])
        if not path:
            return

        def work(conn):
            out = []
            for k, (name, typ) in enumerate(items, 1):
                self.status.config(text=f"읽는 중 {k}/{len(items)} {name}")
                self.root.update()
                text, args = self.db.source(conn, name, typ)
                out.append(f"{'=' * 70}\n-- {typ} {name}" + (f"\n-- 입력값:{args}" if args else "") +
                           f"\n{'=' * 70}\n{text or '(내용을 볼 권한이 없거나 암호화됨)'}\n")
            return out
        try:
            parts = self.run_db(work)
        except Exception as e:
            self.err(e)
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(parts))
        self.status.config(text="")
        messagebox.showinfo(APP_TITLE, f"{len(parts)}개 내용을 저장했습니다.\n{path}")

    def on_export_find(self):
        if not self.find_rows:
            messagebox.showinfo(APP_TITLE, "먼저 조회하세요.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialfile=f"작업지시조사_{self.find_kind}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(self.find_cols)
        for r in self.find_rows:
            ws.append([cell(v, 2000) for v in r])
        for c in ws[1]:
            c.font = Font(bold=True)
        wb.save(path)
        messagebox.showinfo(APP_TITLE, f"저장했습니다.\n{path}")

    # ==========================================================
    # 전후 비교
    # ==========================================================
    def build_diff(self):
        t = self.tabs["diff"]
        guide = ttk.Label(t, foreground="#1F4E79", font=("", 10, "bold"),
                          text="순서:  ① [찍기 전 저장]  →  ② MES 화면에서 작업지시 1건 내기  →  ③ [찍은 후 비교]   "
                               "(②는 빨리 하세요. 그 사이 다른 작업도 같이 잡힙니다)")
        guide.pack(fill="x", padx=4, pady=(8, 2))
        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=(2, 4))
        self.var_scope = tk.StringVar(value="kw")
        ttk.Radiobutton(bar, text="관련 테이블만 (빠름)", value="kw", variable=self.var_scope).pack(side="left", padx=4)
        ttk.Radiobutton(bar, text="스키마 전체 (느림, 빠짐없이)", value="all", variable=self.var_scope).pack(side="left", padx=4)
        ttk.Button(bar, text="① 찍기 전 저장", style="Big.TButton", command=self.on_snap).pack(side="left", padx=(16, 3))
        ttk.Button(bar, text="③ 찍은 후 비교", style="Big.TButton", command=self.on_compare).pack(side="left", padx=3)
        ttk.Button(bar, text="결과 엑셀 저장", command=self.on_export_diff).pack(side="right", padx=3)
        self.var_hide = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="변화 없는 테이블 숨김", variable=self.var_hide,
                        command=self.show_diff).pack(side="right", padx=8)
        self.lbl_diff = ttk.Label(t, text="아직 저장한 상태가 없습니다.", foreground="#555")
        self.lbl_diff.pack(fill="x", padx=4)
        fr, self.tv_diff = self.make_tree(t, ["-"])
        fr.pack(fill="both", expand=True, pady=4)
        self.tv_diff.bind("<Double-1>", self.on_diff_detail)
        self.diff_status = ttk.Label(t, text="")
        self.diff_status.pack(fill="x", padx=4, pady=(0, 4))

    def progress(self, text):
        self.diff_status.config(text=text)
        self.root.update()

    def take_snapshot(self, conn):
        names = self.db.all_tables(conn, self.var_scope.get() == "kw")
        lim = int(self.cfg.get("small_limit") or 30000)
        snap, skipped = {}, []
        for i, n in enumerate(names, 1):
            self.progress(f"상태 저장 중... {i}/{len(names)}  {n}")
            try:
                snap[n] = self.db.snap_table(conn, n, lim)
            except Exception as e:
                skipped.append(f"{n}({str(e)[:40]})")
        return snap, skipped

    def on_snap(self):
        try:
            self.busy(True)
            snap, skipped = self.run_db(self.take_snapshot)
        except Exception as e:
            self.err(e)
            return
        self.snap, self.snap_time, self.diff = snap, dt.datetime.now(), []
        self.fill_tree(self.tv_diff, ["-"], [])
        self.progress("")
        msg = f"{self.snap_time:%H:%M:%S} 상태 저장 완료 ({len(snap)}개 테이블)."
        if skipped:
            msg += f"  읽지 못한 테이블 {len(skipped)}개"
        self.lbl_diff.config(text=msg + "  → 이제 MES에서 작업지시 1건을 내고 [③ 찍은 후 비교]를 누르세요.",
                             foreground="#C00000")
        messagebox.showinfo(APP_TITLE, msg + "\n\n이제 MES 화면에서 작업지시 1건을 내세요.\n다 되면 [③ 찍은 후 비교]를 누르세요.")

    def compare(self, conn):
        lim = int(self.cfg.get("small_limit") or 30000)
        out = []
        names = list(self.snap)
        for i, n in enumerate(names, 1):
            self.progress(f"비교 중... {i}/{len(names)}  {n}")
            before = self.snap[n]
            try:
                after = self.db.snap_table(conn, n, lim)
            except Exception:
                continue
            changed_cols = [c for c in before["sums"] if before["sums"].get(c) != after["sums"].get(c)]
            d = {"table": n, "before": before["count"], "after": after["count"], "cols": changed_cols,
                 "new": [], "upd": [], "del": 0, "hdr": [], "how": ""}
            if before["count"] != after["count"] or changed_cols:
                if before["fp"] is not None and after["fp"] is not None:
                    newids = [r for r in after["fp"] if r not in before["fp"]]
                    updids = [r for r in after["fp"] if r in before["fp"] and after["fp"][r] != before["fp"][r]]
                    d["del"] = sum(1 for r in before["fp"] if r not in after["fp"])
                    if newids:
                        d["hdr"], d["new"] = self.db.rows_by_rowid(conn, n, newids[:500])
                    if updids:
                        h, d["upd"] = self.db.rows_by_rowid(conn, n, updids[:500])
                        d["hdr"] = d["hdr"] or h
                    d["how"] = "줄 단위 비교"
                else:
                    h, rows, tcols = self.db.rows_since(conn, n, self.snap_time)
                    if h:
                        d["hdr"], d["new"] = h, rows
                        d["how"] = f"큰 테이블 - {', '.join(tcols)} 기준 최근 줄"
                    else:
                        d["how"] = "큰 테이블 - 줄 내용은 못 봄 (건수/컬럼만)"
            out.append(d)
        return out

    def on_compare(self):
        if not self.snap:
            messagebox.showinfo(APP_TITLE, "먼저 [① 찍기 전 저장]을 하세요.")
            return
        try:
            self.busy(True)
            self.diff = self.run_db(self.compare)
        except Exception as e:
            self.err(e)
            return
        self.progress("")
        self.show_diff()
        ch = [d for d in self.diff if d["before"] != d["after"] or d["cols"]]
        messagebox.showinfo(APP_TITLE, f"비교 완료.\n바뀐 테이블 {len(ch)}개\n\n"
                                       f"초록 = 새 줄이 생김 / 주황 = 기존 줄이 바뀜 / 빨강 = 줄이 없어짐\n"
                                       f"줄을 더블클릭하면 실제 내용을 봅니다.")

    def show_diff(self):
        if not self.diff:
            return
        rows, tags, self.diff_view = [], [], []
        for d in sorted(self.diff, key=lambda x: (-(abs(x["after"] - x["before"]) + len(x["new"]) + len(x["upd"])),
                                                   x["table"])):
            changed = d["before"] != d["after"] or d["cols"]
            if self.var_hide.get() and not changed:
                continue
            self.diff_view.append(d)
            diffn = d["after"] - d["before"]
            rows.append([d["table"], f"{d['before']:,}", f"{d['after']:,}", f"{diffn:+,}" if diffn else "0",
                         len(d["new"]), len(d["upd"]), d["del"], ", ".join(d["cols"][:12]) +
                         (" …" if len(d["cols"]) > 12 else ""), d["how"]])
            tags.append("del" if d["del"] and not d["new"] else ("new" if d["new"] or diffn > 0 else
                                                                 ("upd" if changed else "")))
        self.fill_tree(self.tv_diff, ["테이블", "전", "후", "증감", "새 줄", "바뀐 줄", "없어진 줄", "바뀐 컬럼", "확인 방법"],
                       rows, [220, 70, 70, 60, 55, 60, 70, 380, 260], tags)
        n = sum(1 for d in self.diff if d["before"] != d["after"] or d["cols"])
        self.lbl_diff.config(text=f"저장 {self.snap_time:%H:%M:%S} 대비 바뀐 테이블 {n}개 / 전체 {len(self.diff)}개.  "
                                  f"더블클릭 = 새 줄·바뀐 줄 내용", foreground="#1F4E79")

    def on_diff_detail(self, _e=None):
        sel = self.tv_diff.selection()
        if not sel:
            return
        d = self.diff_view[int(sel[0])]
        if not d["hdr"]:
            messagebox.showinfo(APP_TITLE, f"{d['table']}: 줄 내용을 볼 수 없습니다.\n({d['how']})\n"
                                           f"바뀐 컬럼: {', '.join(d['cols']) or '없음'}")
            return
        w = tk.Toplevel(self.root)
        w.title(f"{d['table']} - 새 줄 {len(d['new'])} / 바뀐 줄 {len(d['upd'])}")
        w.geometry("1150x560")
        ttk.Label(w, text=f"초록 = 새로 생긴 줄,  주황 = 값이 바뀐 줄 (지금 값).  바뀐 컬럼: {', '.join(d['cols']) or '-'}",
                  foreground="#555").pack(fill="x", padx=8, pady=(8, 0))
        fr, tv = self.make_tree(w, ["-"])
        fr.pack(fill="both", expand=True, padx=8, pady=8)
        rows = [["새 줄"] + list(r) for r in d["new"]] + [["바뀐 줄"] + list(r) for r in d["upd"]]
        tags = ["new"] * len(d["new"]) + ["upd"] * len(d["upd"])
        self.fill_tree(tv, ["구분"] + list(d["hdr"]), rows, tags=tags)

        def vertical(_e=None):
            s = tv.selection()
            if not s:
                return
            r = rows[int(s[0])]
            vw = tk.Toplevel(w)
            vw.title(f"{d['table']} - {r[0]}")
            vw.geometry("620x640")
            vf, vt = self.make_tree(vw, ["-"])
            vf.pack(fill="both", expand=True, padx=8, pady=8)
            self.fill_tree(vt, ["컬럼", "값"], [[h, v] for h, v in zip(d["hdr"], r[1:])], [220, 360])
        tv.bind("<Double-1>", vertical)

    def on_export_diff(self):
        if not self.diff:
            messagebox.showinfo(APP_TITLE, "먼저 비교하세요.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialfile=f"작업지시_전후비교_{dt.datetime.now():%Y%m%d_%H%M}.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "요약"
        ws.append([f"저장 시각 {self.snap_time:%Y-%m-%d %H:%M:%S}"])
        ws.append(["테이블", "전", "후", "증감", "새 줄", "바뀐 줄", "없어진 줄", "바뀐 컬럼", "확인 방법"])
        for c in ws[2]:
            c.font = Font(bold=True)
        used = set()
        for d in self.diff:
            if not (d["before"] != d["after"] or d["cols"]):
                continue
            ws.append([d["table"], d["before"], d["after"], d["after"] - d["before"], len(d["new"]),
                       len(d["upd"]), d["del"], ", ".join(d["cols"]), d["how"]])
            if d["hdr"] and (d["new"] or d["upd"]):
                name = d["table"][:28]
                k = 1
                while name in used:
                    k += 1
                    name = f"{d['table'][:25]}_{k}"
                used.add(name)
                sh = wb.create_sheet(name)
                sh.append(["구분"] + list(d["hdr"]))
                for c in sh[1]:
                    c.font = Font(bold=True)
                for tag, rows, color in (("새 줄", d["new"], "C6EFCE"), ("바뀐 줄", d["upd"], "FCE4D6")):
                    for r in rows:
                        sh.append([tag] + [cell(v, 2000) for v in r])
                        sh.cell(sh.max_row, 1).fill = PatternFill("solid", start_color=color, end_color=color)
        wb.save(path)
        messagebox.showinfo(APP_TITLE, f"저장했습니다.\n{path}")

    # ==========================================================
    # 설정
    # ==========================================================
    SET_FIELDS = [
        ("oracle.dsn", "접속 주소(DSN)", "예: 192.168.1.250:1521/XE"),
        ("oracle.lib_dir", "Instant Client 폴더", "11g는 필수"),
        ("schema", "스키마", "MES 스키마"),
        ("keywords", "찾을 단어", "쉼표로 구분. 테이블/프로그램 이름에 이 단어가 있으면 찾음"),
        ("small_limit", "줄 단위 비교 한도", "이 줄 수 이하 테이블은 새 줄·바뀐 줄 내용까지 비교"),
    ]


# ==============================================================
# 간편 금형등록 (한 화면에서 금형 + 품번연결 + 단위중량을 한 번에, 추가만)
# ==============================================================
RESET_WORDS = ("SHOT", "HIT", "BARCODE", "QR", "LAST_RECEIPT", "LAST_ISSUE", "ACTUAL", "SERIAL", "BUY_DATE", "IMAGE")
QUICK_FIELDS = [("MOLD_CODE", "금형코드 *"), ("MOLD_NAME", "금형명 *"), ("MOLD_GROUP", "톤수(그룹)"),
                ("DRAWING_NO", "도면번호"), ("MOLD_SPEC", "금형사양"), ("CUSTOMER_CODE", "고객"),
                ("SUPPLIER_CODE", "제작처"), ("CAVITY_QTY", "캐비티")]
BOM_KEYS = ("ORGANIZATION_ID", "ITEM_CODE", "CHILD_ITEM_CODE", "WORKSTAGE_CODE", "CHILD_WORKSTAGE_CODE")


class QuickTab:
    def __init__(self, app, parent):
        self.app, self.mes = app, app.mes
        self.base = None            # {"mold": {..}, "mold_desc": [...], "link": {..}, "link_desc", "bom": {..}, "bom_desc"}
        self.items = []             # [{"item", "name", "raw", "qpu", "cav", "note"}]
        self.checked = None
        wrap = ttk.Frame(parent, padding=8)
        wrap.pack(fill="both", expand=True)

        f1 = ttk.LabelFrame(wrap, text=" ① 비슷한 기존 금형 고르기 (회사코드·그룹 등 나머지 칸을 이 금형에서 복사합니다) ", padding=8)
        f1.pack(fill="x")
        self.var_bq = tk.StringVar(value=app.cfg.get("ref_mold", ""))
        e = ttk.Entry(f1, textvariable=self.var_bq, width=24)
        e.pack(side="left")
        e.bind("<Return>", lambda _e: self.find_base())
        ttk.Button(f1, text="찾기", command=self.find_base).pack(side="left", padx=4)
        self.cb_base = ttk.Combobox(f1, width=60, state="readonly")
        self.cb_base.pack(side="left", padx=8)
        self.cb_base.bind("<<ComboboxSelected>>", lambda _e: self.load_base())
        self.lbl_base = ttk.Label(f1, text="금형코드·금형명 일부를 넣고 [찾기] (비우고 누르면 전체 목록)", foreground="#555")
        self.lbl_base.pack(side="left", padx=8)

        f2 = ttk.LabelFrame(wrap, text=" ② 새 금형 정보 ", padding=8)
        f2.pack(fill="x", pady=6)
        self.fvars = {}
        for k, (col, label) in enumerate(QUICK_FIELDS):
            r, c = divmod(k, 4)
            ttk.Label(f2, text=label, font=("", 9, "bold") if "*" in label else ("", 9)).grid(row=r, column=c * 2, sticky="e", padx=(8, 3), pady=3)
            v = tk.StringVar()
            ttk.Entry(f2, textvariable=v, width=22).grid(row=r, column=c * 2 + 1, sticky="w", pady=3)
            self.fvars[col] = v
        self.fvars["MOLD_CODE"].trace_add("write", lambda *_: self.on_code())

        f3 = ttk.LabelFrame(wrap, text=" ③ 이 금형으로 만드는 품번 (품번 하나는 금형 하나에만 연결됩니다) ", padding=8)
        f3.pack(fill="both", expand=True)
        b3 = ttk.Frame(f3)
        b3.pack(fill="x")
        ttk.Button(b3, text="+ 품번 추가", style="Big.TButton", command=self.add_item).pack(side="left")
        ttk.Button(b3, text="선택 삭제", command=self.del_item).pack(side="left", padx=4)
        ttk.Label(b3, text="칸 더블클릭 = 원소재·단위중량·캐비티 고치기.  피치당 소재중량(g) = 두께×폭×피치×7.85÷1000 (스크랩 포함, 캐비티 2↑는 ÷캐비티)",
                  foreground="#555").pack(side="left", padx=10)
        fr, self.tv = app.make_tree(f3, ["품번", "품명", "원소재(코일) 코드", "피치당 소재중량(g)", "캐비티", "상태"],
                                   [150, 220, 180, 90, 70, 460])
        fr.pack(fill="both", expand=True, pady=4)
        self.tv.bind("<Double-1>", self.edit_item)

        f4 = ttk.Frame(wrap)
        f4.pack(fill="x", pady=(4, 0))
        ttk.Button(f4, text="④ 검사", style="Big.TButton", command=self.check).pack(side="left")
        ttk.Button(f4, text="⑤ MES에 등록", style="Big.TButton", command=self.register).pack(side="left", padx=6)
        ttk.Button(f4, text="새로 시작", command=self.reset).pack(side="left", padx=6)
        self.lbl = ttk.Label(f4, text="", font=("", 10, "bold"))
        self.lbl.pack(side="left", padx=10)

    # ---------- 공통 ----------
    def T(self, key):
        return {"mold": self.mes.table("mold"), "link": self.mes.table("item"),
                "bom": self.mes.qualify(self.app.cfg.get("bom_table") or "ICOM_ITEM_CHILD"),
                "mst": self.mes.qualify(self.app.cfg.get("item_master_table") or "ICOM_ITEM_MASTER")}[key]

    @staticmethod
    def fetch_one(conn, sql, binds):
        cur = conn.cursor()
        cur.execute(sql, binds)
        r = cur.fetchone()
        return (dict(zip([d[0] for d in cur.description], r)) if r else None), cur.description

    def invalidate(self):
        self.checked = None
        self.lbl.config(text="(검사 전)", foreground="#333")

    # ---------- ① 기준 금형 ----------
    def find_base(self):
        q = self.var_bq.get().strip().upper()      # 비우고 [찾기] = 전체 금형 목록

        def work(conn):
            cur = conn.cursor()
            cur.execute(f"SELECT * FROM (SELECT MOLD_CODE, MOLD_NAME FROM {self.T('mold')} WHERE UPPER(MOLD_CODE) LIKE :q "
                        f"OR UPPER(MOLD_NAME) LIKE :q ORDER BY MOLD_CODE) WHERE ROWNUM <= 500", q=f"%{q}%")
            return cur.fetchall()
        try:
            rows = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return
        self.base_list = [(show(a).strip(), show(b).strip()) for a, b in rows]
        self.cb_base["values"] = [f"{a}  |  {b}" for a, b in self.base_list]
        if not rows:
            self.lbl_base.config(text="찾은 금형이 없습니다.", foreground="#C00000")
            return
        self.cb_base.current(0)
        self.lbl_base.config(text=f"{len(rows)}개 찾음 - 목록에서 고르면 아래 칸이 채워집니다", foreground="#1F4E79")
        self.load_base()

    def load_base(self):
        i = self.cb_base.current()
        if i < 0:
            return
        code = self.base_list[i][0]

        def work(conn):
            mold, md = self.fetch_one(conn, f"SELECT * FROM {self.T('mold')} WHERE TRIM(MOLD_CODE) = :c AND ROWNUM = 1", {"c": code})
            link, ld = self.fetch_one(conn, f"SELECT * FROM {self.T('link')} WHERE TRIM(MOLD_CODE) = :c AND ROWNUM = 1", {"c": code})
            bom, bd = None, None
            if link:
                bom, bd = self.fetch_one(conn, f"SELECT * FROM {self.T('bom')} WHERE ITEM_CODE = :i AND ROWNUM = 1",
                                         {"i": link["ITEM_CODE"]})
            if not bd:
                _, bd = self.fetch_one(conn, f"SELECT * FROM {self.T('bom')} WHERE ROWNUM = 1", {})
                bom = None
            if not ld:
                _, ld = self.fetch_one(conn, f"SELECT * FROM {self.T('link')} WHERE ROWNUM = 1", {})
            return mold, md, link, ld, bom, bd
        try:
            mold, md, link, ld, bom, bd = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return
        if not mold:
            return
        self.base = {"mold": mold, "mold_desc": md, "link": link, "link_desc": ld, "bom": bom, "bom_desc": bd}
        cols = {d[0] for d in md}
        for col, v in self.fvars.items():          # 기준 금형 값으로 아래 칸 채우기 (고를 때마다 새로 채움)
            if col == "MOLD_CODE":
                continue                            # 금형코드는 새로 넣어야 하므로 그대로 둠
            if col not in cols and col != "CAVITY_QTY":
                v.set("(이 테이블에 없음)")
                continue
            src = mold.get(col) if col in cols else (link or {}).get("CAVITY_QTY")
            v.set(show(src).strip() if src is not None else "")
        warn = "" if link else "  ※ 이 금형은 품번 연결이 없어 품번연결 칸은 다른 금형 형식으로 채웁니다"
        need = "" if self.fvars["MOLD_CODE"].get().strip() else "  → 새 금형코드를 넣고 금형명을 고치세요"
        self.lbl_base.config(text=f"기준: {code} ({show(mold.get('MOLD_NAME')).strip()}){need}{warn}", foreground="#1F4E79")
        self.invalidate()

    def on_code(self):
        """도면번호가 금형코드와 같은 규칙이면 따라가게"""
        if self.base and show(self.base["mold"].get("DRAWING_NO")).strip() == show(self.base["mold"].get("MOLD_CODE")).strip():
            if "DRAWING_NO" in self.fvars:
                self.fvars["DRAWING_NO"].set(self.fvars["MOLD_CODE"].get().strip())
        self.invalidate()

    # ---------- ③ 품번 ----------
    def show_items(self):
        self.tv.delete(*self.tv.get_children())
        for i, it in enumerate(self.items):
            tag = {"err": "err", "warn": "warn", "ok": "ok"}.get(it.get("level", ""), "")
            self.tv.insert("", "end", iid=str(i), tags=(tag,) if tag else (),
                           values=[it["item"], it.get("name", ""), it.get("raw", ""), show(it.get("qpu")) if it.get("qpu") is not None else "",
                                   show(it.get("cav")) if it.get("cav") is not None else "", it.get("note", "")])

    def add_item(self):
        code = simpledialog.askstring("품번 추가", "제품 품번을 입력하세요:", parent=self.app.root)
        if not code or not code.strip():
            return
        code = code.strip()
        if any(it["item"].upper() == code.upper() for it in self.items):
            messagebox.showinfo(APP_TITLE, "이미 목록에 있는 품번입니다.")
            return

        def work(conn):
            cur = conn.cursor()
            cur.execute(f"SELECT MAX(ITEM_NAME) FROM {self.T('mst')} WHERE ITEM_CODE = :i", i=code)
            name = cur.fetchone()[0]
            cur.execute(f"SELECT CHILD_ITEM_CODE, UNIT_PER_QTY FROM {self.T('bom')} WHERE ITEM_CODE = :i "
                        f"ORDER BY NVL(UNIT_PER_QTY, 0) DESC", i=code)
            boms = cur.fetchall()
            return name, boms
        try:
            name, boms = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return
        it = {"item": code, "name": show(name) if name else "(품목 마스터에 없음)", "raw": "", "qpu": None,
              "cav": None, "note": ""}
        if boms:
            it["raw"] = show(boms[0][0]).strip()
            it["qpu"] = boms[0][1]
            it["note"] = f"기존 단위중량 있음 ({len(boms)}건) - 그대로 둡니다"
        elif self.base and self.base.get("bom"):
            it["raw"] = show(self.base["bom"].get("CHILD_ITEM_CODE")).strip()
            it["note"] = "원소재는 기준 금형 값을 넣었습니다. 확인 후 단위중량을 입력하세요"
        cav = self.fvars["CAVITY_QTY"].get().strip()
        it["cav"] = conv({"type": "NUM"}, cav) if cav and cav.replace(".", "").isdigit() else None
        self.items.append(it)
        self.show_items()
        self.invalidate()

    def del_item(self):
        for i in sorted((int(x) for x in self.tv.selection()), reverse=True):
            del self.items[i]
        self.show_items()
        self.invalidate()

    def edit_item(self, e):
        iid, colid = self.tv.identify_row(e.y), self.tv.identify_column(e.x)
        if not iid:
            return
        ci = int(colid[1:]) - 1
        key = {2: "raw", 3: "qpu", 4: "cav"}.get(ci)
        if not key:
            messagebox.showinfo(APP_TITLE, "원소재·단위중량·캐비티 칸만 고칠 수 있습니다. 품번이 틀리면 삭제 후 다시 추가하세요.")
            return
        it = self.items[int(iid)]
        title = {"raw": "원소재(코일) 코드", "qpu": "피치당 소재중량(g) = 두께×폭×피치×7.85÷1000 (캐비티 2 이상이면 ÷캐비티)", "cav": "캐비티"}[key]
        v = simpledialog.askstring(title, f"{it['item']} - {title}:", initialvalue=show(it.get(key)) if it.get(key) is not None else "",
                                   parent=self.app.root)
        if v is None:
            return
        v = v.strip()
        if key == "raw":
            it[key] = v
        else:
            try:
                it[key] = conv({"type": "NUM"}, v) if v else None
                if it[key] is not None and it[key] < 0:
                    raise ValueError("음수")
            except ValueError:
                messagebox.showerror(APP_TITLE, "숫자를 입력하세요.")
                return
        self.show_items()
        self.invalidate()

    # ---------- ④ 검사 ----------
    def check(self, quiet=False):
        if not self.base:
            messagebox.showinfo(APP_TITLE, "① 비슷한 기존 금형을 먼저 고르세요.")
            return None
        code = self.fvars["MOLD_CODE"].get().strip()
        name = self.fvars["MOLD_NAME"].get().strip()
        errs = []
        if not code:
            errs.append("금형코드를 입력하세요.")
        if not name:
            errs.append("금형명을 입력하세요.")
        mcols = {d[0]: d for d in self.base["mold_desc"]}
        if code and "MOLD_CODE" in mcols and mcols["MOLD_CODE"][3] and blen(code, int(self.app.cfg.get("kor_bytes") or 3)) > mcols["MOLD_CODE"][3]:
            errs.append(f"금형코드가 너무 깁니다 (최대 {mcols['MOLD_CODE'][3]}byte).")
        seen = set()
        for it in self.items:
            if it["item"].upper() in seen:
                errs.append(f"품번 {it['item']}가 두 번 있습니다.")
            seen.add(it["item"].upper())

        def work(conn):
            cur = conn.cursor()
            out = {"mold_exists": False}
            if code:
                cur.execute(f"SELECT COUNT(*) FROM {self.T('mold')} WHERE TRIM(MOLD_CODE) = :c", c=code)
                out["mold_exists"] = cur.fetchone()[0] > 0
            for it in self.items:
                it["level"], notes = "ok", []
                cur.execute(f"SELECT MAX(TRIM(MOLD_CODE)) FROM {self.T('link')} WHERE ITEM_CODE = :i", i=it["item"])
                other = cur.fetchone()[0]
                if other:
                    it["level"] = "err"
                    notes.append(f"이미 금형 {other}에 연결된 품번 (MES 규칙: 품번 하나 = 금형 하나)")
                cur.execute(f"SELECT COUNT(*) FROM {self.T('mst')} WHERE ITEM_CODE = :i", i=it["item"])
                if not cur.fetchone()[0]:
                    it["level"] = "err" if it["level"] == "err" else "warn"
                    notes.append("품목 마스터에 없는 품번 (오타 확인)")
                it["bom_action"] = None
                if it.get("raw"):
                    cur.execute(f"SELECT UNIT_PER_QTY FROM {self.T('bom')} WHERE ITEM_CODE = :i AND CHILD_ITEM_CODE = :r",
                                i=it["item"], r=it["raw"])
                    ex = cur.fetchall()
                    if ex:
                        it["bom_action"] = "keep"
                        notes.append(f"단위중량 기존값 {show(ex[0][0])} 유지 (바꾸려면 [단위중량 관리] 탭)")
                    elif it.get("qpu"):
                        it["bom_action"] = "insert"
                        notes.append(f"단위중량 {show(it['qpu'])} 새로 등록")
                        cur.execute(f"SELECT COUNT(*) FROM {self.T('mst')} WHERE ITEM_CODE = :r", r=it["raw"])
                        if not cur.fetchone()[0]:
                            it["level"] = "err" if it["level"] == "err" else "warn"
                            notes.append("원소재 코드가 품목 마스터에 없음 (코일 코드 확인)")
                    else:
                        it["level"] = "err" if it["level"] == "err" else "warn"
                        notes.append("단위중량이 비어 있음 → 소재 사용량이 0으로 쌓입니다")
                else:
                    it["level"] = "err" if it["level"] == "err" else "warn"
                    notes.append("원소재·단위중량 미입력 → 소재 사용량이 0으로 쌓입니다")
                it["note"] = " / ".join(notes) if notes else "OK"
            return out
        try:
            out = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return None
        if out["mold_exists"]:
            errs.append(f"금형코드 {code}는 이미 있습니다.")
        if not self.items:
            errs.append("품번이 없습니다. 품번이 없으면 작업지시를 낼 수 없습니다.")
        self.show_items()
        nerr = len(errs) + sum(1 for it in self.items if it["level"] == "err")
        nwarn = sum(1 for it in self.items if it["level"] == "warn")
        self.checked = None if nerr else code
        self.lbl.config(text=("오류 " + str(nerr) + "건" if nerr else "등록 가능") + (f" / 경고 {nwarn}건" if nwarn else ""),
                        foreground="#C00000" if nerr else "#1F6E1F")
        if errs and not quiet:
            messagebox.showwarning(APP_TITLE, "\n".join(errs))
        return nerr == 0

    # ---------- ⑤ 등록 ----------
    def build(self):
        reg = (self.app.cfg.get("reg_user") or "").strip()

        def prep(desc, base, over):
            vals, auto = {}, {}
            for d in desc:
                c, tn = d[0], str(getattr(d[1], "name", d[1])).upper()
                if any(w in tn for w in ("LOB", "LONG", "ROWID")):
                    continue
                if "DATE" in tn and (RE_AUTO_DATE.search(c) or RE_UPD_DATE.search(c)):
                    auto[c] = "SYSDATE"
                    continue
                if "CHAR" in tn and (RE_AUTO_USER.search(c) or RE_UPD_USER.search(c)) and reg:
                    auto[c] = reg
                    continue
                if c in over:
                    vals[c] = over[c]
                elif base and c in base and not any(w in c for w in RESET_WORDS):
                    vals[c] = base[c]
            return {k: v for k, v in vals.items() if v is not None}, auto

        code = self.fvars["MOLD_CODE"].get().strip()
        mcols = {d[0]: d for d in self.base["mold_desc"]}
        over = {"MOLD_CODE": code}
        for col, v in self.fvars.items():
            val = v.get().strip()
            if col in mcols and col != "MOLD_CODE" and val and val != "(이 테이블에 없음)":
                tn = str(getattr(mcols[col][1], "name", "")).upper()
                over[col] = conv({"type": "NUM" if "NUMBER" in tn else "TEXT"}, val)
        steps = [("금형", self.T("mold"), *prep(self.base["mold_desc"], self.base["mold"], over))]
        link_base = self.base.get("link") or {}
        for it in self.items:
            o = {"MOLD_CODE": code, "ITEM_CODE": it["item"]}
            if it.get("cav") is not None:
                o["CAVITY_QTY"] = it["cav"]
            if "ORGANIZATION_ID" not in link_base and "ORGANIZATION_ID" in self.base["mold"]:
                o["ORGANIZATION_ID"] = self.base["mold"]["ORGANIZATION_ID"]
            steps.append(("품번연결", self.T("link"), *prep(self.base["link_desc"], link_base, o)))
        bom_base = self.base.get("bom") or {}
        for it in self.items:
            if it.get("bom_action") == "insert":
                o = {"ITEM_CODE": it["item"], "CHILD_ITEM_CODE": it["raw"], "UNIT_PER_QTY": it["qpu"]}
                if it.get("cav") is not None:
                    o["CAVITY_QTY"] = it["cav"]
                if "ORGANIZATION_ID" not in bom_base and "ORGANIZATION_ID" in self.base["mold"]:
                    o["ORGANIZATION_ID"] = self.base["mold"]["ORGANIZATION_ID"]
                steps.append(("단위중량", self.T("bom"), *prep(self.base["bom_desc"], bom_base, o)))
        return steps

    def register(self):
        if not self.check(quiet=True) or not self.checked:
            messagebox.showwarning(APP_TITLE, "④ 검사에서 오류가 있습니다. 빨간 줄과 메시지를 확인하세요.")
            self.check()
            return
        steps = self.build()
        summary = "\n".join(f"  {kind}: " + " / ".join(show(v.get(k)) for k in ("MOLD_CODE", "ITEM_CODE", "CHILD_ITEM_CODE", "UNIT_PER_QTY") if k in v)
                            for kind, _, v, _ in steps)
        if not messagebox.askyesno("MES에 등록", f"아래를 한 번에 추가합니다 (기존 내용은 바꾸지 않음):\n\n{summary}\n\n"
                                               f"한 줄이라도 실패하면 전부 취소됩니다. 계속할까요?"):
            return
        try:
            self.app.run_db(lambda conn: self.mes.insert_many(conn, [(t, v, a) for _, t, v, a in steps]))
        except Exception as e:
            self.app.err(e)
            return
        now = dt.datetime.now()
        reg = self.app.cfg.get("reg_user", "")
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", reg, f"간편{kind}등록", t,
                              "/".join(show(v.get(k)) for k in ("MOLD_CODE", "ITEM_CODE", "CHILD_ITEM_CODE") if k in v),
                              "; ".join(f"{k}={cell_text(x)}" for k, x in v.items())] for kind, t, v, a in steps])
        self.app.load_log()
        messagebox.showinfo(APP_TITLE, f"등록 완료: 금형 {self.fvars['MOLD_CODE'].get().strip()} "
                                       f"(품번 {len(self.items)}개, 단위중량 {sum(1 for s in steps if s[0] == '단위중량')}건)\n"
                                       f"[작업지시 관리] 탭에서 이 품번으로 작업지시를 낼 수 있습니다.")
        self.reset(keep_base=True)

    def reset(self, keep_base=False):
        for col, v in self.fvars.items():
            if col in ("MOLD_CODE", "MOLD_NAME") or not keep_base:
                v.set("")
        self.items = []
        self.show_items()
        if not keep_base:
            self.base = None
            self.lbl_base.config(text="기존 금형코드나 금형명 일부를 넣고 [찾기]", foreground="#555")
        self.invalidate()


# ==========================================================
# 달력 입력칸 (추가 설치 없이 쓰는 날짜 선택)
# ==========================================================
def parse_day(text):
    """'2026-09-30', '20260930', '2026.9.30', '2026/9/30' -> date"""
    t = (text or "").strip()
    if re.fullmatch(r"\d{8}", t):
        return dt.date(int(t[:4]), int(t[4:6]), int(t[6:]))
    m = re.fullmatch(r"(\d{4})\D(\d{1,2})\D(\d{1,2})", t)
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    raise ValueError(f"날짜 형식이 틀립니다: {text}  (예: 2026-09-30)")


class DateEntry(ttk.Frame):
    """날짜 입력칸 + [달력] 버튼"""
    DAYS = ("일", "월", "화", "수", "목", "금", "토")

    def __init__(self, parent, value=None):
        super().__init__(parent)
        self.var = tk.StringVar(value=(value or dt.date.today()).strftime("%Y-%m-%d"))
        ttk.Entry(self, textvariable=self.var, width=11).pack(side="left")
        self.btn = ttk.Button(self, text="달력", width=5, command=self.popup)
        self.btn.pack(side="left", padx=(2, 0))
        self.win = None

    def get(self):
        d = parse_day(self.var.get())
        self.set(d)
        return d

    def set(self, d):
        self.var.set(d.strftime("%Y-%m-%d"))

    def popup(self):
        if self.win is not None and self.win.winfo_exists():
            self.win.destroy()
            return
        try:
            cur = parse_day(self.var.get())
        except ValueError:
            cur = dt.date.today()
        self.sel = cur
        self.ym = [cur.year, cur.month]
        win = tk.Toplevel(self)
        self.win = win
        win.title("날짜 선택")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())
        win.geometry(f"+{self.btn.winfo_rootx() - 120}+{self.btn.winfo_rooty() + self.btn.winfo_height() + 2}")
        head = tk.Frame(win, bg="#1F4E79")
        head.pack(fill="x")
        tk.Button(head, text="◀", relief="flat", bg="#1F4E79", fg="white", activebackground="#163A5C",
                  command=lambda: self.move(-1)).pack(side="left", padx=2, pady=2)
        self.lbl_ym = tk.Label(head, bg="#1F4E79", fg="white", font=("맑은 고딕", 11, "bold"))
        self.lbl_ym.pack(side="left", expand=True)
        tk.Button(head, text="▶", relief="flat", bg="#1F4E79", fg="white", activebackground="#163A5C",
                  command=lambda: self.move(1)).pack(side="right", padx=2, pady=2)
        self.body = tk.Frame(win, bg="white", padx=4, pady=4)
        self.body.pack()
        foot = tk.Frame(win, bg="white")
        foot.pack(fill="x")
        tk.Button(foot, text="오늘", relief="groove", command=lambda: self.pick(dt.date.today())).pack(
            side="left", padx=4, pady=4)
        tk.Button(foot, text="닫기", relief="groove", command=win.destroy).pack(side="right", padx=4, pady=4)
        win.bind("<Escape>", lambda e: win.destroy())
        self.draw()
        win.focus_set()

    def move(self, step):
        y, m = self.ym
        m += step
        if m == 0:
            y, m = y - 1, 12
        elif m == 13:
            y, m = y + 1, 1
        self.ym = [y, m]
        self.draw()

    def draw(self):
        import calendar
        for w in self.body.winfo_children():
            w.destroy()
        y, m = self.ym
        self.lbl_ym.config(text=f"{y}년 {m}월")
        for c, d in enumerate(self.DAYS):
            tk.Label(self.body, text=d, width=4, bg="white", font=("맑은 고딕", 9, "bold"),
                     fg="#C62828" if c == 0 else ("#1565C0" if c == 6 else "#333")).grid(row=0, column=c)
        today = dt.date.today()
        for r, week in enumerate(calendar.Calendar(firstweekday=6).monthdayscalendar(y, m), 1):
            for c, day in enumerate(week):
                if not day:
                    continue
                d = dt.date(y, m, day)
                bg, fg = "white", ("#C62828" if c == 0 else ("#1565C0" if c == 6 else "#222"))
                if d == today:
                    bg = "#FFF3C4"
                if d == self.sel:
                    bg, fg = "#1F4E79", "white"
                tk.Button(self.body, text=str(day), width=3, relief="flat", bg=bg, fg=fg,
                          activebackground="#BBDEFB", command=lambda d=d: self.pick(d)).grid(
                    row=r, column=c, padx=1, pady=1)

    def pick(self, d):
        self.set(d)
        if self.win is not None:
            self.win.destroy()
        self.event_generate("<<DatePicked>>")


# ==========================================================
# 소재 관리 탭 (IMTL_RAW_MATERIAL_RECEIPT / STOCK_LOT / STOCK)
# ==========================================================
class ReceiptTab:
    """소재 입고 조회 / 새 입고 / 고치기 / 잘못 입고한 건 삭제 (사용된 LOT은 수정·삭제 제한)"""
    T_RCV = "IMTL_RAW_MATERIAL_RECEIPT"
    T_LOT = "IMTL_RAW_MATERIAL_STOCK_LOT"
    T_STK = "IMTL_RAW_MATERIAL_STOCK"
    T_USE = "IMTL_RAW_SF_USE_CASE_DTL"
    T_ISS = "IMTL_RAW_MATERIAL_ISSUE"
    COLS = ["상태", "입고번호", "SEQ", "입고일", "품번", "품명", "HEAT_NO", "업체",
            "수량", "단위", "순중량(kg)", "총중량(kg)", "구분", "사용/출고", "등록자", "등록일시"]
    WIDTHS = [110, 110, 45, 90, 130, 180, 170, 70, 55, 45, 90, 90, 50, 80, 80, 130]

    def __init__(self, app, t):
        self.app = app
        self.rows = []
        self.unit, self.unit_note = "kg", "아직 판단 전"
        self.unit_ok = False

        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=(6, 2))
        today = dt.date.today()
        ttk.Label(bar, text="입고일").pack(side="left", padx=(4, 2))
        self.d1 = DateEntry(bar, today - dt.timedelta(days=7))
        self.d1.pack(side="left")
        ttk.Label(bar, text=" ~ ").pack(side="left")
        self.d2 = DateEntry(bar, today)
        self.d2.pack(side="left")
        for txt, a, b in (("오늘", 0, 0), ("어제", 1, 1), ("최근 7일", 7, 0), ("최근 30일", 30, 0)):
            ttk.Button(bar, text=txt, width=len(txt) + 3,
                       command=lambda a=a, b=b: self.quick(a, b)).pack(side="left", padx=(6 if txt == "오늘" else 1, 0))
        ttk.Button(bar, text="이번 달", width=7, command=lambda: self.month(0)).pack(side="left", padx=1)
        ttk.Button(bar, text="지난 달", width=7, command=lambda: self.month(-1)).pack(side="left", padx=1)

        bar2 = ttk.Frame(t)
        bar2.pack(fill="x", pady=2)
        self.vars = {}
        for key, label, w in (("item", "품번/품명", 14), ("heat", "HEAT_NO", 14),
                              ("rno", "입고번호", 12), ("vendor", "업체", 8)):
            ttk.Label(bar2, text=label).pack(side="left", padx=(6, 2))
            v = tk.StringVar()
            e = ttk.Entry(bar2, textvariable=v, width=w)
            e.pack(side="left")
            e.bind("<Return>", lambda _e: self.fetch())
            self.vars[key] = v
        self.var_cancel = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar2, text="취소건 포함", variable=self.var_cancel).pack(side="left", padx=8)
        ttk.Button(bar2, text="MES에서 조회", style="Big.TButton", command=self.fetch).pack(side="left", padx=6)
        ttk.Button(bar2, text="엑셀로 저장", command=self.export_excel).pack(side="right", padx=3)

        bar3 = ttk.Frame(t)
        bar3.pack(fill="x", pady=2)
        ttk.Button(bar3, text="+ 새 입고", style="Big.TButton", command=lambda: ReceiptForm(self)).pack(side="left", padx=3)
        ttk.Button(bar3, text="선택 입고 고치기", style="Big.TButton", command=self.edit).pack(side="left", padx=3)
        self.var_mode = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar3, text="삭제 모드", variable=self.var_mode, command=self.on_mode).pack(side="left", padx=(12, 4))
        self.btn_del = ttk.Button(bar3, text="선택 입고 삭제", style="Big.TButton", command=self.delete,
                                  state="disabled")
        self.btn_del.pack(side="right", padx=3)

        ttk.Label(t, text="초록 = 수정·삭제 가능   회색 = 생산에 사용됨(업체·담당자만 수정, 삭제 불가)   "
                          "노랑 = 취소(반품)된 입고   |   줄 더블클릭 = 고치기 창",
                  foreground="#555").pack(fill="x", padx=6)
        fr, self.tv = app.make_tree(t, self.COLS, self.WIDTHS)
        fr.pack(fill="both", expand=True, pady=4)
        self.tv.bind("<Double-1>", lambda e: self.edit())
        self.lbl = ttk.Label(t, text="조회 전", font=("", 10, "bold"))
        self.lbl.pack(fill="x", padx=6, pady=(0, 4))

    # ---------- 중량 단위 / 재고 도우미 ----------
    def detect_unit(self, conn):
        """입고 순중량(NET_WEIGHT)이 kg인지 g인지: 설정값, 없으면 최근 1년 입고 중앙값 크기로 판단"""
        if self.unit_ok:
            return
        want = str(self.app.cfg.get("rcv_weight_unit") or "auto").strip().lower()
        if want in ("kg", "g"):
            self.unit, self.unit_note = want, "환경설정 값"
        else:
            cur = conn.cursor()
            cur.execute(f"SELECT MEDIAN(NET_WEIGHT) FROM {self.T_RCV} WHERE IO_FLAG = 1 AND NET_WEIGHT > 0 "
                        f"AND RECEIPT_DATE >= SYSDATE - 365")
            med = cur.fetchone()[0]
            self.unit = "g" if med and med >= 50000 else "kg"
            self.unit_note = f"자동 판단: 최근 입고 중앙값 {med:,.0f}" if med else "자동 판단: 자료 없음 → kg"
        self.unit_ok = True

    def to_kg(self, v):
        if v is None:
            return None
        return v / 1000 if self.unit == "g" else v

    def from_kg(self, kg):
        return kg * 1000 if self.unit == "g" else kg

    def usage(self, conn, r):
        """생산 사용 / 출고 / 취소줄 건수. 하나도 없으면 {}"""
        cur = conn.cursor()
        b = {"rno": r["RECEIPT_NO"], "seq": r["SEQ"], "heat": r.get("HEAT_NO")}
        cur.execute(f"SELECT COUNT(*) FROM {self.T_USE} WHERE RECEIPT_NO = :rno AND RECEIPT_SEQ = :seq "
                    f"AND NVL(HEAT_NO, '~') = NVL(:heat, '~')", b)
        use = cur.fetchone()[0]
        cur.execute(f"SELECT COUNT(*) FROM {self.T_ISS} WHERE RECEIPT_NO = :rno AND RECEIPT_SEQ = :seq "
                    f"AND NVL(HEAT_NO, '~') = NVL(:heat, '~')", b)
        iss = cur.fetchone()[0]
        cur.execute(f"SELECT COUNT(*) FROM {self.T_RCV} WHERE RECEIPT_NO = :rno AND SEQ = :seq AND IO_FLAG = 2",
                    rno=r["RECEIPT_NO"], seq=r["SEQ"])
        cancel = cur.fetchone()[0]
        return {"use": use, "iss": iss, "cancel": cancel} if (use or iss or cancel) else {}

    def next_no(self, conn, d):
        """입고번호 제안: yymmdd + 일련번호(10001부터) - 소재입고 업로드 프로그램과 같은 방식"""
        prefix = d.strftime("%y%m%d")
        cur = conn.cursor()
        cur.execute(f"SELECT MAX(TO_CHAR(RECEIPT_NO)) FROM {self.T_RCV} WHERE TO_CHAR(RECEIPT_NO) LIKE :p",
                    p=prefix + "%")
        last = cur.fetchone()[0]
        if last and str(last)[len(prefix):].isdigit():
            return f"{prefix}{int(str(last)[len(prefix):]) + 1}", last
        return f"{prefix}10001", None

    def stock_add(self, cur, item, dq, dkg, user, org=1, unit="RL", meta=None):
        """소재 재고(STOCK)에 수량·kg 더하기(빼기는 음수). 품번 줄이 2개 이상이면 오류"""
        cur.execute(f"SELECT ROWIDTOCHAR(ROWID), QTY, STOCK_WEIGHT_KG, STOCK_WEIGHT_G FROM {self.T_STK} "
                    f"WHERE ITEM_CODE = :i", i=item)
        rows = cur.fetchall()
        if len(rows) > 1:
            raise RuntimeError(f"소재 재고({self.T_STK})에 품번 {item} 줄이 {len(rows)}개라 재고를 맞출 수 없습니다.")
        now = dt.datetime.now()
        if not rows:
            if dq <= 0 and dkg <= 0:
                return f"{item} 재고 줄 없음(빼지 않음)"
            meta = meta or table_meta(cur.connection, self.T_STK)
            insert_dict(cur, self.T_STK, meta, {
                "ITEM_CODE": item, "QTY": dq, "QTY_UNIT": unit, "STOCK_WEIGHT_KG": round(dkg, 3),
                "STOCK_WEIGHT_G": round(dkg * 1000), "LAST_RECEIPT_DATE": now, "ORGANIZATION_ID": org,
                "ENTER_DATE": now, "ENTER_BY": user, "LAST_MODIFY_DATE": now, "LAST_MODIFY_BY": user})
            return f"{item} 재고 새로 만듦 {dq}롤 / {dkg:,.1f}kg"
        rid, q0, k0, g0 = rows[0]
        q0, k0, g0 = q0 or 0, k0 or 0, g0 or 0
        q1, k1, g1 = max(q0 + dq, 0), max(round(k0 + dkg, 3), 0), max(round(g0 + dkg * 1000), 0)
        extra = ", LAST_RECEIPT_DATE = SYSDATE" if dq > 0 else ""
        cur.execute(f"UPDATE {self.T_STK} SET QTY = :q, STOCK_WEIGHT_KG = :k, STOCK_WEIGHT_G = :g, "
                    f"LAST_MODIFY_DATE = SYSDATE, LAST_MODIFY_BY = :u{extra} WHERE ROWID = CHARTOROWID(:rid)",
                    q=q1, k=k1, g=g1, u=user, rid=rid)
        if cur.rowcount != 1:
            raise RuntimeError(f"소재 재고 {item} 줄을 찾지 못했습니다.")
        return f"{item} {q0}→{q1}롤, {k0:,.1f}→{k1:,.1f}kg"

    def edit(self):
        sel = self.tv.selection()
        if len(sel) != 1:
            messagebox.showinfo(APP_TITLE, "고칠 입고 1줄을 선택하세요. (먼저 [MES에서 조회])")
            return
        r = self.rows[int(sel[0])]
        if r["IO_FLAG"] != 1:
            messagebox.showinfo(APP_TITLE, "취소(반품) 줄은 고칠 수 없습니다. 원래 입고 줄을 선택하세요.")
            return
        ReceiptForm(self, r)

    # ---------- 날짜 빠른 선택 ----------
    def quick(self, a, b):
        today = dt.date.today()
        self.d1.set(today - dt.timedelta(days=a))
        self.d2.set(today - dt.timedelta(days=b))
        self.fetch()

    def month(self, step):
        first = dt.date.today().replace(day=1)
        if step < 0:
            first = (first - dt.timedelta(days=1)).replace(day=1)
        nxt = (first + dt.timedelta(days=32)).replace(day=1)
        self.d1.set(first)
        self.d2.set(nxt - dt.timedelta(days=1))
        self.fetch()

    def on_mode(self):
        if self.var_mode.get() and not messagebox.askyesno(
                APP_TITLE, "삭제 모드를 켜면 MES 소재 입고를 지울 수 있습니다.\n"
                           "(입고·LOT를 지우고 소재 재고에서 빼 줍니다)\n\n켤까요?"):
            self.var_mode.set(False)
        self.btn_del.config(state="normal" if self.var_mode.get() else "disabled")

    # ---------- 조회 ----------
    def sql(self):
        s = f"""
            SELECT R.RECEIPT_NO, R.SEQ, R.RECEIPT_DATE, R.ITEM_CODE,
                   (SELECT MAX(M.ITEM_NAME) FROM ICOM_ITEM_MASTER M WHERE M.ITEM_CODE = R.ITEM_CODE) ITEM_NAME,
                   R.HEAT_NO, R.VENDOR_SITE_ID, R.QTY, R.QTY_UNIT, R.NET_WEIGHT, R.TOTAL_WEIGHT,
                   R.IO_FLAG, R.ENTER_BY, R.ENTER_DATE,
                   (SELECT COUNT(*) FROM {self.T_USE} U WHERE U.RECEIPT_NO = R.RECEIPT_NO
                       AND U.RECEIPT_SEQ = R.SEQ AND NVL(U.HEAT_NO, '~') = NVL(R.HEAT_NO, '~')) USE_CNT,
                   (SELECT COUNT(*) FROM {self.T_ISS} I WHERE I.RECEIPT_NO = R.RECEIPT_NO
                       AND I.RECEIPT_SEQ = R.SEQ AND NVL(I.HEAT_NO, '~') = NVL(R.HEAT_NO, '~')) ISS_CNT,
                   (SELECT COUNT(*) FROM {self.T_RCV} X WHERE X.RECEIPT_NO = R.RECEIPT_NO
                       AND X.SEQ = R.SEQ AND X.IO_FLAG = 2) CANCEL_CNT
            FROM {self.T_RCV} R
            WHERE R.RECEIPT_DATE >= :sd AND R.RECEIPT_DATE < :ed"""
        d1, d2 = self.d1.get(), self.d2.get()
        if d1 > d2:
            d1, d2 = d2, d1
            self.d1.set(d1)
            self.d2.set(d2)
        p = {"sd": dt.datetime.combine(d1, dt.time()), "ed": dt.datetime.combine(d2 + dt.timedelta(days=1), dt.time())}
        if not self.var_cancel.get():
            s += " AND R.IO_FLAG = 1"
        v = {k: x.get().strip() for k, x in self.vars.items()}
        if v["item"]:
            s += (" AND (UPPER(R.ITEM_CODE) LIKE :item OR EXISTS (SELECT 1 FROM ICOM_ITEM_MASTER M "
                  "WHERE M.ITEM_CODE = R.ITEM_CODE AND UPPER(M.ITEM_NAME) LIKE :item))")
            p["item"] = f"%{v['item'].upper()}%"
        if v["heat"]:
            s += " AND UPPER(R.HEAT_NO) LIKE :heat"
            p["heat"] = f"%{v['heat'].upper()}%"
        if v["rno"]:
            s += " AND TO_CHAR(R.RECEIPT_NO) LIKE :rno"
            p["rno"] = f"%{v['rno']}%"
        if v["vendor"]:
            s += " AND UPPER(R.VENDOR_SITE_ID) LIKE :vd"
            p["vd"] = f"%{v['vendor'].upper()}%"
        s += " ORDER BY R.RECEIPT_DATE DESC, R.RECEIPT_NO DESC, R.SEQ, R.IO_FLAG"
        return s, p

    def fetch(self):
        try:
            s, p = self.sql()
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e))
            return

        def work(conn):
            self.unit_ok = False
            self.detect_unit(conn)
            cur = conn.cursor()
            cur.execute(s, p)
            names = [d[0] for d in cur.description]
            return [dict(zip(names, r)) for r in cur.fetchall()]
        try:
            self.rows = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return
        self.show()

    @staticmethod
    def state(r):
        if (r["USE_CNT"] or 0) or (r["ISS_CNT"] or 0):
            return "사용됨(삭제불가)", "off"
        if r["IO_FLAG"] == 2:
            return "취소 줄", "warn"
        if r["CANCEL_CNT"]:
            return "취소된 입고", "warn"
        return "수정·삭제 가능", "ok"

    def show(self):
        tv = self.tv
        tv.delete(*tv.get_children())
        for i, r in enumerate(self.rows):
            st, tag = self.state(r)
            io = {1: "입고", 2: "취소"}.get(r["IO_FLAG"], show(r["IO_FLAG"]))
            use = f"{r['USE_CNT'] or 0}/{r['ISS_CNT'] or 0}"
            tv.insert("", "end", iid=str(i), tags=(tag,), values=[
                st, show(r["RECEIPT_NO"]), show(r["SEQ"]),
                r["RECEIPT_DATE"].strftime("%Y-%m-%d") if r["RECEIPT_DATE"] else "",
                show(r["ITEM_CODE"]), show(r["ITEM_NAME"]), show(r["HEAT_NO"]), show(r["VENDOR_SITE_ID"]),
                show(r["QTY"]), show(r["QTY_UNIT"]),
                f"{self.to_kg(r['NET_WEIGHT']):,.1f}" if r["NET_WEIGHT"] is not None else "",
                f"{self.to_kg(r['TOTAL_WEIGHT']):,.1f}" if r["TOTAL_WEIGHT"] is not None else "",
                io, use, show(r["ENTER_BY"]),
                r["ENTER_DATE"].strftime("%Y-%m-%d %H:%M") if r["ENTER_DATE"] else ""])
        ins = [r for r in self.rows if r["IO_FLAG"] == 1]
        kg = sum((self.to_kg(r["NET_WEIGHT"]) or 0) for r in ins)
        ok = sum(1 for r in self.rows if self.state(r)[1] == "ok")
        self.lbl.config(text=f"{len(self.rows)}건  (입고 {len(ins)}건 / 순중량 합계 {kg:,.1f} kg)   "
                             f"수정·삭제 가능 {ok}건  ·  사용됨 {sum(1 for r in self.rows if self.state(r)[1] == 'off')}건"
                             f"      (MES 순중량 단위: {self.unit} - {self.unit_note})")

    def label(self, r):
        io = "취소" if r["IO_FLAG"] == 2 else "입고"
        return (f"{show(r['RECEIPT_NO'])}-{show(r['SEQ'])} {io} {show(r['ITEM_CODE'])} "
                f"HEAT {show(r['HEAT_NO'])} {show(r['QTY'])}{show(r['QTY_UNIT'])} {self.to_kg(r['NET_WEIGHT']) or 0:,.1f}kg")

    def detail(self):
        sel = self.tv.selection()
        if not sel:
            return
        r = self.rows[int(sel[0])]
        st, _ = self.state(r)
        messagebox.showinfo(APP_TITLE, f"{self.label(r)}\n\n상태: {st}\n"
                                       f"생산 사용 {r['USE_CNT']}건 / 출고 {r['ISS_CNT']}건 / 같은 입고의 취소 줄 {r['CANCEL_CNT']}건")

    # ---------- 삭제 ----------
    def delete(self):
        if not self.var_mode.get():
            messagebox.showinfo(APP_TITLE, "[삭제 모드]를 체크한 뒤 하세요.")
            return
        sel = sorted(int(i) for i in self.tv.selection())
        if not sel:
            messagebox.showinfo(APP_TITLE, "삭제할 줄을 선택하세요. (여러 줄: Ctrl/Shift+클릭)")
            return
        targets, seen, blocked = [], set(), []
        for i in sel:
            r = self.rows[i]
            if self.state(r)[1] == "off":
                blocked.append(self.label(r))
                continue
            k = (r["RECEIPT_NO"], r["SEQ"])
            if k not in seen:
                seen.add(k)
                targets.append(r)
        if blocked:
            messagebox.showwarning(APP_TITLE, "생산에 사용된 소재라 삭제할 수 없습니다:\n\n" + "\n".join(blocked[:15])
                                   + ("\n\n나머지 줄만 선택해서 다시 하세요." if targets else ""))
            return
        lines = "\n".join(f"  {self.label(r)}" for r in targets[:20]) + \
            (f"\n  ... 외 {len(targets) - 20}건" if len(targets) > 20 else "")
        note = ("\n\n※ 같은 입고번호·SEQ의 입고줄과 취소줄은 함께 지워집니다."
                if any(r["CANCEL_CNT"] or r["IO_FLAG"] == 2 for r in targets) else "")
        typed = simpledialog.askstring(
            "입고 삭제 확인",
            f"MES에서 아래 {len(targets)}건의 소재 입고를 삭제합니다.\n"
            f"(입고 + 재고LOT 삭제, 소재 재고에서 수량·중량 차감){note}\n\n{lines}\n\n"
            f"되돌릴 수 없습니다. 진행하려면 '삭제'라고 입력하세요:", parent=self.app.root)
        if (typed or "").strip() != "삭제":
            return
        now = dt.datetime.now()
        try:
            results = self.app.run_db(lambda c: [(r, *self.delete_one(c, r, now)) for r in targets])
        except Exception as e:
            self.app.err(e)
            return
        user = self.app.cfg.get("reg_user", "")
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", user, "입고삭제", self.T_RCV,
                              f"{show(r['RECEIPT_NO'])}-{show(r['SEQ'])}", msg]
                             for r, ok, msg in results if ok])
        self.app.load_log()
        okn = sum(1 for _, ok, _ in results if ok)
        fails = [f"  {self.label(r)}\n     → {msg}" for r, ok, msg in results if not ok]
        text = f"삭제 완료 {okn}건 / 실패 {len(fails)}건\n백업: {BACKUP_PATH}"
        if fails:
            messagebox.showwarning(APP_TITLE, text + "\n\n실패:\n" + "\n".join(fails[:10]))
        else:
            messagebox.showinfo(APP_TITLE, text)
        self.fetch()

    def delete_one(self, conn, r, now):
        """입고 1건(입고번호+SEQ) 삭제. 건마다 따로 저장/취소. return (성공, 메시지)"""
        rno, seq, heat = r["RECEIPT_NO"], r["SEQ"], r["HEAT_NO"]
        self.detect_unit(conn)
        key = {"rno": rno, "seq": seq}
        keyh = {"rno": rno, "seq": seq, "heat": heat}
        cur = conn.cursor()

        def rows_of(sql, binds):
            cur.execute(sql, binds)
            names = [d[0] for d in cur.description]
            return [dict(zip(names, x)) for x in cur.fetchall()]
        try:
            # 1) 사용 여부 다시 확인 (조회 후 누가 썼을 수 있음)
            cur.execute(f"SELECT COUNT(*), NVL(SUM(USE_WEIGHT), 0) FROM {self.T_USE} WHERE RECEIPT_NO = :rno "
                        f"AND RECEIPT_SEQ = :seq AND NVL(HEAT_NO, '~') = NVL(:heat, '~')", keyh)
            ucnt, uwt = cur.fetchone()
            cur.execute(f"SELECT COUNT(*) FROM {self.T_ISS} WHERE RECEIPT_NO = :rno "
                        f"AND RECEIPT_SEQ = :seq AND NVL(HEAT_NO, '~') = NVL(:heat, '~')", keyh)
            icnt = cur.fetchone()[0]
            if ucnt or icnt:
                return False, f"사용이력 있음 (생산 {ucnt}건/{uwt:,}g, 출고 {icnt}건) - 삭제 안 함"

            # 2) 지울 내용 읽어서 백업
            rcv = rows_of(f"SELECT * FROM {self.T_RCV} WHERE RECEIPT_NO = :rno AND SEQ = :seq", key)
            if not rcv:
                return False, "이미 삭제되었거나 없는 입고입니다."
            lot = rows_of(f"SELECT * FROM {self.T_LOT} WHERE RECEIPT_NO = :rno AND RECEIPT_SEQ = :seq "
                          f"AND NVL(HEAT_NO, '~') = NVL(:heat, '~')", keyh)
            ins = [x for x in rcv if x.get("IO_FLAG") == 1]
            cancel = any(x.get("IO_FLAG") == 2 for x in rcv)
            item = (ins[0] if ins else rcv[0]).get("ITEM_CODE")
            stk = []
            if ins and not cancel and item:
                stk = rows_of(f"SELECT ROWIDTOCHAR(ROWID) RID__, S.* FROM {self.T_STK} S WHERE ITEM_CODE = :ic",
                              {"ic": item})
                if len(stk) > 1:
                    return False, f"소재 재고({self.T_STK})에 품번 {item} 줄이 {len(stk)}개라 재고 차감을 못 합니다 - 삭제 안 함"
            with open(BACKUP_PATH, "a", encoding="utf-8") as f:
                f.write(json.dumps({"time": f"{now:%Y-%m-%d %H:%M:%S}", "user": self.app.cfg.get("reg_user", ""),
                                    "table": self.T_RCV, "kind": "입고삭제",
                                    "receipt": [{k: show(v) for k, v in x.items()} for x in rcv],
                                    "stock_lot": [{k: show(v) for k, v in x.items()} for x in lot],
                                    "stock_before": [{k: show(v) for k, v in x.items() if k != "RID__"} for x in stk]},
                                   ensure_ascii=False) + "\n")

            # 3) LOT, 입고 삭제
            cur.execute(f"DELETE FROM {self.T_LOT} WHERE RECEIPT_NO = :rno AND RECEIPT_SEQ = :seq "
                        f"AND NVL(HEAT_NO, '~') = NVL(:heat, '~')", keyh)
            nlot = cur.rowcount
            cur.execute(f"DELETE FROM {self.T_RCV} WHERE RECEIPT_NO = :rno AND SEQ = :seq", key)
            nrcv = cur.rowcount
            if nrcv != len(rcv):
                conn.rollback()
                return False, f"입고 줄 수가 달라짐({len(rcv)}→{nrcv}) - 취소함"

            # 4) 재고 차감 (취소 안 된 입고만. 취소된 입고는 이미 재고에서 빠져 있음)
            smsg = "재고 차감 없음(취소된 입고)" if cancel else "재고 차감 없음"
            if stk:
                dq = sum((x.get("QTY") or 0) for x in ins)
                dkg = sum((self.to_kg(x.get("NET_WEIGHT")) or 0) for x in ins)
                smsg = "재고 " + self.stock_add(cur, item, -dq, -dkg, "PYDEL")
            conn.commit()
            return True, f"입고 {nrcv}줄 / LOT {nlot}줄 삭제, {smsg}"
        except Exception as e:
            conn.rollback()
            return False, ora_hint(e)
        finally:
            cur.close()

    # ---------- 엑셀 ----------
    def export_excel(self):
        if not self.rows:
            messagebox.showinfo(APP_TITLE, "먼저 [MES에서 조회]를 하세요.")
            return
        if openpyxl is None:
            messagebox.showerror(APP_TITLE, "pip install openpyxl 이 필요합니다.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=f"소재입고_{self.d1.var.get().replace('-', '')}_{self.d2.var.get().replace('-', '')}.xlsx")
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "소재입고"
        ws.append(self.COLS)
        for iid in self.tv.get_children():
            ws.append(list(self.tv.item(iid, "values")))
        for i, w in enumerate(self.WIDTHS, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(8, w // 7)
        try:
            wb.save(path)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"저장 실패 (파일이 열려 있으면 닫으세요)\n{e}")
            return
        open_file(path)


# ==========================================================
# 작업지시 간편 입력창 (새로 만들기 / 고치기)
# ==========================================================
class WoForm:
    """③ 작업지시 관리 탭에서 여는 입력창.
    새로 만들기: 같은 품번(없으면 같은 설비)의 최근 작업지시를 틀로 복사 → 입력한 칸만 바꿔서 INSERT (기존 줄은 안 바뀜)
    고치기: 바꾼 칸만 ROWID + 원래값 조건으로 UPDATE (그 사이 누가 바꿨으면 저장 안 함)"""
    DONE_COLS = ("ACTUAL_QTY", "PASS_QTY", "BAD_QTY", "INPUT_QTY")

    def __init__(self, g, rid=None):
        self.g, self.app, self.rid = g, g.app, rid
        self.orig = None
        try:
            self.tbl = g.table()
            self.app.run_db(self.load)
        except Exception as e:
            self.app.err(e)
            return
        if rid and self.orig is None:
            messagebox.showwarning(APP_TITLE, "선택한 작업지시를 MES에서 찾지 못했습니다. 다시 조회하세요.")
            return
        self.build()

    # ---------- 기초 자료 ----------
    def has(self, c):
        return c in self.meta

    def load(self, conn):
        owner, t = self.tbl.split(".") if "." in self.tbl else (None, self.tbl)
        cur = conn.cursor()
        cur.execute("SELECT COLUMN_NAME, DATA_TYPE, NULLABLE, DATA_DEFAULT, DATA_LENGTH FROM ALL_TAB_COLUMNS "
                    "WHERE OWNER = NVL(:o, USER) AND TABLE_NAME = :t ORDER BY COLUMN_ID", o=owner, t=t)
        self.meta = {n: {"type": ty, "notnull": nl == "N", "hasdef": str(d or "").strip().upper() not in ("", "NULL"),
                         "len": ln} for n, ty, nl, d, ln in cur.fetchall()}
        tb = self.tbl
        cur.execute(f"SELECT MACHINE_CODE, MAX(MACHINE_NAME) FROM {tb} WHERE WORK_ORDER_DATE >= SYSDATE - 365 "
                    f"AND MACHINE_CODE IS NOT NULL GROUP BY MACHINE_CODE ORDER BY MACHINE_CODE")
        self.machines = [(show(a), show(b)) for a, b in cur.fetchall()]
        self.shifts, self.st1, self.st2 = [], [], []
        if self.has("WORK_SHIFT"):
            cur.execute(f"SELECT WORK_SHIFT, COUNT(*) FROM {tb} WHERE WORK_ORDER_DATE >= SYSDATE - 180 "
                        f"GROUP BY WORK_SHIFT ORDER BY 2 DESC")
            self.shifts = [show(a) for a, _ in cur.fetchall() if a is not None]
        self.def_st = {}
        for c in ("WORK_ORDER_STATUS", "PLAN_STATUS"):
            if not self.has(c):
                continue
            cur.execute(f"SELECT {c}, COUNT(*) FROM {tb} WHERE WORK_ORDER_DATE >= SYSDATE - 180 "
                        f"GROUP BY {c} ORDER BY 2 DESC")
            vals = [show(a) for a, _ in cur.fetchall() if a is not None]
            (self.st1 if c == "WORK_ORDER_STATUS" else self.st2).extend(vals)
            done = " OR ".join(f"NVL({x}, 0) <> 0" for x in self.DONE_COLS if self.has(x)) or "1 = 0"
            cur.execute(f"SELECT {c} FROM (SELECT {c}, COUNT(*) N FROM {tb} WHERE WORK_ORDER_DATE >= SYSDATE - 90 "
                        f"AND NOT ({done}) GROUP BY {c} ORDER BY N DESC) WHERE ROWNUM = 1")
            r = cur.fetchone()
            self.def_st[c] = show(r[0]) if r and r[0] is not None else (vals[0] if vals else "")
        order = "ENTER_DATE DESC NULLS LAST, WORK_ORDER_NO DESC" if self.has("ENTER_DATE") else "WORK_ORDER_NO DESC"
        cur.execute(f"SELECT WORK_ORDER_NO FROM (SELECT WORK_ORDER_NO FROM {tb} WHERE WORK_ORDER_DATE >= SYSDATE - 30 "
                    f"ORDER BY {order}) WHERE ROWNUM = 1")
        r = cur.fetchone()
        self.last_no = show(r[0]) if r else "(없음)"
        if self.rid:
            self.orig = self.row_by_rid(conn, self.rid)

    def row_by_rid(self, conn, rid):
        cur = conn.cursor()
        cur.execute(f"SELECT * FROM {self.tbl} WHERE ROWID = CHARTOROWID(:r)", r=rid)
        r = cur.fetchone()
        return dict(zip([d[0] for d in cur.description], r)) if r else None

    # ---------- 화면 ----------
    def build(self):
        o = self.orig or {}
        new = self.rid is None
        win = tk.Toplevel(self.app.root)
        self.win = win
        win.title("작업지시 새로 만들기" if new else f"작업지시 고치기 - {show(o.get('WORK_ORDER_NO'))}")
        win.geometry("900x580")
        win.transient(self.app.root)
        bg = "#EEF4FB"
        head = tk.Frame(win, bg=bg, highlightbackground="#B7CCE4", highlightthickness=1)
        head.pack(fill="x", padx=8, pady=8)
        tk.Label(head, bg=bg, fg="#1F4E79", font=("맑은 고딕", 11, "bold"), anchor="w",
                 text="새 작업지시 만들기" if new else "작업지시 고치기").pack(fill="x", padx=10, pady=(6, 0))
        tk.Label(head, bg=bg, fg="#333", anchor="w", justify="left", font=("맑은 고딕", 9), text=(
            "품번을 넣고 Enter(또는 [품번 찾기]) → 설비·수량 확인 → [MES에 등록].\n"
            "나머지 칸은 같은 품번(없으면 같은 설비)의 최근 작업지시에서 복사합니다. 금형은 품번 연결로 자동으로 들어갑니다."
            if new else
            "바꿀 칸만 고치고 [수정 저장]. 바뀐 칸만 저장하며, 그 사이 다른 곳에서 바뀌었으면 저장하지 않습니다.\n"
            "품번을 바꾸면 금형이, 설비를 바꾸면 설비명이 MES에서 자동으로 바뀝니다.")).pack(fill="x", padx=10, pady=(0, 6))

        body = ttk.Frame(win, padding=(16, 4))
        body.pack(fill="both", expand=True)
        body.columnconfigure(2, weight=1)
        r = 0

        def row(label, widget, note=None):
            nonlocal r
            ttk.Label(body, text=label, font=("", 10, "bold")).grid(row=r, column=0, sticky="w", pady=5)
            widget.grid(row=r, column=1, sticky="w", padx=8)
            if note is not None:
                note.configure(wraplength=420, justify="left")
                note.grid(row=r, column=2, sticky="w")
            r += 1

        d0 = o.get("WORK_ORDER_DATE")
        self.d = DateEntry(body, d0.date() if d0 else dt.date.today())
        self.d.var.trace_add("write", lambda *_: self.on_date())
        row("지시일", self.d)

        if self.shifts or self.has("WORK_SHIFT"):
            self.v_shift = tk.StringVar(value=show(o.get("WORK_SHIFT")) if o else (self.shifts[0] if self.shifts else "1"))
            row("근무조", ttk.Combobox(body, textvariable=self.v_shift, values=self.shifts, width=8))
        else:
            self.v_shift = None

        itf = ttk.Frame(body)
        self.v_item = tk.StringVar(value=show(o.get("ITEM_CODE")))
        e = ttk.Entry(itf, textvariable=self.v_item, width=22)
        e.pack(side="left")
        e.bind("<Return>", lambda _e: self.lookup_item())
        e.bind("<FocusOut>", lambda _e: self.lookup_item(quiet=True))
        ttk.Button(itf, text="품번 찾기", command=self.find_item).pack(side="left", padx=4)
        row("품번", itf)
        self.lbl_item = tk.Label(body, text="", anchor="w", justify="left", font=("맑은 고딕", 9))
        self.lbl_item.grid(row=r, column=1, columnspan=2, sticky="w", padx=8)
        r += 1

        self.mc_list = [f"{a} | {b}" if b else a for a, b in self.machines]
        mc0 = show(o.get("MACHINE_CODE"))
        self.v_mc = tk.StringVar(value=next((x for x in self.mc_list if x.split(" | ")[0] == mc0), mc0))
        row("설비", ttk.Combobox(body, textvariable=self.v_mc, values=self.mc_list, width=40),
            ttk.Label(body, text="목록에서 고르거나 설비코드 입력", foreground="#777"))

        self.v_qty = tk.StringVar(value=show(o.get("PLAN_QTY")))
        row("계획수량", ttk.Entry(body, textvariable=self.v_qty, width=12))

        self.v_st = {}
        for c, lab, lst in (("WORK_ORDER_STATUS", "상태", self.st1), ("PLAN_STATUS", "계획상태", self.st2)):
            if not self.has(c):
                continue
            v = tk.StringVar(value=show(o.get(c)) if o else self.def_st.get(c, ""))
            self.v_st[c] = v
            row(lab, ttk.Combobox(body, textvariable=v, values=lst, width=12),
                ttk.Label(body, text=("보통 그대로 두세요 (최근 새 작업지시에 가장 많이 쓴 값)" if new
                                      else "바꾸면 현장 화면에서 보이거나 사라질 수 있음"), foreground="#777"))

        self.v_no = tk.StringVar(value=show(o.get("WORK_ORDER_NO")))
        self.no_auto = new
        ent_no = ttk.Entry(body, textvariable=self.v_no, width=22, state="normal" if new else "readonly")
        ent_no.bind("<Key>", lambda _e: setattr(self, "no_auto", False))
        self.lbl_no = ttk.Label(body, text="" if new else "번호는 실적과 연결되어 바꿀 수 없습니다", foreground="#777")
        row("작업지시번호", ent_no, self.lbl_no)

        if not new:
            done = {c: o.get(c) for c in self.DONE_COLS if o.get(c) not in (None, 0)}
            txt = ("실적 있음: " + ", ".join(f"{c}={show(v)}" for c, v in done.items())) if done else "실적 없음"
            ttk.Label(body, text=f"금형: {show(o.get('MOLD_CODE'))}   설비명: {show(o.get('MACHINE_NAME'))}   {txt}",
                      foreground="#C62828" if done else "#555").grid(row=r, column=0, columnspan=3, sticky="w", pady=(10, 0))
            r += 1

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=8)
        if new:
            ttk.Button(bar, text="MES에 등록", style="Big.TButton", command=self.save_new).pack(side="left", padx=3)
        else:
            ttk.Button(bar, text="수정 저장", style="Big.TButton", command=self.save_edit).pack(side="left", padx=3)
        ttk.Button(bar, text="닫기", command=win.destroy).pack(side="right", padx=3)
        if new:
            self.on_date()
        if self.v_item.get():
            self.lookup_item(quiet=True)

    # ---------- 입력 도우미 ----------
    def mc_code(self):
        return self.v_mc.get().split(" | ")[0].strip()

    def on_date(self):
        if self.rid is not None or not self.no_auto:
            return
        try:
            d = parse_day(self.d.var.get())
        except ValueError:
            return
        day = dt.datetime.combine(d, dt.time())

        def work(conn):
            cur = conn.cursor()
            cur.execute(f"SELECT MAX(WORK_ORDER_NO) FROM {self.tbl} WHERE WORK_ORDER_DATE >= :a AND WORK_ORDER_DATE < :b",
                        a=day, b=day + dt.timedelta(days=1))
            return wo_next_no(conn, self.tbl, day), cur.fetchone()[0]
        try:
            nxt, same = self.app.run_db(work)
        except Exception:
            return
        self.v_no.set(nxt or "")
        self.no_auto = True
        if same:
            self.lbl_no.config(text=f"제안 번호 (같은 날 마지막: {same})", foreground="#2E7D32")
        else:
            self.lbl_no.config(text=f"※ 이 날 첫 번호라 앞부분은 추정입니다. 최근 번호 {self.last_no} 와 형식을 비교하세요",
                               foreground="#C62828")

    def lookup_item(self, quiet=False):
        it = self.v_item.get().strip()
        if not it:
            self.lbl_item.config(text="")
            return
        tb = self.tbl
        mst = self.app.mes.qualify(self.app.cfg.get("item_master_table") or "ICOM_ITEM_MASTER")
        link = self.app.mes.qualify(self.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")

        def work(conn):
            cur = conn.cursor()
            cur.execute(f"SELECT MAX(ITEM_NAME), COUNT(*) FROM {mst} WHERE ITEM_CODE = :i", i=it)
            name, n = cur.fetchone()
            cur.execute(f"SELECT MAX(MOLD_CODE) FROM {link} WHERE ITEM_CODE = :i", i=it)
            mold = cur.fetchone()[0]
            cur.execute(f"SELECT MACHINE_CODE, PLAN_QTY FROM (SELECT MACHINE_CODE, PLAN_QTY FROM {tb} "
                        f"WHERE ITEM_CODE = :i ORDER BY WORK_ORDER_DATE DESC, WORK_ORDER_NO DESC) WHERE ROWNUM = 1", i=it)
            last = cur.fetchone()
            return name, n, mold, last
        try:
            name, n, mold, last = self.app.run_db(work)
        except Exception as e:
            if not quiet:
                self.app.err(e)
            return
        if not n:
            self.lbl_item.config(text=f"품목 마스터에 없는 품번입니다: {it}", fg="#C62828")
            return
        parts = [f"품명: {show(name)}"]
        if mold:
            parts.append(f"금형: {mold} (자동으로 들어감)")
        else:
            parts.append("금형 연결 없음 → 금형이 '*'로 들어갑니다. [① 금형·품번 관리]에서 먼저 연결하세요")
        if last:
            parts.append(f"최근 설비: {show(last[0])}")
        self.lbl_item.config(text="   ".join(parts), fg="#2E7D32" if mold else "#C62828")
        if self.rid is None and last:
            if not self.v_mc.get().strip():
                mc = show(last[0])
                self.v_mc.set(next((x for x in self.mc_list if x.split(" | ")[0] == mc), mc))
            if not self.v_qty.get().strip() and last[1] is not None:
                self.v_qty.set(show(last[1]))

    def find_item(self):
        q = self.v_item.get().strip().upper()
        mst = self.app.mes.qualify(self.app.cfg.get("item_master_table") or "ICOM_ITEM_MASTER")
        link = self.app.mes.qualify(self.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")

        def work(conn):
            cur = conn.cursor()
            cur.execute(f"SELECT * FROM (SELECT M.ITEM_CODE, MAX(M.ITEM_NAME), "
                        f"(SELECT MAX(X.MOLD_CODE) FROM {link} X WHERE X.ITEM_CODE = M.ITEM_CODE) "
                        f"FROM {mst} M WHERE UPPER(M.ITEM_CODE) LIKE :q OR UPPER(M.ITEM_NAME) LIKE :q "
                        f"GROUP BY M.ITEM_CODE ORDER BY M.ITEM_CODE) WHERE ROWNUM <= 300", q=f"%{q}%")
            return cur.fetchall()
        try:
            rows = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return
        w = tk.Toplevel(self.win)
        w.title("품번 찾기 (더블클릭 = 선택)")
        w.geometry("640x420")
        w.transient(self.win)
        fr, tv = self.app.make_tree(w, ["품번", "품명", "연결 금형"], [170, 300, 140])
        fr.pack(fill="both", expand=True, padx=6, pady=6)
        for i, (a, b, m) in enumerate(rows):
            tv.insert("", "end", iid=str(i), values=[show(a), show(b), show(m) or "(없음)"],
                      tags=() if m else ("warn",))

        def pick(_e=None):
            s = tv.selection()
            if s:
                self.v_item.set(show(rows[int(s[0])][0]))
                w.destroy()
                self.lookup_item()
        tv.bind("<Double-1>", pick)
        tv.bind("<Return>", pick)
        ttk.Label(w, text=f"{len(rows)}건 (최대 300)   노랑 = 금형 연결 없음", foreground="#555").pack(anchor="w", padx=6)

    # ---------- 값 모으기 ----------
    def cast(self, c, text):
        m = self.meta.get(c, {})
        t = (text or "").strip()
        if t == "":
            return None
        if m.get("type") == "NUMBER":
            try:
                f = float(t.replace(",", ""))
            except ValueError:
                raise ValueError(f"{c}: 숫자를 넣으세요 ({t})")
            return int(f) if f.is_integer() else f
        if m.get("type") == "DATE":
            return dt.datetime.combine(parse_day(t), dt.time())
        if m.get("len") and blen(t, int(self.app.cfg.get("kor_bytes") or 3)) > int(m["len"]):
            raise ValueError(f"{c}: 너무 깁니다 (최대 {m['len']}byte)")
        return t

    def inputs(self):
        """화면 값 -> {컬럼: 값}"""
        day = self.d.get()
        od = (self.orig or {}).get("WORK_ORDER_DATE")
        v = {"WORK_ORDER_DATE": od if od and od.date() == day else dt.datetime.combine(day, dt.time()),
             "ITEM_CODE": self.cast("ITEM_CODE", self.v_item.get()),
             "MACHINE_CODE": self.cast("MACHINE_CODE", self.mc_code()),
             "PLAN_QTY": self.cast("PLAN_QTY", self.v_qty.get())}
        if self.v_shift is not None and self.has("WORK_SHIFT"):
            v["WORK_SHIFT"] = self.cast("WORK_SHIFT", self.v_shift.get())
        for c, var in self.v_st.items():
            v[c] = self.cast(c, var.get())
        for c, lab in (("ITEM_CODE", "품번"), ("MACHINE_CODE", "설비"), ("PLAN_QTY", "계획수량")):
            if v[c] is None:
                raise ValueError(f"{lab}을(를) 넣으세요.")
        if v["PLAN_QTY"] <= 0:
            raise ValueError("계획수량은 1 이상이어야 합니다.")
        return v

    def auto_cols(self):
        reg = (self.app.cfg.get("reg_user") or "").strip() or "PYWO"
        auto = {}
        for c, m in self.meta.items():
            if m["type"] == "DATE" and (RE_AUTO_DATE.search(c) or RE_UPD_DATE.search(c)):
                auto[c] = "SYSDATE"
            elif m["type"].startswith("VARCHAR") and (RE_AUTO_USER.search(c) or RE_UPD_USER.search(c)):
                auto[c] = reg
        return auto

    def check_common(self, conn, v, own_rid=None):
        """경고 모음 (중복 / 금형 없음). return [문장]"""
        cur = conn.cursor()
        warn = []
        link = self.app.mes.qualify(self.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")
        mst = self.app.mes.qualify(self.app.cfg.get("item_master_table") or "ICOM_ITEM_MASTER")
        cur.execute(f"SELECT COUNT(*) FROM {mst} WHERE ITEM_CODE = :i", i=v["ITEM_CODE"])
        if not cur.fetchone()[0]:
            raise ValueError(f"품목 마스터에 없는 품번입니다: {v['ITEM_CODE']}")
        cur.execute(f"SELECT MAX(MOLD_CODE) FROM {link} WHERE ITEM_CODE = :i", i=v["ITEM_CODE"])
        if not cur.fetchone()[0]:
            warn.append(f"품번 {v['ITEM_CODE']}에 연결된 금형이 없어 금형이 '*'로 들어갑니다.")
        d = v["WORK_ORDER_DATE"]
        b = {"a": d, "b": d + dt.timedelta(days=1), "m": v["MACHINE_CODE"], "i": v["ITEM_CODE"]}
        own = ""
        if own_rid:
            own, b["r"] = " AND ROWID <> CHARTOROWID(:r)", own_rid
        cur.execute(f"SELECT WORK_ORDER_NO FROM {self.tbl} WHERE WORK_ORDER_DATE >= :a AND WORK_ORDER_DATE < :b "
                    f"AND MACHINE_CODE = :m AND ITEM_CODE = :i{own}", b)
        same = [show(x[0]) for x in cur.fetchall()]
        if same:
            warn.append(f"같은 날·설비·품번 작업지시가 이미 있습니다: {', '.join(same[:5])}")
        return warn

    # ---------- 새로 만들기 ----------
    def template(self, conn, v):
        cur = conn.cursor()
        for cond, b in (("ITEM_CODE = :x", v["ITEM_CODE"]), ("MACHINE_CODE = :x", v["MACHINE_CODE"]), ("1 = 1", None)):
            sql = (f"SELECT * FROM (SELECT * FROM {self.tbl} WHERE {cond} "
                   f"ORDER BY WORK_ORDER_DATE DESC, WORK_ORDER_NO DESC) WHERE ROWNUM = 1")
            cur.execute(sql, {"x": b} if b is not None else {})
            r = cur.fetchone()
            if r:
                return dict(zip([d[0] for d in cur.description], r)), cond.split(" ")[0]
        raise RuntimeError("틀로 쓸 작업지시가 하나도 없습니다.")

    def save_new(self):
        try:
            v = self.inputs()
            no = self.cast("WORK_ORDER_NO", self.v_no.get())
            if not no:
                raise ValueError("작업지시번호를 넣으세요.")
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return

        def prep(conn):
            cur = conn.cursor()
            cur.execute(f"SELECT COUNT(*) FROM {self.tbl} WHERE WORK_ORDER_NO = :n", n=no)
            if cur.fetchone()[0]:
                raise ValueError(f"작업지시번호 {no}는 이미 있습니다. 번호를 바꾸세요.")
            return self.check_common(conn, v), self.template(conn, v)
        try:
            warn, (tpl, how) = self.app.run_db(prep)
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return
        except Exception as e:
            self.app.err(e)
            return

        vals = dict(tpl)
        for c in WO_RESET + ("WIP_ENTITY_ID",):
            if c in vals:
                vals[c] = None
        old_no, old_item, old_day = tpl.get("WORK_ORDER_NO"), tpl.get("ITEM_CODE"), tpl.get("WORK_ORDER_DATE")
        vals.update(v)
        vals["WORK_ORDER_NO"] = no
        if "MODEL_NAME" in vals and vals.get("MODEL_NAME") == old_item:
            vals["MODEL_NAME"] = v["ITEM_CODE"]
        if "SET_WORK_ORDER" in vals and vals.get("SET_WORK_ORDER") is not None:
            if vals["SET_WORK_ORDER"] == old_no:
                vals["SET_WORK_ORDER"] = no
            elif not self.meta.get("SET_WORK_ORDER", {}).get("notnull"):
                vals["SET_WORK_ORDER"] = None
        if old_day:                                   # 시작·종료 시각은 틀과 같은 간격으로 옮김
            delta = v["WORK_ORDER_DATE"] - dt.datetime.combine(old_day.date(), dt.time())
            for c in ("START_DATE", "END_DATE"):
                if isinstance(vals.get(c), dt.datetime):
                    vals[c] = vals[c] + delta
        auto = self.auto_cols()
        for c in auto:
            vals.pop(c, None)
        ins = {c: x for c, x in vals.items() if x is not None and c in self.meta}
        miss = [c for c, m in self.meta.items() if m["notnull"] and not m["hasdef"] and c not in ins and c not in auto]
        if miss:
            messagebox.showerror(APP_TITLE, "반드시 필요한 칸이 비었습니다:\n  " + ", ".join(miss), parent=self.win)
            return
        src = {"ITEM_CODE": "같은 품번", "MACHINE_CODE": "같은 설비", "1": "가장 최근"}[how]
        msg = (f"MES에 작업지시 1건을 새로 추가합니다. (기존 작업지시는 바뀌지 않음)\n\n"
               f"  번호: {no}\n  지시일: {v['WORK_ORDER_DATE']:%Y-%m-%d}   근무조: {show(v.get('WORK_SHIFT'))}\n"
               f"  설비: {v['MACHINE_CODE']}\n  품번: {v['ITEM_CODE']}\n  계획수량: {v['PLAN_QTY']:,}\n"
               + "".join(f"  {c}: {show(v[c])}\n" for c in self.v_st)
               + f"\n나머지 칸은 {src} 최근 작업지시({show(old_no)})에서 복사했습니다.")
        if warn:
            msg += "\n\n※ 확인하세요:\n  " + "\n  ".join(warn)
        if not messagebox.askyesno("작업지시 등록", msg + "\n\n등록할까요?", parent=self.win,
                                   icon="warning" if warn else "question"):
            return

        def do(conn):
            self.app.mes.insert_row(conn, self.tbl, ins, auto, wo_derive)
            cur = conn.cursor()
            cur.execute(f"SELECT MOLD_CODE, MACHINE_NAME, ITEM_NAME FROM {self.tbl} WHERE WORK_ORDER_NO = :n", n=no)
            return cur.fetchone()
        try:
            got = self.app.run_db(do)
        except Exception as e:
            messagebox.showerror(APP_TITLE, ora_hint(e), parent=self.win)
            return
        now = dt.datetime.now()
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", auto.get("ENTER_BY", ""), "작업지시등록", self.tbl, no,
                              "; ".join(f"{c}={cell_text(x)}" for c, x in ins.items())]])
        self.app.load_log()
        self.win.destroy()
        messagebox.showinfo(APP_TITLE, f"작업지시를 등록했습니다: {no}\n\n"
                                       f"금형: {show(got[0]) if got else '?'}   설비명: {show(got[1]) if got else '?'}\n"
                                       f"품명: {show(got[2]) if got else '?'}\n\n현장 화면(생산현황판)에 보이는지 확인하세요.")
        g = self.g
        g.fvars["d1"].set(min(v["WORK_ORDER_DATE"].date(), parse_day(g.fvars["d1"].get() or "2000-01-01"))
                          .strftime("%Y-%m-%d"))
        if parse_day(g.fvars["d2"].get() or "2000-01-01") < v["WORK_ORDER_DATE"].date():
            g.fvars["d2"].set(v["WORK_ORDER_DATE"].strftime("%Y-%m-%d"))
        g.fetch(ask=False)

    # ---------- 고치기 ----------
    def save_edit(self):
        o = self.orig
        try:
            v = self.inputs()
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return
        ch = {c: (o.get(c), x) for c, x in v.items() if cell_text(o.get(c)) != cell_text(x)}
        if not ch:
            messagebox.showinfo(APP_TITLE, "바뀐 내용이 없습니다.", parent=self.win)
            return
        if "WORK_ORDER_DATE" in ch:
            delta = v["WORK_ORDER_DATE"] - dt.datetime.combine(o["WORK_ORDER_DATE"].date(), dt.time())
            for c in ("START_DATE", "END_DATE"):
                if isinstance(o.get(c), dt.datetime):
                    ch[c] = (o[c], o[c] + delta)
        extra_warn = []
        if "ITEM_CODE" in ch and "MODEL_NAME" in o and o.get("MODEL_NAME") == o.get("ITEM_CODE"):
            ch["MODEL_NAME"] = (o["MODEL_NAME"], v["ITEM_CODE"])
        done = {c: o.get(c) for c in self.DONE_COLS if o.get(c) not in (None, 0)}
        if done and set(ch) & {"ITEM_CODE", "MACHINE_CODE", "WORK_ORDER_DATE"}:
            extra_warn.append("이 작업지시는 실적이 있습니다 (" + ", ".join(f"{c}={show(x)}" for c, x in done.items()) +
                              "). 품번·설비·날짜를 바꾸면 실적 집계가 달라집니다.")

        def prep(conn):
            w = self.check_common(conn, v, own_rid=self.rid) if set(ch) & {"ITEM_CODE", "MACHINE_CODE",
                                                                           "WORK_ORDER_DATE"} else []
            der = wo_derive(conn, {**o, **v}) if "ITEM_CODE" in ch else {}
            return w, der
        try:
            warn, der = self.app.run_db(prep)
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return
        except Exception as e:
            self.app.err(e)
            return
        for c in ("ITEM_NAME", "ITEM_SPEC"):
            if c in der and c in o and cell_text(der[c]) != cell_text(o.get(c)):
                ch[c] = (o.get(c), der[c])
        names = GRID_SPECS["wo"]["names"]
        lines = "\n".join(f"  {names.get(c, c)}: {cell_text(a) or '(빈칸)'} → {cell_text(b) or '(빈칸)'}"
                          for c, (a, b) in ch.items())
        msg = f"작업지시 {show(o.get('WORK_ORDER_NO'))}를 아래처럼 고칩니다.\n\n{lines}"
        allw = extra_warn + warn
        if allw:
            msg += "\n\n※ 확인하세요:\n  " + "\n  ".join(allw)
        if not messagebox.askyesno("작업지시 수정", msg + "\n\n저장할까요?", parent=self.win,
                                   icon="warning" if allw else "question"):
            return
        auto = {c: a for c, a in self.auto_cols().items() if RE_UPD_DATE.search(c) or RE_UPD_USER.search(c)}

        def do(conn):
            sets, where, b = [], ["ROWID = CHARTOROWID(:rid)"], {"rid": self.rid}
            for i, (c, (a, x)) in enumerate(ch.items()):
                sets.append(f"{ident(c, '컬럼')} = :n{i}")
                b[f"n{i}"] = x
                if a is None:
                    where.append(f"{ident(c, '컬럼')} IS NULL")
                else:
                    where.append(f"{ident(c, '컬럼')} = :o{i}")
                    b[f"o{i}"] = a
            for j, (c, how) in enumerate(auto.items()):
                if c in ch:
                    continue
                if how == "SYSDATE":
                    sets.append(f"{ident(c, '컬럼')} = SYSDATE")
                else:
                    sets.append(f"{ident(c, '컬럼')} = :u{j}")
                    b[f"u{j}"] = how
            cur = conn.cursor()
            try:
                cur.execute(f"UPDATE {self.tbl} SET {', '.join(sets)} WHERE {' AND '.join(where)}", b)
                if cur.rowcount != 1:
                    raise RuntimeError("그 사이 다른 곳(MES 화면·현장)에서 바뀌었거나 없어진 작업지시입니다.\n"
                                       "다시 조회한 뒤 고치세요. (저장 안 함)")
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            return self.row_by_rid(conn, self.rid)
        try:
            new = self.app.run_db(do)
        except Exception as e:
            messagebox.showerror(APP_TITLE, ora_hint(e), parent=self.win)
            return
        now = dt.datetime.now()
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", self.app.cfg.get("reg_user", ""), "작업지시수정", self.tbl,
                              show(o.get("WORK_ORDER_NO")),
                              "; ".join(f"{c}: {cell_text(a)} -> {cell_text(x)}" for c, (a, x) in ch.items())]])
        self.app.load_log()
        self.win.destroy()
        info = f"저장했습니다: {show(o.get('WORK_ORDER_NO'))}"
        if new:
            info += f"\n\n금형: {show(new.get('MOLD_CODE'))}   설비명: {show(new.get('MACHINE_NAME'))}"
        messagebox.showinfo(APP_TITLE, info)
        self.g.fetch(ask=False)


def wo_open_new(g):
    WoForm(g, None)


def wo_open_edit(g, idx=None):
    if idx is None:
        sel = g.tv.selection()
        if len(sel) != 1:
            messagebox.showinfo(APP_TITLE, "고칠 작업지시 1줄을 선택하세요. (먼저 [MES에서 조회])")
            return
        idx = int(sel[0])
    if g.edits:
        messagebox.showinfo(APP_TITLE, "표에서 고친 뒤 저장 안 한 내용이 있습니다. 먼저 [수정 저장] 또는 [수정 취소]를 하세요.")
        return
    WoForm(g, g.rids[idx])


# ==========================================================
# 소재 입고 입력창 (새 입고 / 고치기)
# ==========================================================
def table_meta(conn, tbl):
    """ALL_TAB_COLUMNS -> {컬럼: {type, notnull, hasdef, len}}"""
    owner, t = tbl.split(".") if "." in tbl else (None, tbl)
    cur = conn.cursor()
    cur.execute("SELECT COLUMN_NAME, DATA_TYPE, NULLABLE, DATA_DEFAULT, DATA_LENGTH FROM ALL_TAB_COLUMNS "
                "WHERE OWNER = NVL(:o, USER) AND TABLE_NAME = :t ORDER BY COLUMN_ID", o=owner, t=t)
    return {n: {"type": ty, "notnull": nl == "N", "hasdef": str(d or "").strip().upper() not in ("", "NULL"),
                "len": ln} for n, ty, nl, d, ln in cur.fetchall()}


def fit_type(meta, col, v):
    """컬럼 형식에 맞게 값 바꾸기 (숫자 칸에 '1' 등)"""
    if v is None:
        return None
    t = meta.get(col, {}).get("type", "")
    if t == "NUMBER" and isinstance(v, str):
        f = float(v)
        return int(f) if f.is_integer() else f
    if t.startswith(("VARCHAR", "CHAR", "NVARCHAR")) and not isinstance(v, str):
        return show(v)
    return v


def insert_dict(cur, tbl, meta, vals):
    """meta에 있는 컬럼만 INSERT. 빠진 필수칸이 있으면 오류"""
    vals = {c: fit_type(meta, c, v) for c, v in vals.items() if c in meta and v is not None}
    miss = [c for c, m in meta.items() if m["notnull"] and not m["hasdef"] and c not in vals]
    if miss:
        raise RuntimeError(f"{tbl}: 반드시 필요한 칸이 비었습니다 - {', '.join(miss)}")
    cols = list(vals)
    cur.execute(f"INSERT INTO {tbl} ({', '.join(ident(c, '컬럼') for c in cols)}) "
                f"VALUES ({', '.join(':' + str(i) for i in range(1, len(cols) + 1))})", [vals[c] for c in cols])
    if cur.rowcount != 1:
        raise RuntimeError(f"{tbl}: 등록되지 않았습니다.")
    return vals


class ReceiptForm:
    """④ 소재 관리 탭의 입력창.
    새 입고: 입고(RECEIPT) + 재고LOT(STOCK_LOT) 추가 + 소재 재고(STOCK) 더하기 (한 묶음, 실패하면 전부 취소)
    고치기: 입고·LOT를 바뀐 칸만 고치고, 품번·수량·중량이 바뀌면 재고도 차이만큼 맞춤"""
    FIELDS_LOCK_USED = ("RECEIPT_DATE", "ITEM_CODE", "HEAT_NO", "QTY", "NET_WEIGHT")

    def __init__(self, tab, row=None):
        self.tab, self.app = tab, tab.app
        self.row = row
        self.orig = None
        try:
            self.app.run_db(self.load)
        except Exception as e:
            self.app.err(e)
            return
        if row is not None and self.orig is None:
            messagebox.showwarning(APP_TITLE, "선택한 입고(입고 줄)를 MES에서 찾지 못했습니다. 다시 조회하세요.")
            return
        self.build()

    # ---------- 기초 자료 ----------
    def load(self, conn):
        t = self.tab
        t.detect_unit(conn)
        self.m_rcv, self.m_lot, self.m_stk = (table_meta(conn, x) for x in (t.T_RCV, t.T_LOT, t.T_STK))
        cur = conn.cursor()
        cur.execute(f"SELECT ITEM_CODE, COUNT(*) FROM {t.T_RCV} WHERE RECEIPT_DATE >= SYSDATE - 365 AND IO_FLAG = 1 "
                    f"GROUP BY ITEM_CODE ORDER BY 2 DESC")
        self.items = [show(a) for a, _ in cur.fetchall() if a]
        self.vendors = []
        try:
            cur.execute("SELECT VENDOR_SITE_ID, MAX(VENDOR_NAME) FROM ICOM_VENDORS GROUP BY VENDOR_SITE_ID "
                        "ORDER BY MAX(VENDOR_NAME)")
            self.vendors = [(show(a), show(b)) for a, b in cur.fetchall() if a is not None]
        except Exception:
            pass
        if not self.vendors:
            cur.execute(f"SELECT DISTINCT VENDOR_SITE_ID FROM {t.T_RCV} WHERE VENDOR_SITE_ID IS NOT NULL "
                        f"AND RECEIPT_DATE >= SYSDATE - 365")
            self.vendors = [(show(a), "") for (a,) in cur.fetchall()]
        order = "ENTER_DATE DESC NULLS LAST" if "ENTER_DATE" in self.m_rcv else "RECEIPT_NO DESC"
        cur.execute(f"SELECT RECEIPT_NO FROM (SELECT RECEIPT_NO FROM {t.T_RCV} ORDER BY {order}) WHERE ROWNUM = 1")
        r = cur.fetchone()
        self.last_no = show(r[0]) if r else "(없음)"
        if self.row is not None:
            cur.execute(f"SELECT ROWIDTOCHAR(ROWID) RID__, R.* FROM {t.T_RCV} R "
                        f"WHERE RECEIPT_NO = :rno AND SEQ = :seq AND IO_FLAG = 1",
                        rno=self.row["RECEIPT_NO"], seq=self.row["SEQ"])
            rows = cur.fetchall()
            if len(rows) == 1:
                self.orig = dict(zip([d[0] for d in cur.description], rows[0]))
                self.used = t.usage(conn, self.orig)

    def vlabel(self, vid):
        return next((f"{a} | {b}" if b else a for a, b in self.vendors if a == vid), vid)

    # ---------- 화면 ----------
    def build(self):
        o = self.orig or {}
        new = self.orig is None
        t = self.tab
        win = tk.Toplevel(self.app.root)
        self.win = win
        win.title("소재 새 입고" if new else f"소재 입고 고치기 - {show(o.get('RECEIPT_NO'))}")
        win.geometry("900x600")
        win.transient(self.app.root)
        bg = "#EEF4FB"
        head = tk.Frame(win, bg=bg, highlightbackground="#B7CCE4", highlightthickness=1)
        head.pack(fill="x", padx=8, pady=8)
        tk.Label(head, bg=bg, fg="#1F4E79", font=("맑은 고딕", 11, "bold"), anchor="w",
                 text="소재(코일) 새 입고" if new else "소재 입고 고치기").pack(fill="x", padx=10, pady=(6, 0))
        tk.Label(head, bg=bg, fg="#333", anchor="w", justify="left", font=("맑은 고딕", 9), text=(
            "입고일 → 품번 → HEAT_NO(소재로트) → 업체 → 중량(kg) → [MES에 등록].\n"
            "입고 + 재고LOT가 추가되고 소재 재고에 수량·중량이 더해집니다. 코일 여러 개는 '창 유지'를 켜고 계속 입력하세요."
            if new else
            "바꿀 칸만 고치고 [수정 저장]. 입고와 재고LOT를 같이 고치고, 품번·수량·중량이 바뀌면 소재 재고도 차이만큼 맞춥니다.\n"
            "생산에 사용·출고된 소재는 업체·담당자만 고칠 수 있습니다.")).pack(fill="x", padx=10, pady=(0, 6))

        body = ttk.Frame(win, padding=(16, 4))
        body.pack(fill="both", expand=True)
        body.columnconfigure(2, weight=1)
        self.r = 0
        lock = (not new) and bool(self.used)
        st = "disabled" if lock else "normal"

        def row(label, widget, note=None):
            ttk.Label(body, text=label, font=("", 10, "bold")).grid(row=self.r, column=0, sticky="w", pady=5)
            widget.grid(row=self.r, column=1, sticky="w", padx=8)
            if note is not None:
                note.configure(wraplength=320, justify="left")
                note.grid(row=self.r, column=2, sticky="w")
            self.r += 1

        d0 = o.get("RECEIPT_DATE")
        self.d = DateEntry(body, d0.date() if d0 else dt.date.today())
        if lock:
            for w in self.d.winfo_children():
                w.configure(state="disabled")
        row("입고일", self.d)

        itf = ttk.Frame(body)
        self.v_item = tk.StringVar(value=show(o.get("ITEM_CODE")))
        cb = ttk.Combobox(itf, textvariable=self.v_item, values=self.items, width=26, state=st)
        cb.pack(side="left")
        cb.bind("<Return>", lambda _e: self.lookup_item())
        cb.bind("<<ComboboxSelected>>", lambda _e: self.lookup_item())
        cb.bind("<FocusOut>", lambda _e: self.lookup_item())
        ttk.Button(itf, text="품번 찾기", command=self.find_item, state=st).pack(side="left", padx=4)
        row("품번(소재)", itf, ttk.Label(body, text="목록 = 최근 1년 입고한 소재 (많이 쓴 순)", foreground="#777"))
        self.lbl_item = tk.Label(body, text="", anchor="w", justify="left", font=("맑은 고딕", 9))
        self.lbl_item.grid(row=self.r, column=1, columnspan=2, sticky="w", padx=8)
        self.r += 1

        self.v_heat = tk.StringVar(value=show(o.get("HEAT_NO")))
        row("HEAT_NO(로트)", ttk.Entry(body, textvariable=self.v_heat, width=28, state=st))

        self.v_vendor = tk.StringVar(value=self.vlabel(show(o.get("VENDOR_SITE_ID"))))
        row("업체", ttk.Combobox(body, textvariable=self.v_vendor, width=40,
                                values=[f"{a} | {b}" if b else a for a, b in self.vendors]))

        kg0 = t.to_kg(o.get("NET_WEIGHT")) if o else None
        self.v_kg = tk.StringVar(value="" if kg0 is None else f"{kg0:g}")
        row("중량(kg)", ttk.Entry(body, textvariable=self.v_kg, width=12, state=st),
            ttk.Label(body, text=f"MES 저장 단위: {t.unit} ({t.unit_note})", foreground="#777"))

        self.v_qty = tk.StringVar(value=show(o.get("QTY")) if o else "1")
        row("수량(롤)", ttk.Entry(body, textvariable=self.v_qty, width=8, state=st))

        self.v_chg = tk.StringVar(value=show(o.get("CHARGER")) if o else
                                  ((self.app.cfg.get("reg_user") or "").strip() or "ADMIN"))
        row("담당자", ttk.Entry(body, textvariable=self.v_chg, width=16))

        self.v_no = tk.StringVar(value=show(o.get("RECEIPT_NO")))
        ent = ttk.Entry(body, textvariable=self.v_no, width=20, state="normal" if new else "readonly")
        self.lbl_no = ttk.Label(body, text="" if new else "입고번호는 바꿀 수 없습니다", foreground="#777")
        row("입고번호", ent, self.lbl_no)
        if new:
            self.no_auto = True
            ent.bind("<Key>", lambda _e: setattr(self, "no_auto", False))
            self.d.var.trace_add("write", lambda *_: self.suggest_no())
            self.suggest_no()

        if not new:
            u = self.used
            txt = (f"생산 사용 {u['use']}건 / 출고 {u['iss']}건 / 취소줄 {u['cancel']}건 → 업체·담당자만 고칠 수 있습니다"
                   if u else "사용 이력 없음 - 모든 칸을 고칠 수 있습니다")
            ttk.Label(body, text=txt, foreground="#C62828" if u else "#2E7D32").grid(
                row=self.r, column=0, columnspan=3, sticky="w", pady=(10, 0))
            self.r += 1

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=8)
        if new:
            ttk.Button(bar, text="MES에 등록", style="Big.TButton", command=self.save_new).pack(side="left", padx=3)
            self.v_keep = tk.BooleanVar(value=False)
            ttk.Checkbutton(bar, text="등록 후 창 유지 (다음 코일 계속 입력)", variable=self.v_keep).pack(side="left", padx=10)
        else:
            ttk.Button(bar, text="수정 저장", style="Big.TButton", command=self.save_edit).pack(side="left", padx=3)
        ttk.Button(bar, text="닫기", command=win.destroy).pack(side="right", padx=3)
        if self.v_item.get():
            self.lookup_item()

    # ---------- 입력 도우미 ----------
    def suggest_no(self):
        if not self.no_auto:
            return
        try:
            d = parse_day(self.d.var.get())
        except ValueError:
            return
        try:
            no, same = self.app.run_db(lambda c: self.tab.next_no(c, d))
        except Exception:
            return
        self.v_no.set(no)
        self.no_auto = True
        if same:
            self.lbl_no.config(text=f"제안 번호 (같은 날 마지막: {same})", foreground="#2E7D32")
        else:
            self.lbl_no.config(text=f"이 날 첫 입고 - 최근 번호 {self.last_no} 와 형식을 비교하세요", foreground="#C62828")

    def lookup_item(self):
        it = self.v_item.get().strip()
        if not it:
            self.lbl_item.config(text="")
            return
        t = self.tab

        def work(conn):
            cur = conn.cursor()
            cur.execute("SELECT MAX(ITEM_NAME), COUNT(*) FROM ICOM_ITEM_MASTER WHERE ITEM_CODE = :i", i=it)
            name, n = cur.fetchone()
            cur.execute(f"SELECT VENDOR_SITE_ID, NET_WEIGHT FROM (SELECT VENDOR_SITE_ID, NET_WEIGHT FROM {t.T_RCV} "
                        f"WHERE ITEM_CODE = :i AND IO_FLAG = 1 ORDER BY RECEIPT_DATE DESC, RECEIPT_NO DESC) "
                        f"WHERE ROWNUM = 1", i=it)
            last = cur.fetchone()
            cur.execute(f"SELECT NVL(SUM(QTY), 0), NVL(SUM(STOCK_WEIGHT_KG), 0) FROM {t.T_STK} WHERE ITEM_CODE = :i",
                        i=it)
            stk = cur.fetchone()
            return name, n, last, stk
        try:
            name, n, last, stk = self.app.run_db(work)
        except Exception:
            return
        if not n:
            self.lbl_item.config(text=f"품목 마스터에 없는 품번입니다: {it}  (ERP 품목 등록 확인)", fg="#C62828")
            return
        parts = [f"품명: {show(name)}", f"현재 재고: {show(stk[0])}롤 / {float(stk[1] or 0):,.1f}kg"]
        if last:
            parts.append(f"최근 업체: {show(last[0])}")
        self.lbl_item.config(text="   ".join(parts), fg="#2E7D32")
        if self.orig is None and last and not self.v_vendor.get().strip() and last[0] is not None:
            self.v_vendor.set(self.vlabel(show(last[0])))

    def find_item(self):
        q = self.v_item.get().strip().upper()

        def work(conn):
            cur = conn.cursor()
            cur.execute("SELECT * FROM (SELECT ITEM_CODE, MAX(ITEM_NAME) FROM ICOM_ITEM_MASTER "
                        "WHERE UPPER(ITEM_CODE) LIKE :q OR UPPER(ITEM_NAME) LIKE :q "
                        "GROUP BY ITEM_CODE ORDER BY ITEM_CODE) WHERE ROWNUM <= 300", q=f"%{q}%")
            return cur.fetchall()
        try:
            rows = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return
        w = tk.Toplevel(self.win)
        w.title("품번 찾기 (더블클릭 = 선택)")
        w.geometry("560x420")
        w.transient(self.win)
        fr, tv = self.app.make_tree(w, ["품번", "품명"], [220, 300])
        fr.pack(fill="both", expand=True, padx=6, pady=6)
        for i, (a, b) in enumerate(rows):
            tv.insert("", "end", iid=str(i), values=[show(a), show(b)])

        def pick(_e=None):
            s = tv.selection()
            if s:
                self.v_item.set(show(rows[int(s[0])][0]))
                w.destroy()
                self.lookup_item()
        tv.bind("<Double-1>", pick)
        tv.bind("<Return>", pick)
        ttk.Label(w, text=f"{len(rows)}건 (최대 300)", foreground="#555").pack(anchor="w", padx=6)

    def inputs(self):
        t = self.tab
        d = self.d.get()
        od = (self.orig or {}).get("RECEIPT_DATE")
        v = {"RECEIPT_DATE": od if od and od.date() == d else dt.datetime.combine(d, dt.time()),
             "ITEM_CODE": self.v_item.get().strip(), "HEAT_NO": self.v_heat.get().strip() or None,
             "VENDOR_SITE_ID": self.v_vendor.get().split(" | ")[0].strip() or None,
             "CHARGER": self.v_chg.get().strip() or None}
        if not v["ITEM_CODE"]:
            raise ValueError("품번을 넣으세요.")
        if not v["HEAT_NO"]:
            raise ValueError("HEAT_NO(소재로트)를 넣으세요.")
        try:
            kg = float(self.v_kg.get().replace(",", ""))
            qty = float(self.v_qty.get().replace(",", ""))
        except ValueError:
            raise ValueError("중량(kg)과 수량은 숫자로 넣으세요.")
        if kg <= 0 or qty <= 0:
            raise ValueError("중량과 수량은 0보다 커야 합니다.")
        if t.unit == "kg" and kg > 30000:
            raise ValueError(f"중량 {kg:,} kg 이 너무 큽니다. kg 단위로 넣으세요.")
        qty = int(qty) if qty.is_integer() else qty
        net = t.from_kg(kg)
        if isinstance(net, float) and net.is_integer():
            net = int(net)
        v.update(QTY=qty, NET_WEIGHT=net)
        v["_kg"] = kg
        return v

    # ---------- 새 입고 ----------
    def save_new(self):
        t = self.tab
        try:
            v = self.inputs()
            no = self.v_no.get().strip()
            if not no:
                raise ValueError("입고번호를 넣으세요.")
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return

        def check(conn):
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM ICOM_ITEM_MASTER WHERE ITEM_CODE = :i", i=v["ITEM_CODE"])
            if not cur.fetchone()[0]:
                raise ValueError(f"품목 마스터에 없는 품번입니다: {v['ITEM_CODE']}")
            cur.execute(f"SELECT COUNT(*) FROM {t.T_RCV} WHERE TO_CHAR(RECEIPT_NO) = :n", n=no)
            if cur.fetchone()[0]:
                raise ValueError(f"입고번호 {no}는 이미 있습니다.")
            warn = []
            cur.execute(f"SELECT RECEIPT_NO, RECEIPT_DATE FROM {t.T_RCV} WHERE ITEM_CODE = :i AND HEAT_NO = :h "
                        f"AND IO_FLAG = 1", i=v["ITEM_CODE"], h=v["HEAT_NO"])
            same = cur.fetchall()
            if same:
                warn.append("같은 품번·HEAT_NO 입고가 이미 있습니다: " + ", ".join(
                    f"{show(a)}({b:%m/%d})" if b else show(a) for a, b in same[:5]))
            cur.execute(f"SELECT COUNT(*) FROM {t.T_STK} WHERE ITEM_CODE = :i", i=v["ITEM_CODE"])
            n = cur.fetchone()[0]
            if n > 1:
                raise ValueError(f"소재 재고({t.T_STK})에 품번 {v['ITEM_CODE']} 줄이 {n}개라 재고를 더할 수 없습니다.")
            if not n:
                warn.append("이 소재는 재고 줄이 없어 새로 만듭니다.")
            return warn
        try:
            warn = self.app.run_db(check)
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return
        except Exception as e:
            self.app.err(e)
            return
        msg = (f"MES에 소재 입고 1건을 추가합니다.\n\n  입고번호: {no}\n  입고일: {v['RECEIPT_DATE']:%Y-%m-%d}\n"
               f"  품번: {v['ITEM_CODE']}\n  HEAT_NO: {v['HEAT_NO']}\n  업체: {show(v['VENDOR_SITE_ID'])}\n"
               f"  수량: {v['QTY']}롤   중량: {v['_kg']:,} kg\n\n입고 + 재고LOT 추가, 소재 재고에 더하기 (기존 입고는 안 바뀜)")
        if warn:
            msg += "\n\n※ 확인하세요:\n  " + "\n  ".join(warn)
        if not messagebox.askyesno("소재 입고 등록", msg + "\n\n등록할까요?", parent=self.win,
                                   icon="warning" if warn else "question"):
            return
        user = (self.app.cfg.get("reg_user") or "").strip() or v["CHARGER"] or "ADMIN"
        now = dt.datetime.now()

        def do(conn):
            cur = conn.cursor()
            try:
                cur.execute(f"SELECT * FROM (SELECT * FROM {t.T_RCV} WHERE ITEM_CODE = :i AND IO_FLAG = 1 "
                            f"ORDER BY RECEIPT_DATE DESC) WHERE ROWNUM = 1", i=v["ITEM_CODE"])
                r = cur.fetchone()
                tpl = dict(zip([d[0] for d in cur.description], r)) if r else {}
                org = tpl.get("ORGANIZATION_ID") or 1
                unit = tpl.get("QTY_UNIT") or "RL"
                base = {c: tpl.get(c) for c, m in self.m_rcv.items() if m["notnull"] and not m["hasdef"]}
                rcv = dict(base)
                rcv.update(RECEIPT_NO=no, SEQ=1, RECEIPT_DATE=v["RECEIPT_DATE"], ITEM_CODE=v["ITEM_CODE"],
                           VENDOR_SITE_ID=v["VENDOR_SITE_ID"], QTY=v["QTY"], QTY_UNIT=unit,
                           NET_WEIGHT=v["NET_WEIGHT"], TOTAL_WEIGHT=v["NET_WEIGHT"], HEAT_NO=v["HEAT_NO"],
                           CHARGER=v["CHARGER"], ORGANIZATION_ID=org, ENTER_DATE=now, ENTER_BY=user,
                           LAST_MODIFY_DATE=now, LAST_MODIFY_BY=user, IO_FLAG=1, DIVIDE_FLAG="N")
                rcv = insert_dict(cur, t.T_RCV, self.m_rcv, rcv)
                cur.execute(f"SELECT * FROM (SELECT * FROM {t.T_LOT} ORDER BY ENTER_DATE DESC) WHERE ROWNUM = 1") \
                    if "ENTER_DATE" in self.m_lot else cur.execute(f"SELECT * FROM {t.T_LOT} WHERE ROWNUM = 1")
                r = cur.fetchone()
                ltpl = dict(zip([d[0] for d in cur.description], r)) if r else {}
                lot = {c: ltpl.get(c) for c, m in self.m_lot.items() if m["notnull"] and not m["hasdef"]}
                lot.update(RECEIPT_NO=no, RECEIPT_SEQ=1, HEAT_NO=v["HEAT_NO"], ITEM_CODE=v["ITEM_CODE"],
                           VENDOR_SITE_ID=v["VENDOR_SITE_ID"], QTY=v["QTY"], QTY_UNIT=unit,
                           UNIT_WEIGHT=v["NET_WEIGHT"], TOTAL_WEIGHT=v["NET_WEIGHT"], ORGANIZATION_ID=org,
                           ENTER_DATE=now, ENTER_BY=user, LAST_MODIFY_DATE=now, LAST_MODIFY_BY=user, DIVIDE_FLAG="N")
                insert_dict(cur, t.T_LOT, self.m_lot, lot)
                smsg = t.stock_add(cur, v["ITEM_CODE"], v["QTY"], v["_kg"], user, org=org, unit=unit, meta=self.m_stk)
                conn.commit()
                return rcv, smsg
            except Exception:
                conn.rollback()
                raise
        try:
            rcv, smsg = self.app.run_db(do)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "등록 중 오류 -> 전부 취소 (MES는 그대로)\n\n" + ora_hint(e), parent=self.win)
            return
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", user, "소재입고등록", t.T_RCV, no,
                              f"{v['ITEM_CODE']} HEAT {v['HEAT_NO']} {v['QTY']}롤 {v['_kg']}kg; {smsg}"]])
        self.app.load_log()
        t.d2.set(max(t.d2.get(), v["RECEIPT_DATE"].date()))
        t.d1.set(min(t.d1.get(), v["RECEIPT_DATE"].date()))
        t.fetch()
        if self.v_keep.get():
            self.v_heat.set("")
            self.no_auto = True
            self.suggest_no()
            self.lbl_item.config(text=f"✔ {no} 등록 완료 ({smsg}).  다음 코일 HEAT_NO·중량을 넣으세요.", fg="#1565C0")
        else:
            self.win.destroy()
            messagebox.showinfo(APP_TITLE, f"소재 입고를 등록했습니다: {no}\n{smsg}")

    # ---------- 고치기 ----------
    def save_edit(self):
        t, o = self.tab, self.orig
        try:
            v = self.inputs()
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return
        kg_new = v.pop("_kg")
        ch = {c: (o.get(c), x) for c, x in v.items() if c in o and cell_text(o.get(c)) != cell_text(x)}
        if "NET_WEIGHT" in ch and "TOTAL_WEIGHT" in o and cell_text(o.get("TOTAL_WEIGHT")) == cell_text(o.get("NET_WEIGHT")):
            ch["TOTAL_WEIGHT"] = (o["TOTAL_WEIGHT"], v["NET_WEIGHT"])
        if not ch:
            messagebox.showinfo(APP_TITLE, "바뀐 내용이 없습니다.", parent=self.win)
            return
        if self.used and set(ch) & set(self.FIELDS_LOCK_USED):
            messagebox.showwarning(APP_TITLE, "사용·출고·취소 이력이 있는 입고는 업체·담당자만 고칠 수 있습니다.", parent=self.win)
            return
        names = {"RECEIPT_DATE": "입고일", "ITEM_CODE": "품번", "HEAT_NO": "HEAT_NO", "VENDOR_SITE_ID": "업체",
                 "QTY": "수량", "NET_WEIGHT": f"순중량({t.unit})", "TOTAL_WEIGHT": f"총중량({t.unit})", "CHARGER": "담당자"}
        lines = "\n".join(f"  {names.get(c, c)}: {cell_text(a) or '(빈칸)'} → {cell_text(b) or '(빈칸)'}"
                          for c, (a, b) in ch.items())
        stock_ch = bool(set(ch) & {"ITEM_CODE", "QTY", "NET_WEIGHT"})
        if not messagebox.askyesno("소재 입고 수정", f"입고 {show(o['RECEIPT_NO'])}를 아래처럼 고칩니다.\n\n{lines}\n\n"
                                   f"재고LOT도 같이 고칩니다." + ("\n소재 재고도 차이만큼 맞춥니다." if stock_ch else "")
                                   + "\n\n저장할까요?", parent=self.win):
            return
        user = (self.app.cfg.get("reg_user") or "").strip() or "PYEDIT"

        def do(conn):
            cur = conn.cursor()
            try:
                u = t.usage(conn, o)
                if u and set(ch) & set(self.FIELDS_LOCK_USED):
                    raise RuntimeError("그 사이 생산에 사용·출고되었습니다. 업체·담당자만 고칠 수 있습니다.")
                # 1) 입고
                sets, where, b = [], ["ROWID = CHARTOROWID(:rid)"], {"rid": o["RID__"]}
                for i, (c, (a, x)) in enumerate(ch.items()):
                    sets.append(f"{c} = :n{i}")
                    b[f"n{i}"] = fit_type(self.m_rcv, c, x)
                    if a is None:
                        where.append(f"{c} IS NULL")
                    else:
                        where.append(f"{c} = :o{i}")
                        b[f"o{i}"] = a
                for c, how in (("LAST_MODIFY_DATE", "SYSDATE"), ("LAST_MODIFY_BY", user)):
                    if c in self.m_rcv:
                        sets.append(f"{c} = SYSDATE" if how == "SYSDATE" else f"{c} = :mb")
                        if how != "SYSDATE":
                            b["mb"] = how
                cur.execute(f"UPDATE {t.T_RCV} SET {', '.join(sets)} WHERE {' AND '.join(where)}", b)
                if cur.rowcount != 1:
                    raise RuntimeError("그 사이 다른 곳에서 바뀌었거나 없어진 입고입니다. 다시 조회하세요.")
                # 2) 재고LOT
                lmap = {"HEAT_NO": "HEAT_NO", "ITEM_CODE": "ITEM_CODE", "VENDOR_SITE_ID": "VENDOR_SITE_ID",
                        "QTY": "QTY", "NET_WEIGHT": "UNIT_WEIGHT"}
                lset, lb = [], {"rno": o["RECEIPT_NO"], "seq": o["SEQ"], "heat": o.get("HEAT_NO")}
                for c, lc in lmap.items():
                    if c in ch and lc in self.m_lot:
                        lset.append(f"{lc} = :l_{lc}")
                        lb[f"l_{lc}"] = fit_type(self.m_lot, lc, ch[c][1])
                if "NET_WEIGHT" in ch and "TOTAL_WEIGHT" in self.m_lot:
                    lset.append("TOTAL_WEIGHT = :l_tw")
                    lb["l_tw"] = fit_type(self.m_lot, "TOTAL_WEIGHT", ch["NET_WEIGHT"][1])
                lmsg = "LOT 변경 없음"
                if lset:
                    if "LAST_MODIFY_DATE" in self.m_lot:
                        lset.append("LAST_MODIFY_DATE = SYSDATE")
                    cur.execute(f"UPDATE {t.T_LOT} SET {', '.join(lset)} WHERE RECEIPT_NO = :rno AND RECEIPT_SEQ = :seq "
                                f"AND NVL(HEAT_NO, '~') = NVL(:heat, '~')", lb)
                    if cur.rowcount != 1:
                        raise RuntimeError(f"재고LOT 줄이 {cur.rowcount}개라 고칠 수 없습니다 (1개여야 함). 저장 안 함.")
                    lmsg = "LOT 1줄 수정"
                # 3) 소재 재고 (차이만큼)
                smsg = ""
                if stock_ch:
                    kg_old = t.to_kg(o.get("NET_WEIGHT")) or 0
                    q_old = o.get("QTY") or 0
                    s1 = t.stock_add(cur, o["ITEM_CODE"], -q_old, -kg_old, user, meta=self.m_stk)
                    s2 = t.stock_add(cur, v["ITEM_CODE"], v["QTY"], kg_new, user, org=o.get("ORGANIZATION_ID") or 1,
                                     unit=o.get("QTY_UNIT") or "RL", meta=self.m_stk)
                    smsg = f"; 재고: {s1} / {s2}" if o["ITEM_CODE"] != v["ITEM_CODE"] else f"; 재고: {s2}"
                conn.commit()
                return lmsg + smsg
            except Exception:
                conn.rollback()
                raise
        try:
            res = self.app.run_db(do)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "수정 중 오류 -> 전부 취소 (MES는 그대로)\n\n" + ora_hint(e), parent=self.win)
            return
        now = dt.datetime.now()
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", user, "소재입고수정", t.T_RCV, show(o["RECEIPT_NO"]),
                              "; ".join(f"{c}: {cell_text(a)} -> {cell_text(x)}" for c, (a, x) in ch.items()) + f"; {res}"]])
        self.app.load_log()
        self.win.destroy()
        messagebox.showinfo(APP_TITLE, f"저장했습니다: {show(o['RECEIPT_NO'])}\n{res}")
        t.fetch()


WO_RE = re.compile(r"^WO([0-9A-Z]+?)P(\d+)$")

# 작업지시번호 앞부분(WO 다음 ~ P 앞) 만드는 규칙 후보. 과거 번호로 어느 규칙인지 알아냄
WO_RULES = (
    ("년2+월(1~12)+일2", lambda d: f"{d:%y}{d.month}{d:%d}"),
    ("년2+월2+일2", lambda d: f"{d:%y}{d:%m}{d:%d}"),
    ("년2+월(1~9,A~C)+일2", lambda d: f"{d:%y}{'123456789ABC'[d.month - 1]}{d:%d}"),
    ("년1+월2+일2", lambda d: f"{d.year % 10}{d:%m}{d:%d}"),
    ("년1+월(1~9,A~C)+일2", lambda d: f"{d.year % 10}{'123456789ABC'[d.month - 1]}{d:%d}"),
)


def wo_learn_rule(conn, tbl):
    """과거 작업지시번호와 지시일을 비교해서 번호 규칙을 찾음 (10~12월 자료를 가장 중요하게 봄)"""
    cur = conn.cursor()
    cur.execute(f"SELECT WORK_ORDER_NO, WORK_ORDER_DATE FROM (SELECT WORK_ORDER_NO, WORK_ORDER_DATE FROM {tbl} "
                f"WHERE WORK_ORDER_DATE >= ADD_MONTHS(SYSDATE, -15) AND WORK_ORDER_NO LIKE 'WO%' "
                f"ORDER BY DBMS_RANDOM.VALUE) WHERE ROWNUM <= 4000")
    score = [[0, 0] for _ in WO_RULES]       # [10~12월 맞은 수, 전체 맞은 수]
    n_late = 0
    for no, d in cur.fetchall():
        m = WO_RE.match(str(no or "").strip())
        if not m or not d:
            continue
        late = d.month >= 10
        n_late += late
        for i, (_, f) in enumerate(WO_RULES):
            if f(d) == m.group(1):
                score[i][1] += 1
                score[i][0] += late
    best = max(range(len(WO_RULES)), key=lambda i: (score[i][0], score[i][1]))
    if score[best][1] == 0:
        best = 0
    note = (f"{WO_RULES[best][0]} (과거 번호 {score[best][1]}건 일치"
            + (f", 10~12월 {score[best][0]}/{n_late}건" if n_late else ", 10~12월 자료 없음") + ")")
    return WO_RULES[best][1], note


def wo_numbers(conn, tbl, dates):
    """날짜 목록 -> 작업지시번호 목록 (같은 날 번호가 있으면 그 앞부분, 없으면 과거 규칙. 일련번호는 이어서)"""
    cur = conn.cursor()
    cur.execute(f"SELECT WORK_ORDER_NO, WORK_ORDER_DATE FROM {tbl} WHERE WORK_ORDER_DATE >= :d",
                d=dt.datetime.now() - dt.timedelta(days=90))
    seq, same = 0, {}
    for no, d in cur.fetchall():
        m = WO_RE.match(str(no or "").strip())
        if not m:
            continue
        seq = max(seq, int(m.group(2)))
        if d:
            same.setdefault(d.date(), m.group(1))
    rule, note = None, ""
    out = []
    for d in dates:
        pre = same.get(d.date())
        if pre is None:
            if rule is None:
                rule, note = wo_learn_rule(conn, tbl)
            pre = rule(d)
        seq += 1
        out.append(f"WO{pre}P{seq}")
    return out, note


# ==========================================================
# 글자 치면 바로 걸러지는 입력칸 (목록에서 고르기 + 직접 입력)
# ==========================================================
class FilterEntry(ttk.Frame):
    """items: [(값, 보이는글)]  - 칸에 글자를 치면 값/보이는글에 그 글자가 든 것만 아래 목록에 보여줌.
    ↓ = 목록으로, Enter/클릭 = 고르기, Esc = 닫기. 목록에 없어도 직접 입력한 글자를 그대로 씀."""

    def __init__(self, parent, items=(), width=20, on_pick=None, max_show=300, on_change=None, rows=12):
        super().__init__(parent)
        self.items = list(items)
        self.on_pick = on_pick
        self.on_change = on_change
        self.last_text = ""
        self.rows = rows                       # 목록에 한 번에 보이는 줄 수
        self.max_show = max_show
        self.var = tk.StringVar()
        self.ent = ttk.Entry(self, textvariable=self.var, width=width)
        self.ent.pack(side="left")
        self.btn = ttk.Button(self, text="▼", width=2, command=self.toggle)
        self.btn.pack(side="left")
        self.pop = None
        self.shown = []
        self.ent.bind("<KeyRelease>", self.on_key)
        self.ent.bind("<Down>", lambda e: self.focus_list())
        self.ent.bind("<Escape>", lambda e: self.close())
        self.ent.bind("<FocusOut>", lambda e: self.after(200, self.close_if_away))

    def get(self):
        return self.var.get().strip()

    def set(self, v):
        self.var.set(v)

    def set_items(self, items):
        self.items = list(items)

    def toggle(self):
        if self.pop:
            self.close()
        else:
            self.open("")                      # ▼ = 지금 칸 글자와 상관없이 목록 전체
            self.ent.focus_set()
            cur = self.get().upper()
            for i, it in enumerate(self.shown):
                if str(it[0]).upper() == cur:
                    self.lb.selection_set(i)
                    self.lb.see(i)
                    break

    def on_key(self, e):
        if e.keysym in ("Return", "KP_Enter", "Escape", "Down", "Up", "Tab", "Shift_L", "Shift_R"):
            if e.keysym in ("Return", "KP_Enter") and self.pop and len(self.shown) == 1:
                self.pick(0)
            return
        t = self.get()
        if t != self.last_text:
            self.last_text = t
            if self.on_change:
                self.on_change(t)
        self.open(t)

    def open(self, text):
        t = text.upper().replace(" ", "")
        self.shown = [it for it in self.items
                      if not t or t in str(it[0]).upper().replace(" ", "") or t in str(it[1]).upper().replace(" ", "")][:self.max_show]
        if not self.shown:
            self.close()
            return
        if self.pop is None:
            self.pop = tk.Toplevel(self)
            self.pop.wm_overrideredirect(True)
            self.pop.attributes("-topmost", True)
            fr = ttk.Frame(self.pop)
            fr.pack(fill="both", expand=True)
            self.lb = tk.Listbox(fr, height=self.rows, font=("맑은 고딕", 9), activestyle="dotbox", exportselection=False)
            sb = ttk.Scrollbar(fr, orient="vertical", command=self.lb.yview)
            self.lb.configure(yscrollcommand=sb.set)
            self.lb.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            self.lb.bind("<ButtonRelease-1>", lambda e: self.pick(self.lb.nearest(e.y)))
            self.lb.bind("<Return>", lambda e: self.pick(self.lb.curselection()[0] if self.lb.curselection() else 0))
            self.lb.bind("<Escape>", lambda e: (self.close(), self.ent.focus_set()))
            self.lb.bind("<FocusOut>", lambda e: self.after(200, self.close_if_away))
        self.lb.delete(0, "end")
        for v, d in self.shown:
            self.lb.insert("end", d if d else v)
        x = self.ent.winfo_rootx()
        y = self.ent.winfo_rooty() + self.ent.winfo_height()
        # 폭 = 가장 긴 글자에 맞춤 (칸 폭보다 좁아지지는 않음), 높이 = 보이는 줄 수에 맞춤
        try:
            import tkinter.font as tkfont
            f = tkfont.Font(font=self.lb.cget("font"))
            tw = max(f.measure(d if d else v) for v, d in self.shown)
        except Exception:
            tw = 300
        w = max(self.ent.winfo_width() + self.btn.winfo_width(), min(tw + 40, 700))
        self.lb.configure(height=min(self.rows, len(self.shown)))
        self.pop.update_idletasks()
        h = self.lb.winfo_reqheight() + 4
        sh = self.winfo_screenheight()
        if y + h > sh - 40:                    # 화면 아래로 넘치면 칸 위쪽에 띄움
            y = max(0, self.ent.winfo_rooty() - h)
        self.pop.geometry(f"{w}x{h}+{x}+{y}")

    def focus_list(self):
        if not self.pop:
            self.open(self.get())
        if self.pop:
            self.lb.focus_set()
            self.lb.selection_clear(0, "end")
            self.lb.selection_set(0)
            self.lb.activate(0)
        return "break"

    def pick(self, i):
        if not self.shown or i is None or i >= len(self.shown):
            return
        v = self.shown[i][0]
        self.var.set(v)
        self.last_text = v
        self.close()
        self.ent.focus_set()
        self.ent.icursor("end")
        if self.on_pick:
            self.on_pick(v)

    def close_if_away(self):
        try:
            f = self.focus_get()
        except Exception:
            f = None
        if self.pop and f not in (self.ent, getattr(self, "lb", None)):
            self.close()

    def close(self):
        if self.pop is not None:
            self.pop.destroy()
            self.pop = None



def wo_open_recent(g):
    """최근 30일 호기별 최종 작업지시를 불러오기 전
    ① 기본 계획일자 → ② 정상/잔업 기준 순서로 먼저 선택한다.
    이 단계에서는 MES를 변경하지 않는다.
    """
    result = {"date": None, "hours": None}

    # ------------------------------------------------------
    # 1단계: 기본 계획일자 먼저 선택
    # ------------------------------------------------------
    win = tk.Toplevel(g.app.root)
    win.title("최근 작업지시 불러오기 - 기본 계획일자 선택")
    win.geometry("470x205")
    win.resizable(False, False)
    win.transient(g.app.root)
    win.grab_set()

    fr = ttk.Frame(win, padding=16)
    fr.pack(fill="both", expand=True)
    ttk.Label(fr, text="적용할 작업일자를 선택하세요",
              font=("맑은 고딕", 12, "bold")).pack(anchor="w", pady=(0, 8))
    ttk.Label(fr, text="최근 30일의 호기별 최종 작업지시를 불러온 뒤\n"
                       "아래 날짜를 새 작업지시의 기본 계획일자로 사용합니다.",
              foreground="#555", justify="left").pack(anchor="w", pady=(0, 12))

    df = ttk.Frame(fr)
    df.pack(anchor="w", pady=3)
    ttk.Label(df, text="기본 계획일자", font=("", 10, "bold")).pack(side="left", padx=(0, 8))
    de = DateEntry(df, dt.date.today())
    de.pack(side="left")

    def choose_date():
        try:
            result["date"] = de.get()
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=win)
            return
        win.destroy()

    bf = ttk.Frame(fr)
    bf.pack(fill="x", pady=(14, 0))
    ttk.Button(bf, text="다음", style="Big.TButton",
               command=choose_date).pack(side="left", expand=True, fill="x", padx=(0, 5))
    ttk.Button(bf, text="취소", command=win.destroy).pack(side="left", expand=True, fill="x", padx=(5, 0))

    win.protocol("WM_DELETE_WINDOW", win.destroy)
    try:
        win.update_idletasks()
        x = g.app.root.winfo_rootx() + max(0, (g.app.root.winfo_width() - win.winfo_width()) // 2)
        y = g.app.root.winfo_rooty() + max(0, (g.app.root.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")
    except Exception:
        pass
    win.wait_window()

    if result["date"] is None:
        return

    # ------------------------------------------------------
    # 2단계: 정상 / 잔업 기준 선택
    # ------------------------------------------------------
    win = tk.Toplevel(g.app.root)
    win.title("최근 작업지시 불러오기 - 근무 기준 선택")
    win.geometry("490x285")
    win.resizable(False, False)
    win.transient(g.app.root)
    win.grab_set()

    fr = ttk.Frame(win, padding=16)
    fr.pack(fill="both", expand=True)
    ttk.Label(fr, text=f"계획일자 {result['date']:%Y-%m-%d}",
              font=("맑은 고딕", 10, "bold"), foreground="#1F4E79").pack(anchor="w", pady=(0, 7))
    ttk.Label(fr, text="어느 근무시간까지 작업지시를 만들까요?",
              font=("맑은 고딕", 12, "bold")).pack(anchor="w", pady=(0, 8))
    ttk.Label(fr, text="선택한 시간 기준으로 최근 SPM × 근무시간 × 캐비티를 계산해\n"
                       "계획수량을 미리 배분합니다. 아직 MES에는 등록되지 않습니다.",
              foreground="#555", justify="left").pack(anchor="w", pady=(0, 14))

    def choose_hours(hours):
        result["hours"] = hours
        win.destroy()

    bf = ttk.Frame(fr)
    bf.pack(fill="x", pady=4)
    ttk.Button(bf, text="정상까지 (8시간)", style="Big.TButton",
               command=lambda: choose_hours(8)).pack(side="left", expand=True, fill="x", padx=(0, 5))
    ttk.Button(bf, text="잔업까지 (10시간)", style="Big.TButton",
               command=lambda: choose_hours(10)).pack(side="left", expand=True, fill="x", padx=(5, 0))
    mf = ttk.Frame(fr)                                  # 시간 직접 입력
    mf.pack(fill="x", pady=(10, 0))
    ttk.Label(mf, text="직접 입력", font=("", 10, "bold")).pack(side="left")
    v_man = tk.StringVar()
    e_man = ttk.Entry(mf, textvariable=v_man, width=7)
    e_man.pack(side="left", padx=(8, 4))
    ttk.Label(mf, text="시간").pack(side="left")

    def choose_manual(_e=None):
        try:
            h = float(v_man.get().replace(",", "").strip())
            if not 0 < h <= 24:
                raise ValueError
        except ValueError:
            messagebox.showwarning(APP_TITLE, "근무시간은 0보다 크고 24 이하인 숫자로 넣으세요. (예: 9, 11.5)", parent=win)
            return
        choose_hours(h)
    e_man.bind("<Return>", choose_manual)
    ttk.Button(mf, text="이 시간으로", command=choose_manual).pack(side="left", padx=8)
    ttk.Button(fr, text="취소", command=win.destroy).pack(side="right", pady=(12, 0))

    win.protocol("WM_DELETE_WINDOW", win.destroy)
    try:
        win.update_idletasks()
        x = g.app.root.winfo_rootx() + max(0, (g.app.root.winfo_width() - win.winfo_width()) // 2)
        y = g.app.root.winfo_rooty() + max(0, (g.app.root.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")
    except Exception:
        pass
    win.wait_window()

    if result["hours"]:
        WoBatch(g, recent_hours=result["hours"], recent_plan_date=result["date"])


# ==========================================================
# SPM 초기값: spm.xlsx 가 없을 때 처음 한 번만 이 값으로 파일을 만듦. 이후에는 [⑤ SPM 관리] 탭에서 관리
# ==========================================================
SPM_DEFAULT = {
    "1515-196": 45, "1547-102-1": 35, "1547-102-2": 45, "1549-156": 40, "1549-157": 35, "1551-189": 35,
    "1553-101-1": 35, "1553-101-2": 35, "1642-118-1": 30, "1654-119-1": 36, "1654-119-2": 32, "1657-127-1": 35,
    "1657-130": 35, "1660-131": 35, "1660-132": 35, "1664-128-1": 35, "1664-128-2": 35, "1752-297": 35,
    "1755-296": 35, "1781-0001-00": 35, "1807-653": 25, "1816-0010-00": 35, "1839-629": 28, "1837-0033-01": 35,
    "1853-515-1": 35, "1853-515-2": 35, "1853-586": 35, "1856-627-2": 60, "1856-628-1": 35, "1858-618": 35,
    "1858-618-2": 35, "1858-619": 35, "1861-596": 35, "1861-597": 35, "1862-0041-01": 30, "1862-0041-21": 25,
    "1862-0042-21": 25, "1862-604": 35, "1862-624": 35, "1862-651": 35, "1866-599": 30, "1866-617-1": 35,
    "1866-617-2": 35, "1869-621": 35, "1870-602-1": 35, "1870-602-2": 35, "1871-0043-01": 25,
    "1871-0043-21": 25, "1873-0044-21": 25, "1873-0045-01": 25, "1874-623": 35, "1888-0005": 20,
    "1945-349-1": 35, "1945-349-2": 40, "1946-0053-01": 35, "1947-353": 35, "1950-354-1": 35, "1950-354-2": 35,
    "1950-374-1": 35, "1950-374-2": 35, "1951-386": 40, "1952-0056-01": 30, "1952-0056-21": 25,
    "1952-346-1": 35, "1952-346-2": 35, "1952-367-2": 30, "1958-365-1": 40, "1958-365-2(2호금형)": 30,
    "1958-373-1": 35, "1958-373-2": 35, "1959-0024-01": 30, "1959-0024-21": 25, "1959-360-1": 35,
    "1959-360-2": 35, "1960-0031-01": 28, "1960-0031-21": 25, "1960-361": 35, "1962-0030": 35,
    "1963-0055-01": 30, "1963-0055-21": 25, "1965-326-1": 35, "1965-326-2": 35, "1967-0052-01": 25,
    "1967-0052-21": 25, "1967-391": 35, "1967-392": 35, "1968-0016-01": 30, "1968-0016-02": 40, "1968-362": 35,
    "1968-364-1": 40, "1968-364-2": 24, "1969-0030-01": 35, "1970-363": 35, "1970-372-1": 35, "1970-372-2": 35,
    "1972-383": 40, "1972-383(2호금형)": 40, "1973-325": 30, "1973-370(375)": 35, "1974-0023-01": 25,
    "1974-0023-21": 25, "1976-0004-01": 35, "1979-0051-01": 20, "1982-371": 25, "54610-C1000-(TR)": 22,
    "54610-C1001-(BK)": 25, "54610-D3500-(TR)": 22, "54610-D3501-(BK)": 25, "54610-D3000-(TR)": 23,
    "54610-D3001-(BK)": 22, "AB3038-SC": 23, "AB3248-SC-01": 22, "AT3037-GC-00": 23, "IJ223101-SC-00": 25,
    "C657-127": 35, "C873-622-2": 35, "1GE0-0009-01": 30, "116090-ST2": 28, "116080-ST0": 25,
    "IW/OW-BDA-1106-1": 38, "IW/OW-BDA-1106-2": 38, "IW/OW-BDA-1106-3": 38, "IW/OW-BDA-1113-1": 38,
    "IW/OW-BDA-1113-2": 38, "IW/OW-BDA-1113D": 38, "IW/OW-BDA-1124 BA": 35, "IW/OW-BDA-1156": 38,
    "IW/OW-BDA-1159B": 38, "IW/OW-BDA-1165A": 38, "IW/OW-BDA-1179": 38, "IW/OW-BDA-1179-2": 38,
    "MAG624828": 30, "MAG629430": 10, "MJH637370": 25, "48338-2H010": 36, "45631-26010 A": 25,
    "45631-26010 B/C": 25, "45631-26010-C": 25, "45624-2F010-U": 25, "45624-2F010-L": 25, "45524-26210": 25,
    "45464-4G600": 25, "45424-4G100L-A": 25, "45424-4G100L-B": 15, "45424-4G100L-C": 15, "D000136433": 22,
    "B17101000-0013-00": 22,
}
SPM_XLSX = os.path.join(BASE, "spm.xlsx")
SPM_HEAD = ["품번", "SPM", "수정일", "수정자", "비고"]


def spm_load():
    """spm.xlsx -> [{"item", "spm", "date", "user", "note"}]. 파일이 없으면 처음 한 번 기본표로 만들어 둠"""
    if openpyxl is None:
        raise RuntimeError("openpyxl이 없어 SPM 파일을 읽을 수 없습니다.")
    if not os.path.exists(SPM_XLSX):
        now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        spm_save([{"item": k, "spm": float(v), "date": now, "user": "초기값", "note": ""}
                  for k, v in SPM_DEFAULT.items()])
    wb = openpyxl.load_workbook(SPM_XLSX, read_only=True, data_only=True)
    out, seen = [], set()
    for r in wb.worksheets[0].iter_rows(min_row=2, values_only=True):
        r = list(r or []) + [None] * 5
        it = str(r[0] or "").strip()
        if not it or it.upper() in seen:
            continue
        try:
            v = float(str(r[1]).replace(",", ""))
        except (TypeError, ValueError):
            continue
        seen.add(it.upper())
        out.append({"item": it, "spm": v, "date": show(r[2]), "user": show(r[3]), "note": show(r[4])})
    wb.close()
    return out


def spm_save(rows):
    """[{"item", "spm", ...}] -> spm.xlsx (품번 순)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SPM"
    ws.append(SPM_HEAD)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in sorted(rows, key=lambda x: x["item"].upper()):
        v = r["spm"]
        ws.append([r["item"], int(v) if float(v).is_integer() else v, r.get("date", ""), r.get("user", ""),
                   r.get("note", "")])
    for col, w in zip("ABCDE", (22, 8, 18, 12, 30)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    try:
        wb.save(SPM_XLSX)
    except PermissionError:
        raise RuntimeError(f"spm.xlsx 가 엑셀에서 열려 있어 저장하지 못했습니다. 엑셀을 닫고 다시 하세요.\n{SPM_XLSX}")


def spm_table():
    """품번(대문자) -> SPM  ([⑤ SPM 관리] 탭에서 관리하는 spm.xlsx)"""
    try:
        return {r["item"].upper(): r["spm"] for r in spm_load()}, "SPM 관리"
    except Exception:
        return {}, "SPM 관리(읽기 실패)"


class SpmTab:
    """⑤ SPM 관리: 품번별 SPM 추가·수정·삭제 (spm.xlsx 에 저장, MES는 안 바뀜)"""
    COLS = ["품번", "SPM", "수정일", "수정자", "비고"]
    WIDTHS = [200, 70, 140, 90, 360]

    def __init__(self, app, t):
        self.app = app
        self.rows = []
        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=(6, 2))
        ttk.Label(bar, text="찾기").pack(side="left", padx=(4, 2))
        self.var_q = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.var_q, width=24)
        e.pack(side="left")
        self.var_q.trace_add("write", lambda *_: self.show())
        ttk.Button(bar, text="+ 새로 등록", style="Big.TButton", command=lambda: self.edit(None)).pack(side="left", padx=(12, 3))
        ttk.Button(bar, text="선택 SPM 고치기", style="Big.TButton", command=self.edit_sel).pack(side="left", padx=3)
        ttk.Button(bar, text="선택 줄 삭제", style="Big.TButton", command=self.delete).pack(side="left", padx=3)
        ttk.Button(bar, text="SPM 없는 품번 찾기 (최근 30일 작업)", command=self.find_missing).pack(side="left", padx=(12, 3))
        ttk.Button(bar, text="새로고침", command=self.load).pack(side="right", padx=3)
        ttk.Button(bar, text="엑셀 파일 열기", command=self.open_xlsx).pack(side="right", padx=3)
        ttk.Button(bar, text="엑셀에서 가져오기", command=self.import_xlsx).pack(side="right", padx=3)
        ttk.Label(t, text="줄 더블클릭 = 고치기.  계획수량 = SPM × 60 × 근무시간 × 캐비티.  "
                          f"저장 파일: {SPM_XLSX}  (MES는 바뀌지 않음)", foreground="#555").pack(fill="x", padx=6)
        fr, self.tv = app.make_tree(t, self.COLS, self.WIDTHS)
        fr.pack(fill="both", expand=True, pady=4)
        self.tv.bind("<Double-1>", lambda e: self.edit_sel())
        self.tv.bind("<Delete>", lambda e: self.delete())
        self.lbl = ttk.Label(t, text="", font=("", 10, "bold"))
        self.lbl.pack(fill="x", padx=6, pady=(0, 4))
        self.load()

    def load(self):
        try:
            self.rows = spm_load()
        except Exception as e:
            self.rows = []
            self.lbl.config(text=f"SPM 파일을 읽지 못했습니다: {e}", foreground="#C62828")
            return
        self.show()

    def save(self):
        spm_save(self.rows)

    def show(self):
        self.tv.delete(*self.tv.get_children())
        q = self.var_q.get().strip().upper().replace(" ", "")
        self.view = [r for r in sorted(self.rows, key=lambda x: x["item"].upper())
                     if not q or q in r["item"].upper().replace(" ", "") or q in r.get("note", "").upper()]
        for i, r in enumerate(self.view):
            self.tv.insert("", "end", iid=str(i), values=[r["item"], f"{r['spm']:g}", r.get("date", ""),
                                                         r.get("user", ""), r.get("note", "")])
        self.lbl.config(text=f"SPM 등록 {len(self.rows):,}개" + (f"  /  찾기 {len(self.view):,}개" if q else ""),
                        foreground="#1F4E79")

    def find(self, item):
        return next((r for r in self.rows if r["item"].upper() == str(item).strip().upper()), None)

    def edit_sel(self):
        sel = self.tv.selection()
        if len(sel) != 1:
            messagebox.showinfo(APP_TITLE, "고칠 줄 1개를 선택하세요.")
            return
        self.edit(self.view[int(sel[0])])

    def edit(self, r, item="", note=""):
        new = r is None
        w = tk.Toplevel(self.app.root)
        w.title("SPM 새로 등록" if new else f"SPM 고치기 - {r['item']}")
        w.resizable(False, False)
        w.transient(self.app.root)
        fr = ttk.Frame(w, padding=14)
        fr.pack(fill="both", expand=True)
        v_it = tk.StringVar(value=item if new else r["item"])
        v_spm = tk.StringVar(value="" if new else f"{r['spm']:g}")
        v_note = tk.StringVar(value=note if new else r.get("note", ""))
        ttk.Label(fr, text="품번", font=("", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        e_it = ttk.Entry(fr, textvariable=v_it, width=26, state="normal" if new else "readonly")
        e_it.grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(fr, text="SPM", font=("", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        e_spm = ttk.Entry(fr, textvariable=v_spm, width=10)
        e_spm.grid(row=1, column=1, sticky="w", padx=8)
        ttk.Label(fr, text="(분당 타수)", foreground="#777").grid(row=1, column=2, sticky="w")
        ttk.Label(fr, text="비고", font=("", 10, "bold")).grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(fr, textvariable=v_note, width=36).grid(row=2, column=1, columnspan=2, sticky="w", padx=8)

        def ok(_e=None):
            it = v_it.get().strip().upper()
            try:
                s = float(v_spm.get().replace(",", "").strip())
                if not 0 < s <= 1000:
                    raise ValueError
            except ValueError:
                messagebox.showwarning(APP_TITLE, "SPM은 0보다 큰 숫자로 넣으세요.", parent=w)
                return
            if not it:
                messagebox.showwarning(APP_TITLE, "품번을 넣으세요.", parent=w)
                return
            old = self.find(it)
            if new and old:
                if not messagebox.askyesno(APP_TITLE, f"품번 {it}는 이미 SPM {old['spm']:g} 로 등록되어 있습니다.\n"
                                                      f"{s:g} 로 바꿀까요?", parent=w):
                    return
            if new and not old:
                try:
                    n = self.app.run_db(lambda c: c.cursor().execute(
                        "SELECT COUNT(*) FROM ICOM_ITEM_MASTER WHERE ITEM_CODE = :i", i=it).fetchone()[0])
                    if not n and not messagebox.askyesno(APP_TITLE, f"품목 마스터에 없는 품번입니다: {it}\n그래도 등록할까요?",
                                                         parent=w):
                        return
                except Exception:
                    pass
            row = old if old else {"item": it}
            if not old:
                self.rows.append(row)
            prev = dict(row)
            row.update(spm=s, note=v_note.get().strip(), date=dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
                       user=(self.app.cfg.get("reg_user") or "").strip())
            try:
                self.save()
            except Exception as e:
                if old:
                    row.clear()
                    row.update(prev)
                else:
                    self.rows.remove(row)
                messagebox.showerror(APP_TITLE, str(e), parent=w)
                return
            self.app.append_log([[f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}", row["user"], "SPM등록" if not old else "SPM수정",
                                  "spm.xlsx", it, f"SPM {prev.get('spm', '') if old else '(없음)'} -> {s:g}"]])
            self.app.load_log()
            w.destroy()
            self.show()
            for iid in self.tv.get_children():
                if self.tv.item(iid, "values")[0] == row["item"]:
                    self.tv.selection_set(iid)
                    self.tv.see(iid)
                    break
        bf = ttk.Frame(fr)
        bf.grid(row=3, column=0, columnspan=3, sticky="w", pady=(12, 0))
        ttk.Button(bf, text="저장", style="Big.TButton", command=ok).pack(side="left", padx=(0, 4))
        ttk.Button(bf, text="취소", command=w.destroy).pack(side="left")
        for e in (e_it, e_spm):
            e.bind("<Return>", ok)
        w.bind("<Escape>", lambda e: w.destroy())
        try:
            w.grab_set()
        except Exception:
            pass
        (e_it if new and not item else e_spm).focus_set()

    def delete(self):
        sel = [self.view[int(i)] for i in self.tv.selection()]
        if not sel:
            messagebox.showinfo(APP_TITLE, "삭제할 줄을 선택하세요. (여러 줄: Ctrl/Shift+클릭)")
            return
        names = ", ".join(r["item"] for r in sel[:10]) + (f" 외 {len(sel) - 10}개" if len(sel) > 10 else "")
        if not messagebox.askyesno(APP_TITLE, f"SPM {len(sel)}개를 삭제합니다.\n{names}\n\n계속할까요?", icon="warning"):
            return
        keep = [r for r in self.rows if r not in sel]
        try:
            spm_save(keep)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        self.rows = keep
        self.app.append_log([[f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}", self.app.cfg.get("reg_user", ""), "SPM삭제",
                              "spm.xlsx", r["item"], f"SPM {r['spm']:g}"] for r in sel])
        self.app.load_log()
        self.show()

    def open_xlsx(self):
        if not os.path.exists(SPM_XLSX):
            self.save()
        messagebox.showinfo(APP_TITLE, "엑셀에서 고친 뒤 저장하고 닫은 다음 [새로고침]을 누르세요.\n"
                                       "(A열 품번, B열 SPM. 엑셀이 열려 있는 동안은 여기서 저장이 안 됩니다)")
        open_file(SPM_XLSX)

    def import_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            got = []
            for r in wb.worksheets[0].iter_rows(values_only=True):
                if not r or len(r) < 2 or r[0] in (None, ""):
                    continue
                try:
                    got.append((str(r[0]).strip().upper(), float(str(r[1]).replace(",", ""))))
                except (TypeError, ValueError):
                    continue                        # 제목줄 등
            wb.close()
        except Exception as e:
            self.app.err(e)
            return
        if not got:
            messagebox.showinfo(APP_TITLE, "가져올 줄이 없습니다. (A열 품번, B열 SPM 숫자)")
            return
        add = [g for g in got if not self.find(g[0])]
        chg = [g for g in got if self.find(g[0]) and self.find(g[0])["spm"] != g[1]]
        if not add and not chg:
            messagebox.showinfo(APP_TITLE, f"{len(got)}줄 모두 지금 값과 같습니다.")
            return
        if not messagebox.askyesno(APP_TITLE, f"엑셀 {len(got)}줄 중\n  새 품번 {len(add)}개 추가\n  SPM 바뀜 {len(chg)}개\n\n"
                                              f"반영할까요? (엑셀에 없는 기존 품번은 그대로 둡니다)"):
            return
        now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        user = (self.app.cfg.get("reg_user") or "").strip()
        rows = [dict(r) for r in self.rows]
        by = {r["item"].upper(): r for r in rows}
        for it, s in add + chg:
            r = by.get(it)
            if r is None:
                r = {"item": it, "note": ""}
                rows.append(r)
                by[it] = r
            r.update(spm=s, date=now, user=user)
        try:
            spm_save(rows)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        self.rows = rows
        self.app.append_log([[f"{dt.datetime.now():%Y-%m-%d %H:%M:%S}", user, "SPM가져오기", "spm.xlsx", os.path.basename(path),
                              f"추가 {len(add)} / 변경 {len(chg)}"]])
        self.app.load_log()
        self.show()
        messagebox.showinfo(APP_TITLE, f"추가 {len(add)}개, 변경 {len(chg)}개 반영했습니다.")

    def find_missing(self):
        """최근 30일 작업지시 품번 중 SPM이 없는 것"""
        tb = self.app.mes.qualify(self.app.cfg.get("wo_table") or "IPLN_WORK_ORDER_MASTER")

        def work(conn):
            cur = conn.cursor()
            cur.execute(f"SELECT ITEM_CODE, MAX(ITEM_NAME), COUNT(*), MAX(WORK_ORDER_DATE), "
                        f"MAX(MACHINE_NAME) KEEP (DENSE_RANK LAST ORDER BY WORK_ORDER_DATE) FROM {tb} "
                        f"WHERE WORK_ORDER_DATE >= TRUNC(SYSDATE) - 30 AND ITEM_CODE IS NOT NULL "
                        f"GROUP BY ITEM_CODE ORDER BY ITEM_CODE")
            return cur.fetchall()
        try:
            rows = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return
        miss = [r for r in rows if not self.find(show(r[0]).strip())]
        if not miss:
            messagebox.showinfo(APP_TITLE, f"최근 30일 작업 품번 {len(rows)}개 모두 SPM이 등록되어 있습니다.")
            return
        w = tk.Toplevel(self.app.root)
        w.title(f"SPM 없는 품번 - 최근 30일 작업 {len(miss)}개")
        w.geometry("760x480")
        w.transient(self.app.root)
        ttk.Label(w, text="줄 더블클릭 = 이 품번 SPM 등록.  SPM이 없으면 최근 생산 기록으로 계산합니다.",
                  foreground="#555").pack(fill="x", padx=8, pady=(8, 0))
        fr, tv = self.app.make_tree(w, ["품번", "품명", "작업 횟수", "마지막 작업일", "최근 설비"], [170, 250, 70, 100, 100])
        fr.pack(fill="both", expand=True, padx=8, pady=8)
        for i, r in enumerate(miss):
            tv.insert("", "end", iid=str(i), values=[show(r[0]), show(r[1]), show(r[2]),
                                                     r[3].strftime("%Y-%m-%d") if r[3] else "", show(r[4])])

        def add(_e=None):
            s = tv.selection()
            if s:
                r = miss[int(s[0])]
                self.edit(None, item=show(r[0]).strip(), note=show(r[1]).strip())
        tv.bind("<Double-1>", add)


def hours_text(h):
    h = float(h)
    return "정상 8시간" if h == 8 else ("잔업까지 10시간" if h == 10 else f"{h:g}시간 (직접 입력)")


class WoBatch:
    """③ 작업지시 관리 - 엑셀 붙여넣기 일괄등록 (MES '작업지시일괄처리'와 같은 방식)
    1) 엑셀에서 복사 → Ctrl+V   2) [검사]   3) [W/O 번호생성]   4) [작업지시 반영]
    반영 = IPLN_WORK_ORDER_UPLOADER 에 넣고 MES 프로그램 P_PLN_WORKORDER_BATCH_CREATE 실행 (MES 화면과 동일)"""
    T_UP = "IPLN_WORK_ORDER_UPLOADER"
    PROC = "P_PLN_WORKORDER_BATCH_CREATE"
    IN_COLS = ["설비명", "품번", "계획수량"]
    COLS = ["No", "설비(입력)", "계획일자", "교대조", "품번", "계획수량", "순위", "설비코드", "품명", "금형", "작업지시번호", "상태", "최근작업일", "최근SPM"]
    WIDTHS = [40, 90, 95, 55, 160, 75, 50, 75, 200, 120, 130, 330, 90, 85]

    def __init__(self, g, recent_hours=None, recent_plan_date=None):
        self.g, self.app = g, g.app
        try:
            self.recent_hours = float(recent_hours) if recent_hours and float(recent_hours) > 0 else None
        except (TypeError, ValueError):
            self.recent_hours = None
        self.recent_plan_date = recent_plan_date if isinstance(recent_plan_date, dt.date) else None
        self.rows = []             # dict: line, date, shift, item, qty, prio, mc, mname, iname, mold, no, st, lv
        self.checked = False
        self.sort_col = None        # 표 헤드 클릭 정렬 컬럼 번호
        self.sort_reverse = False   # False=오름차순, True=내림차순
        win = tk.Toplevel(self.app.root)
        self.win = win
        win.title("최근 30일 작업지시 - 등록 예정" if self.recent_hours else "작업지시 여러 건 일괄등록")
        win.geometry("1480x800" if self.recent_hours else "1380x720")
        win.transient(self.app.root)
        bg = "#EEF4FB"
        head = tk.Frame(win, bg=bg, highlightbackground="#B7CCE4", highlightthickness=1)
        head.pack(fill="x", padx=8, pady=8)
        if self.recent_hours:
            mode = hours_text(self.recent_hours)
            plan_day = (self.recent_plan_date or dt.date.today()).strftime("%Y-%m-%d")
            tk.Label(head, bg=bg, fg="#1F4E79", font=("맑은 고딕", 11, "bold"), anchor="w",
                     text=f"최근 30일 작업지시 불러오기 - 계획일자 {plan_day} / {mode} 기준 (아직 MES 미등록)").pack(fill="x", padx=10, pady=(6, 0))
            tk.Label(head, bg=bg, fg="#333", anchor="w", justify="left", font=("맑은 고딕", 9), text=(
                "최근 30일 작업지시에서 호기마다 마지막 1건(품번)을 불러옵니다.  "
                "계획수량 = SPM([⑤ SPM 관리] 값, 없으면 최근 생산 기록) × 60 × 근무시간 × 캐비티\n"
                "필요 없는 줄 = [선택 줄 빼기],  바꿀 줄 = [선택 줄 수정] 또는 설비·품번·수량 칸 더블클릭.  "
                "[작업지시 반영]을 누르기 전까지 MES는 바뀌지 않습니다.\n"
                "순서: 고치기·빼기 → [검사] → [W/O 번호생성] → 번호·상태 확인 → [작업지시 반영]")
            ).pack(fill="x", padx=10, pady=(0, 4))
        else:
            tk.Label(head, bg=bg, fg="#1F4E79", font=("맑은 고딕", 11, "bold"), anchor="w",
                     text="작업지시 여러 건 일괄등록 - 설비명·품번·계획수량만 넣으면 됩니다 (MES 작업지시일괄처리와 같은 방식)").pack(fill="x", padx=10, pady=(6, 0))
            tk.Label(head, bg=bg, fg="#333", anchor="w", justify="left", font=("맑은 고딕", 9), text=(
                "설비명 · 품번 · 계획수량 만 넣으면 됩니다. 계획일자·교대조는 아래 기본값, 순위는 설비별 순서대로 자동.\n"
                "[빠른 입력]: 설비 고르기 → 그 설비에서 최근 2년 만든 품번 목록 → 품번 고르면 최근 2년 그 품번을 만든 호기만 설비 목록에 남음\n"
                "   → SPM × 근무(정상 8h / 잔업 10h / 직접 h 입력)로 수량 자동 → Enter = 추가.  (▼ = 목록 전체, '전체 설비'/'전체 품번' = 안 거름, ↺ 처음부터·모두 비우기 = 거르기 초기화)\n"
                "엑셀에서 설비명·품번·수량 3칸을 복사해 표에 Ctrl+V 해도 됩니다.  "
                "MES 양식(라인번호·계획일자·교대조·품목코드·계획수량·순위)도 그대로 붙여넣을 수 있습니다.\n"
                "→ [검사] → [W/O 번호생성] → [작업지시 반영].  칸 더블클릭 = 고치기, Delete = 줄 지우기")
            ).pack(fill="x", padx=10, pady=(0, 4))
        lg = tk.Frame(head, bg=bg)                       # 행 색상 설명
        lg.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(lg, text="행 색상 ([검사] 후):", bg=bg, fg="#333", font=("맑은 고딕", 9, "bold")).pack(side="left")
        for color, txt in (("#C6EFCE", "초록 = 정상"),
                           ("#DDEBF7", "하늘 = 같은 날 이 품번 작업지시가 이미 있음 (중복 확인)"),
                           ("#FFEB9C", "노랑 = 확인 필요 (금형 연결 없음 / 소재중량 없음 / 표 안 중복)"),
                           ("#FFC7CE", "빨강 = 오류, 반영 안 됨 (설비·품번·수량 고치기)")):
            tk.Label(lg, text=f" {txt} ", bg=color, fg="#222", font=("맑은 고딕", 9),
                     relief="solid", bd=1).pack(side="left", padx=(6, 0))
        self.machines, self.item_list, self.item_last, self.mc_items, self.item_names = {}, [], {}, {}, {}
        self.item_mcs = {}
        try:
            self.app.run_db(self.load_lists)
        except Exception as e:
            messagebox.showwarning(APP_TITLE, f"설비·품번 목록을 읽지 못했습니다 (직접 입력은 됩니다).\n{ora_hint(e)}", parent=win)
        opt = ttk.Frame(win)
        opt.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Label(opt, text="기본 계획일자", font=("", 10, "bold")).pack(side="left")
        self.d_def = DateEntry(opt, self.recent_plan_date or dt.date.today())
        self.d_def.pack(side="left", padx=(4, 12))
        ttk.Label(opt, text="기본 교대조", font=("", 10, "bold")).pack(side="left")
        self.v_shift = tk.StringVar(value="1")
        ttk.Combobox(opt, textvariable=self.v_shift, values=["1", "2", "3"], width=4).pack(side="left", padx=(4, 16))
        ttk.Separator(opt, orient="vertical").pack(side="left", fill="y", padx=6)
        qf = tk.Frame(win, bg="#FFF7E6", highlightbackground="#F0C36D", highlightthickness=1)
        qf.pack(fill="x", padx=8, pady=(0, 4))
        Q = {"bg": "#FFF7E6"}
        tk.Label(qf, text="빠른 입력", font=("맑은 고딕", 10, "bold"), fg="#C55A11", **Q).pack(side="left", padx=(8, 10))
        tk.Label(qf, text="설비", **Q).pack(side="left")
        self.mc_all = sorted(((n or c), f"{n or c}   ({c})") for c, n in self.machines.items())
        self.f_m = FilterEntry(qf, self.mc_all, width=11, on_pick=self.on_machine_pick, on_change=self.on_mc_typed, rows=30)
        self.f_m.pack(side="left", padx=4, pady=6)
        self.v_allmc = tk.BooleanVar(value=False)
        tk.Checkbutton(qf, text="전체 설비", variable=self.v_allmc, command=self.refresh_machines, **Q).pack(side="left")
        self.v_allitem = tk.BooleanVar(value=False)
        tk.Checkbutton(qf, text="전체 품번", variable=self.v_allitem, command=self.refresh_items, **Q).pack(side="left")
        tk.Label(qf, text="   품번", **Q).pack(side="left")
        self.f_i = FilterEntry(qf, self.item_list, width=20, on_pick=self.on_item_pick, on_change=self.on_item_typed)
        self.f_i.pack(side="left", padx=4)
        tk.Label(qf, text="  근무", **Q).pack(side="left")
        self.v_hours = tk.DoubleVar(value=self.recent_hours or 8)
        self.v_hman = tk.StringVar(value="" if (self.recent_hours or 8) in (8, 10) else f"{self.recent_hours:g}")

        def pick_radio():
            self.v_hman.set("")                         # 정상/잔업 고르면 직접 입력 칸 비움
            self.calc_qty()
        for h, t in ((8, "정상 8h"), (10, "잔업 10h")):
            rb = tk.Radiobutton(qf, text=t, value=h, variable=self.v_hours, command=pick_radio, **Q)
            rb.pack(side="left")
            if self.recent_hours:
                rb.config(state="disabled")
        tk.Label(qf, text=" 직접", **Q).pack(side="left")
        e_h = ttk.Entry(qf, textvariable=self.v_hman, width=5)
        e_h.pack(side="left", padx=(2, 0))
        tk.Label(qf, text="h", **Q).pack(side="left")
        if self.recent_hours:
            e_h.config(state="disabled")

        def on_hman(_e=None):
            txt = self.v_hman.get().replace(",", "").strip()
            if not txt:
                return
            try:
                h = float(txt)
            except ValueError:
                return
            if 0 < h <= 24:
                self.v_hours.set(h)                     # 8·10이 아니면 라디오 선택이 풀림
                self.calc_qty()
        e_h.bind("<KeyRelease>", on_hman)
        e_h.bind("<FocusOut>", on_hman)
        tk.Label(qf, text="  SPM", **Q).pack(side="left")
        self.v_spm = tk.StringVar()
        e_spm = ttk.Entry(qf, textvariable=self.v_spm, width=6)
        e_spm.pack(side="left", padx=2)
        e_spm.bind("<KeyRelease>", lambda e: self.calc_qty())
        self.lbl_spm = tk.Label(qf, text="", fg="#777", font=("맑은 고딕", 8), **Q)
        self.lbl_spm.pack(side="left")
        tk.Label(qf, text="  수량", **Q).pack(side="left")
        self.v_qq = tk.StringVar()
        e_qq = ttk.Entry(qf, textvariable=self.v_qq, width=8)
        e_qq.pack(side="left", padx=4)
        self.e_qq = e_qq
        self.spm_info = (None, 1, "")                 # (SPM, 캐비티, 출처)
        e_qq.bind("<Return>", lambda e: self.quick_add())
        self.f_m.ent.bind("<Return>", lambda e: (self.f_m.pick(0) if self.f_m.pop else None, self.f_i.ent.focus_set()))
        self.f_i.ent.bind("<Return>", lambda e: (self.f_i.pick(0) if self.f_i.pop else None, e_qq.focus_set()))
        ttk.Button(qf, text="+ 추가", style="Big.TButton", command=self.quick_add).pack(side="left", padx=6)
        ttk.Button(qf, text="↺ 처음부터", command=self.reset_quick).pack(side="left", padx=2)
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8)
        ttk.Button(bar, text="붙여넣기 (Ctrl+V)", command=self.paste).pack(side="left", padx=3)
        ttk.Button(bar, text="선택 줄 수정", command=self.edit_selected).pack(side="left", padx=3)
        ttk.Button(bar, text="선택 줄 빼기", command=self.del_rows).pack(side="left", padx=3)
        ttk.Button(bar, text="모두 비우기", command=self.clear).pack(side="left", padx=3)
        ttk.Button(bar, text="작업지시 반영", style="Big.TButton", command=self.apply).pack(side="right", padx=3)
        ttk.Button(bar, text="W/O 번호생성", style="Big.TButton", command=self.make_no).pack(side="right", padx=3)
        ttk.Button(bar, text="검사", style="Big.TButton", command=self.check).pack(side="right", padx=3)
        fr, self.tv = self.app.make_tree(win, self.COLS, self.WIDTHS)
        fr.pack(fill="both", expand=True, padx=8, pady=6)
        self.set_sort_headers()
        self.tv.tag_configure("dup", background="#DDEBF7")
        self.lbl = ttk.Label(win, text="엑셀에서 복사한 뒤 표를 클릭하고 Ctrl+V 를 누르세요.", font=("", 10, "bold"))
        self.lbl.pack(fill="x", padx=10, pady=(0, 8))
        for w in (win, self.tv):
            w.bind("<Control-v>", lambda e: self.paste())
            w.bind("<Control-V>", lambda e: self.paste())
        self.tv.bind("<Double-1>", self.edit_cell)
        self.tv.bind("<Delete>", lambda e: self.del_rows())
        self.tv.focus_set()
        if self.recent_hours:
            self.prefill_recent(self.recent_hours)

    # ---------- 입력 ----------
    @staticmethod
    def parse_date(t):
        t = str(t or "").strip().split(" ")[0]
        if not t:
            return None
        if re.fullmatch(r"\d{1,2}[/.-]\d{1,2}", t):
            m, d = re.split(r"[/.-]", t)
            return dt.datetime(dt.date.today().year, int(m), int(d))
        if re.fullmatch(r"\d{5}(\.0)?", t):                     # 엑셀 날짜 숫자
            return dt.datetime(1899, 12, 30) + dt.timedelta(days=int(float(t)))
        return dt.datetime.combine(parse_day(t), dt.time())

    def load_lists(self, conn):
        """설비 목록 + 품번 목록(최근 2년 작업지시 품번 + 금형 연결 품번, 품명·최근 설비 같이)"""
        tb = self.g.table()
        link = self.app.mes.qualify(self.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")
        mst = self.app.mes.qualify(self.app.cfg.get("item_master_table") or "ICOM_ITEM_MASTER")
        cur = conn.cursor()
        cur.execute(f"SELECT MACHINE_CODE, MAX(MACHINE_NAME) FROM {tb} WHERE WORK_ORDER_DATE >= SYSDATE - 730 "
                    f"AND MACHINE_CODE IS NOT NULL GROUP BY MACHINE_CODE ORDER BY MACHINE_CODE")
        self.machines = {show(a).strip(): show(b).strip() for a, b in cur.fetchall()}
        cur.execute(f"SELECT ITEM_CODE, MAX(ITEM_NAME), MAX(MACHINE_CODE) KEEP (DENSE_RANK LAST ORDER BY WORK_ORDER_DATE), "
                    f"COUNT(*) FROM {tb} WHERE WORK_ORDER_DATE >= SYSDATE - 730 AND ITEM_CODE IS NOT NULL "
                    f"GROUP BY ITEM_CODE ORDER BY COUNT(*) DESC")
        seen = {}
        for it, nm, mc, n in cur.fetchall():
            it = show(it).strip()
            seen[it] = (show(nm).strip(), show(mc).strip())
        cur.execute(f"SELECT MACHINE_CODE, ITEM_CODE, COUNT(*), MAX(WORK_ORDER_DATE) FROM {tb} "
                    f"WHERE WORK_ORDER_DATE >= SYSDATE - 730 AND ITEM_CODE IS NOT NULL AND MACHINE_CODE IS NOT NULL "
                    f"GROUP BY MACHINE_CODE, ITEM_CODE ORDER BY MAX(WORK_ORDER_DATE) DESC")
        self.mc_items = {}                              # 설비 -> 최근 2년 그 설비에서 만든 품번 (최근 순)
        self.item_mcs = {}                              # 품번 -> 최근 2년 그 품번을 만든 설비 (최근 순)
        for mc, it, n, last in cur.fetchall():
            mc, it = show(mc).strip(), show(it).strip()
            self.mc_items.setdefault(mc, []).append((it, n, last))
            self.item_mcs.setdefault(it.upper(), []).append((mc, n, last))
        try:
            cur.execute(f"SELECT L.ITEM_CODE, (SELECT MAX(M.ITEM_NAME) FROM {mst} M WHERE M.ITEM_CODE = L.ITEM_CODE) "
                        f"FROM {link} L")
            for it, nm in cur.fetchall():
                it = show(it).strip()
                if it and it not in seen:
                    seen[it] = (show(nm).strip(), "")
        except Exception:
            pass
        self.item_last = {k: v[1] for k, v in seen.items() if v[1]}
        self.item_names = {k: v[0] for k, v in seen.items()}
        self.item_list = []
        for k, (nm, mc) in seen.items():
            mn = self.machines.get(mc, mc)
            self.item_list.append((k, f"{k}   |  {nm or '-'}" + (f"   |  최근 {mn}" if mn else "")))

    # ---------- 최근 30일 작업지시 자동 불러오기 ----------
    def recent_rows(self, conn, hours):
        """최근 30일에서 호기별 최종 작업지시 1건만 가져와 계획수량을 계산."""
        tb = self.g.table()
        cur = conn.cursor()
        cur.execute(f"""
            SELECT MACHINE_CODE, MACHINE_NAME, ITEM_CODE, ITEM_NAME,
                   WORK_ORDER_DATE AS LAST_DAY, WORK_ORDER_NO AS LAST_WO
              FROM (
                    SELECT T.MACHINE_CODE, T.MACHINE_NAME, T.ITEM_CODE, T.ITEM_NAME,
                           T.WORK_ORDER_DATE, T.WORK_ORDER_NO,
                           ROW_NUMBER() OVER (
                               PARTITION BY T.MACHINE_CODE
                               ORDER BY T.WORK_ORDER_DATE DESC,
                                        T.WORK_ORDER_NO DESC,
                                        T.ROWID DESC
                           ) AS RN
                      FROM {tb} T
                     WHERE T.WORK_ORDER_DATE >= TRUNC(SYSDATE) - 30
                       AND T.MACHINE_CODE IS NOT NULL
                       AND T.ITEM_CODE IS NOT NULL
                   )
             WHERE RN = 1
             ORDER BY LAST_DAY DESC, MACHINE_CODE
        """)
        base = cur.fetchall()
        out = []
        for mc, mname, it, iname, last_day, last_wo in base:
            mc, it = show(mc).strip(), show(it).strip().upper()
            if not mc or not it:
                continue
            try:
                spm, cav, src = self.spm_of(conn, mc, it, days=30)
            except Exception as e:
                spm, cav, src = None, 1, f"SPM 조회 실패: {str(e)[:50]}"
            qty = ""
            if spm:
                q = float(spm) * 60 * float(hours) * int(cav or 1)
                qty = str(int(round(q / 10.0) * 10))
            line = self.machines.get(mc) or show(mname).strip() or mc
            out.append({
                "line": line, "date_t": "", "shift": "", "item": it, "qty_t": qty,
                "prio": "", "st": (f"최근 {last_day:%m/%d} · {src}" if last_day else src),
                "lv": "" if qty else "warn", "recent_date": last_day, "recent_wo": show(last_wo),
                "spm": (float(spm) if spm else None), "spm_src": src,
                "recent_machine": mc, "recent_item_name": show(iname).strip(),
            })
        return out

    def prefill_recent(self, hours):
        mode = hours_text(hours)
        self.lbl.config(text=f"최근 30일 자료를 불러오는 중... ({mode})")
        self.win.update_idletasks()
        try:
            rows = self.app.run_db(lambda c: self.recent_rows(c, hours))
        except Exception as e:
            messagebox.showerror(APP_TITLE, "최근 작업지시를 불러오지 못했습니다.\n\n" + ora_hint(e), parent=self.win)
            self.lbl.config(text="최근 작업지시 불러오기 실패")
            return
        self.rows = rows
        self.checked = False
        self.show()
        no_spm = sum(1 for r in rows if not r.get("qty_t"))
        self.lbl.config(text=f"최근 30일 호기별 최종 작업지시 {len(rows)}건 불러옴 · {mode} 기준 계획수량 자동계산"
                             + (f" · SPM 없어 직접 입력 필요 {no_spm}개" if no_spm else "")
                             + "  → 필요 없는 줄은 빼고, 수정 후 [검사]")
        if not rows:
            messagebox.showinfo(APP_TITLE, "최근 30일 작업지시가 없습니다.", parent=self.win)

    # ---------- 빠른 입력: 설비 → 그 설비 품번 → SPM으로 수량 ----------
    def cur_mc(self):
        t = self.f_m.get()
        if t in self.machines:
            return t
        for c, n in self.machines.items():
            if n and n.replace(" ", "").upper() == t.replace(" ", "").upper():
                return c
        return ""

    def on_machine_pick(self, _v=None):
        self.refresh_items()
        it = self.f_i.get().upper()
        mc = self.cur_mc()
        if it and mc and any(m == mc for m, _n, _l in self.item_mcs.get(it, [])):
            self.update_spm(it)                         # 고른 품번을 만든 다른 호기로 바꾼 경우: 품번 유지, SPM 다시
            return
        self.f_i.set("")
        self.refresh_machines("")
        self.f_i.ent.focus_set()
        self.f_i.open("")

    def on_mc_typed(self, t):
        """설비 칸 글자를 고치기 시작하면 품번 목록은 다시 전체 (서로 계속 좁아지는 것 방지)"""
        if self.cur_mc() == "":
            self.f_i.set_items(self.item_list)

    def on_item_typed(self, t):
        """품번 칸 글자를 고치기 시작하면 설비 목록은 다시 전체"""
        self.f_m.set_items(self.mc_all)

    def reset_quick(self):
        """빠른 입력 칸·목록 거르기를 처음 상태로"""
        for f in (self.f_m, self.f_i):
            f.close()
            f.set("")
            f.last_text = ""
        self.v_allmc.set(False)
        self.v_allitem.set(False)
        self.f_m.set_items(self.mc_all)
        self.f_i.set_items(self.item_list)
        self.v_spm.set("")
        self.v_qq.set("")
        self.lbl_spm.config(text="")
        self.spm_info = (None, 1, "")
        self.f_m.ent.focus_set()

    def refresh_machines(self, item=None):
        """품번이 있으면 최근 2년 그 품번을 만든 호기만, 없거나 '전체 설비' 체크면 모든 설비"""
        item = (self.f_i.get() if item is None else item).upper()
        hist = self.item_mcs.get(item) if item else None
        if self.v_allmc.get() or not hist:
            self.f_m.set_items(self.mc_all)
            return []
        items = []
        for mc, n, last in hist:
            nm = self.machines.get(mc) or mc
            items.append((nm, f"{nm}   ({mc})   |  {n}회" + (f", 마지막 {last:%y/%m/%d}" if last else "")))
        self.f_m.set_items(items)
        return hist

    def refresh_items(self):
        mc = self.cur_mc()
        if self.v_allitem.get() or not mc or mc not in self.mc_items:
            self.f_i.set_items(self.item_list)
            return
        items = []
        for it, n, last in self.mc_items[mc]:
            items.append((it, f"{it}   |  {self.item_names.get(it) or '-'}   |  {n}회, 마지막 {last:%m/%d}" if last else it))
        self.f_i.set_items(items)

    def spm_of(self, conn, mc, it, days=90):
        """최근 SPM 찾기: 생산현황판 SPM → 작업지시 ST. days 범위 우선."""
        days = max(1, min(int(days or 90), 730))
        tb = self.g.table()
        cur = conn.cursor()
        old = getattr(conn, "call_timeout", 0)
        try:
            conn.call_timeout = 10000
        except Exception:
            pass
        try:
            if not hasattr(self, "spm_src"):
                self.spm_src = None
                owner = (self.app.cfg.get("schema") or "").strip().upper() or None
                cur.execute("SELECT TABLE_NAME, MAX(CASE WHEN COLUMN_NAME LIKE '%SPM%' THEN COLUMN_NAME END), "
                            "MAX(CASE WHEN COLUMN_NAME = 'WORK_ORDER_NO' THEN 1 END) FROM ALL_TAB_COLUMNS "
                            "WHERE OWNER = NVL(:o, USER) AND TABLE_NAME IN ('ICOM_PRODUCTION_BOARD_WO_HIST', "
                            "'ICOM_PRODUCTION_BOARD_HIST', 'ICOM_PRODUCTION_BOARD') GROUP BY TABLE_NAME", o=owner)
                for t, col, has_wo in cur.fetchall():
                    if col and has_wo:
                        self.spm_src = (self.app.mes.qualify(t), col, t)
                        break
            cav = 1
            try:
                link = self.app.mes.qualify(self.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")
                cur.execute(f"SELECT MAX(CAVITY_QTY) FROM {link} WHERE ITEM_CODE = :i", i=it)
                cav = int(cur.fetchone()[0] or 1) or 1
            except Exception:
                cav = 1
            if not hasattr(self, "spm_ref"):
                self.spm_ref, self.spm_ref_src = spm_table()
            v = self.spm_ref.get(str(it or "").strip().upper())
            if v:                                       # 1순위: SPM 기준표
                return float(v), cav, f"{self.spm_ref_src} {v:g} × 캐비티 {cav}"
            if self.spm_src:
                t, col, tname = self.spm_src
                for cond, label in (("AND W.MACHINE_CODE = :m", "이 설비"), ("", "다른 설비 포함")):
                    b = {"i": it}
                    if cond:
                        b["m"] = mc
                    cur.execute(f"SELECT AVG(V), COUNT(*) FROM (SELECT P.{col} V FROM {t} P WHERE P.WORK_ORDER_NO IN "
                                f"(SELECT W.WORK_ORDER_NO FROM {tb} W WHERE W.ITEM_CODE = :i {cond} "
                                f"AND W.WORK_ORDER_DATE >= SYSDATE - {days}) AND P.{col} > 0)", b)
                    v, n = cur.fetchone()
                    if v:
                        return float(v), cav, f"{tname} 최근 90일 평균 ({label}, {n}건) × 캐비티 {cav}"
            cur.execute(f"SELECT MAX(ST_VALUE) KEEP (DENSE_RANK LAST ORDER BY WORK_ORDER_DATE) FROM {tb} "
                        f"WHERE ITEM_CODE = :i AND MACHINE_CODE = :m AND WORK_ORDER_DATE >= SYSDATE - {days} "
                        f"AND NVL(ST_VALUE, 0) > 0", i=it, m=mc)
            st = cur.fetchone()[0]
            if not st:
                cur.execute(f"SELECT MAX(ST_VALUE) KEEP (DENSE_RANK LAST ORDER BY WORK_ORDER_DATE) FROM {tb} "
                            f"WHERE ITEM_CODE = :i AND WORK_ORDER_DATE >= SYSDATE - {days} AND NVL(ST_VALUE, 0) > 0", i=it)
                st = cur.fetchone()[0]
            if st:
                return 60.0 / float(st), 1, f"최근 {days}일 작업지시 ST {show(st)}초/개 → 60÷ST"
            return None, cav, f"최근 {days}일 SPM 자료 없음 - 계획수량 직접 입력"
        finally:
            try:
                conn.call_timeout = old
            except Exception:
                pass

    def calc_qty(self):
        try:
            spm = float(self.v_spm.get().replace(",", ""))
        except ValueError:
            return
        cav = self.spm_info[1] or 1
        q = spm * 60 * float(self.v_hours.get()) * cav
        self.v_qq.set(str(int(round(q / 10.0) * 10)))      # 10개 단위로 반올림

    def on_item_pick(self, item):
        self.f_i.last_text = item
        """품번을 고르면: 설비 목록을 최근 2년 그 품번을 만든 호기로 거르고,
        지금 설비가 그 안에 없으면(또는 비어 있으면) 가장 최근 호기로 채움 → 최근 SPM × 근무시간으로 수량 계산"""
        item = item.upper()
        hist = self.refresh_machines(item)
        mc = self.cur_mc()
        if hist and not self.v_allmc.get() and mc not in [m for m, _n, _l in hist]:
            mc = hist[0][0]
            self.f_m.set(self.machines.get(mc) or mc)
        elif not hist and not self.f_m.get() and self.item_last.get(item):
            mc = self.item_last[item]
            self.f_m.set(self.machines.get(mc) or mc)
        self.update_spm(item)

    def update_spm(self, item):
        mc = self.cur_mc()
        try:
            self.spm_info = self.app.run_db(lambda c: self.spm_of(c, mc, item))
        except Exception as e:
            self.spm_info = (None, 1, f"SPM 조회 실패: {str(e)[:40]}")
        spm, cav, src = self.spm_info
        self.v_spm.set(f"{spm:.1f}" if spm else "")
        self.lbl_spm.config(text=src[:60])
        if spm:
            self.calc_qty()
        else:
            self.v_qq.set("")
        self.e_qq.focus_set()
        self.e_qq.select_range(0, "end")

    def quick_add(self):
        m, it, q = self.f_m.get(), self.f_i.get().upper(), self.v_qq.get().strip()
        if not (m and it and q):
            messagebox.showinfo(APP_TITLE, "설비명·품번·수량을 모두 넣으세요.", parent=self.win)
            return
        self.rows.append({"line": m, "date_t": "", "shift": "", "item": it, "qty_t": q, "prio": "", "st": "", "lv": ""})
        self.checked = False
        self.show()
        self.f_i.set("")
        self.v_qq.set("")
        self.v_spm.set("")
        self.lbl_spm.config(text="")
        self.refresh_machines("")                       # 다음 줄을 위해 설비 목록 다시 전체로
        self.f_i.ent.focus_set()
        self.lbl.config(text=f"{len(self.rows)}줄  →  다 넣었으면 [검사]")

    def paste(self):
        try:
            text = self.win.clipboard_get()
        except tk.TclError:
            messagebox.showinfo(APP_TITLE, "복사한 내용이 없습니다. 엑셀에서 범위를 복사(Ctrl+C)한 뒤 다시 하세요.", parent=self.win)
            return "break"
        added, skipped = 0, 0
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            if not line.strip():
                continue
            cells = [c.strip() for c in line.split("\t")]
            while cells and cells[-1] == "":
                cells.pop()
            if len(cells) >= 7 and re.fullmatch(r"\d*", cells[0] or "") and not re.search(r"[/.-]", cells[0]):
                if len(cells) == 7 or not self.looks_date(cells[1]):
                    cells = cells[1:]                         # Upload Seq 칸 빼기
            if len(cells) >= 5 and self.looks_date(cells[1]):    # MES 양식: 라인·일자·교대·품번·수량·순위
                cells += [""] * (6 - len(cells))
                self.rows.append({"line": cells[0], "date_t": cells[1], "shift": cells[2], "item": cells[3],
                                  "qty_t": cells[4], "prio": cells[5], "st": "", "lv": ""})
                added += 1
                continue
            if len(cells) >= 3 and re.fullmatch(r"[\d,]+(\.\d+)?", cells[2]):   # 간단 양식: 설비명·품번·수량
                self.rows.append({"line": cells[0], "date_t": "", "shift": "", "item": cells[1],
                                  "qty_t": cells[2], "prio": "", "st": "", "lv": ""})
                added += 1
                continue
            skipped += 1                                         # 제목줄 등
        self.checked = False
        self.show()
        self.lbl.config(text=f"{added}줄 붙여넣음" + (f" ({skipped}줄 건너뜀: 제목줄·빈칸 등)" if skipped else "")
                        + "  →  [검사]를 누르세요.")
        return "break"

    def looks_date(self, t):
        try:
            return self.parse_date(t) is not None
        except Exception:
            return False


    def edit_selected(self):
        """선택한 등록 예정 1줄을 별도 팝업에서 수정. 아직 MES에는 반영하지 않음."""
        sel = self.tv.selection()
        if len(sel) != 1:
            messagebox.showinfo(APP_TITLE, "수정할 줄 1개를 선택하세요.", parent=self.win)
            return
        i = int(sel[0])
        r0 = self.rows[i]
        w = tk.Toplevel(self.win)
        w.title("등록 예정 작업지시 수정 - MES 미반영")
        w.geometry("470x385")
        w.resizable(False, False)
        w.transient(self.win)
        w.grab_set()
        body = ttk.Frame(w, padding=12)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="이 창의 수정은 등록 예정 목록만 바꿉니다. MES는 아직 변경되지 않습니다.",
                  foreground="#C62828").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        fields = [
            ("설비", "line", r0.get("line", "")),
            ("계획일자", "date_t", r0.get("date_t") or self.d_def.var.get()),
            ("교대조", "shift", r0.get("shift") or self.v_shift.get()),
            ("품번", "item", r0.get("item", "")),
            ("계획수량", "qty_t", r0.get("qty_t", "")),
            ("순위", "prio", r0.get("prio", "")),
        ]
        vars_ = {}
        fes = {}
        for rowno, (lab, key, val) in enumerate(fields, 1):
            ttk.Label(body, text=lab, width=12).grid(row=rowno, column=0, sticky="e", padx=(0, 8), pady=5)
            if key in ("line", "item"):          # 설비·품번 = 목록에서 고르기 + 직접 입력 + 글자 치면 걸러짐
                fe = FilterEntry(body, width=31, rows=30)
                fe.set(show(val))
                fe.last_text = show(val)
                fe.grid(row=rowno, column=1, sticky="w", pady=5)
                fes[key] = fe
                vars_[key] = fe.var
                continue
            v = tk.StringVar(value=show(val))
            vars_[key] = v
            ttk.Entry(body, textvariable=v, width=34).grid(row=rowno, column=1, sticky="ew", pady=5)
        fes["line"].set_items([(x, x) for x in self.line_choices("")[0]])      # 전체 설비, 설비명만, 순서대로
        fes["item"].set_items([(x, x) for x in self.item_choices("")[0]])      # 전체 품번, 품번만, 순서대로
        recent = r0.get("recent_date")
        ttk.Label(body, text=f"최근 작업: {recent:%Y-%m-%d}" if recent else "최근 작업: -",
                  foreground="#666").grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(body, text=(f"최근 SPM: {r0.get('spm'):.1f}" if r0.get("spm") else "최근 SPM: 없음"),
                  foreground="#666").grid(row=8, column=0, columnspan=2, sticky="w")

        def save_row():
            item = vars_["item"].get().strip().upper()
            qty = vars_["qty_t"].get().strip()
            line = vars_["line"].get().strip()
            if not line or not item or not qty:
                messagebox.showwarning(APP_TITLE, "설비·품번·계획수량은 비울 수 없습니다.", parent=w)
                return
            try:
                q = float(qty.replace(",", ""))
                if q <= 0:
                    raise ValueError
            except Exception:
                messagebox.showwarning(APP_TITLE, "계획수량은 1 이상의 숫자로 넣으세요.", parent=w)
                return
            r0.update({k: v.get().strip() for k, v in vars_.items()})
            r0["item"] = item
            for k in ("mc", "mname", "iname", "mold", "date", "qty", "shift_v", "prio_auto", "no"):
                r0.pop(k, None)
            r0["st"], r0["lv"] = "수정됨 - 검사 필요", ""
            self.checked = False
            self.show()
            w.destroy()
            self.lbl.config(text="등록 예정 목록을 수정했습니다. [검사]를 다시 하세요.")

        bf = ttk.Frame(body)
        bf.grid(row=9, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(bf, text="저장", style="Big.TButton", command=save_row).pack(side="left", padx=4)
        ttk.Button(bf, text="취소", command=w.destroy).pack(side="left", padx=4)

    def del_rows(self):
        sel = sorted((int(i) for i in self.tv.selection()), reverse=True)
        for i in sel:
            del self.rows[i]
        self.checked = False
        self.show()

    def clear(self):
        if self.rows and not messagebox.askyesno(APP_TITLE, "표를 모두 비울까요?", parent=self.win):
            return
        self.rows.clear()
        self.checked = False
        self.show()
        self.reset_quick()                               # 설비·품번 거르기도 처음부터
        self.lbl.config(text="비웠습니다. 설비부터 다시 고르세요.")

    # ---------- 표 칸 고치기: 목록에서 고르기 ----------
    def mc_code_of(self, text):
        """설비명·설비코드 글자 -> 설비코드 (못 찾으면 '')"""
        t = str(text or "").strip()
        if t in self.machines:
            return t
        n = t.replace(" ", "").upper()
        for c, nm in self.machines.items():
            if nm and nm.replace(" ", "").upper() == n:
                return c
        return ""

    def item_choices(self, line):
        """품번 목록 (글자만, 가나다·ABC 순). return (전체 품번, 이 설비에서 최근 2년 만든 품번)"""
        allv = sorted({k for k, _ in self.item_list if k}, key=str.upper)
        mc = self.mc_code_of(line)
        hist = sorted({it for it, _n, _l in self.mc_items.get(mc, [])}, key=str.upper) if mc else []
        return allv, hist

    def line_choices(self, item):
        """설비 목록 (설비명만, 순서대로). return (전체 설비, 이 품번을 최근 2년 만든 설비)"""
        allv = sorted({k for k, _ in self.mc_all if k}, key=str.upper)
        hist = self.item_mcs.get(str(item or "").strip().upper(), [])
        hv = sorted({self.machines.get(mc) or mc for mc, _n, _l in hist}, key=str.upper)
        return allv, hv

    def ask_pick(self, title, label, allv, hist, init, hist_text):
        """창 안에 목록(30줄)을 바로 보여줌. 글자 치면 걸러짐, 직접 입력도 됨. 고르면 값, 취소면 None"""
        w = tk.Toplevel(self.win)
        w.title(title)
        w.resizable(False, True)
        w.transient(self.win)
        fr = ttk.Frame(w, padding=12)
        fr.pack(fill="both", expand=True)
        ttk.Label(fr, text=label, font=("", 10, "bold")).pack(anchor="w")
        ttk.Label(fr, text="글자를 치면 바로 걸러집니다.  ↓ = 목록으로,  더블클릭/Enter = 고르기,  "
                           "목록에 없으면 직접 입력 후 [확인]", foreground="#666").pack(anchor="w", pady=(2, 6))
        res = {"v": None}
        var = tk.StringVar(value="")             # 기본: 입력칸 비움
        ent = ttk.Entry(fr, textvariable=var, width=36, font=("맑은 고딕", 10))
        ent.pack(fill="x")
        v_hist = tk.BooleanVar(value=bool(hist))   # 기본: 최근 만든 것만 체크
        if hist:
            ttk.Checkbutton(fr, text=f"{hist_text}만 보기 ({len(hist)}개)", variable=v_hist,
                            command=lambda: refill()).pack(anchor="w", pady=(6, 0))
        lf = ttk.Frame(fr)
        lf.pack(fill="both", expand=True, pady=(6, 0))
        lb = tk.Listbox(lf, height=30, width=36, font=("맑은 고딕", 10), activestyle="dotbox", exportselection=False)
        sb = ttk.Scrollbar(lf, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        cnt = ttk.Label(fr, text="", foreground="#666")
        cnt.pack(anchor="w", pady=(4, 0))
        shown = []

        def refill(*_):
            src = hist if v_hist.get() else allv
            t = var.get().strip().upper().replace(" ", "")
            shown[:] = [x for x in src if not t or t in x.upper().replace(" ", "")]
            lb.delete(0, "end")
            for x in shown:
                lb.insert("end", x)
            cur = var.get().strip().upper()
            for i, x in enumerate(shown):
                if x.upper() == cur:
                    lb.selection_set(i)
                    lb.see(i)
                    break
            cnt.config(text=f"{len(shown):,}개 / 전체 {len(src):,}개")

        def ok(v=None):
            res["v"] = var.get().strip() if v is None else v
            w.destroy()

        def pick_sel(_e=None):
            s = lb.curselection()
            if s:
                ok(shown[s[0]])
            return "break"

        def on_enter(_e=None):
            t = var.get().strip().upper()
            if len(shown) == 1:
                ok(shown[0])
            elif any(x.upper() == t for x in shown) or not shown:
                ok()
            else:
                s = lb.curselection()
                ok(shown[s[0]] if s else shown[0])
            return "break"

        def to_list(_e=None):
            if shown:
                lb.focus_set()
                if not lb.curselection():
                    lb.selection_set(0)
                lb.activate(lb.curselection()[0])
            return "break"
        var.trace_add("write", refill)
        ent.bind("<Return>", on_enter)
        ent.bind("<KP_Enter>", on_enter)
        ent.bind("<Down>", to_list)
        lb.bind("<Double-1>", pick_sel)
        lb.bind("<Return>", pick_sel)
        bf = ttk.Frame(fr)
        bf.pack(fill="x", pady=(8, 0))
        def confirm():
            """확인: 목록에서 고른 줄이 있으면 그 값, 없으면 직접 입력한 글자"""
            t = var.get().strip()
            s = lb.curselection()
            if s and (lb_clicked["on"] or not t or t.upper() == shown[s[0]].upper()):
                return ok(shown[s[0]])          # 목록에서 고른 값
            if t and len(shown) == 1:
                return ok(shown[0])             # 글자로 걸러서 1개만 남은 값
            if not t:
                messagebox.showinfo(APP_TITLE, "목록에서 고르거나 값을 입력하세요.", parent=w)
                return
            ok(t)                               # 직접 입력한 값

        lb_clicked = {"on": False}
        lb.bind("<<ListboxSelect>>", lambda e: lb_clicked.__setitem__("on", True))
        ent.bind("<Key>", lambda e: lb_clicked.__setitem__("on", False), add="+")
        ttk.Button(bf, text="확인", command=confirm).pack(side="left", padx=(0, 4))
        ttk.Button(bf, text="취소", command=w.destroy).pack(side="left")
        w.bind("<Escape>", lambda e: w.destroy())
        refill()
        try:
            w.update_idletasks()
            x = self.win.winfo_rootx() + max(0, (self.win.winfo_width() - w.winfo_width()) // 2)
            y = max(0, self.win.winfo_rooty() + max(0, (self.win.winfo_height() - w.winfo_height()) // 2))
            w.geometry(f"+{x}+{y}")
            w.grab_set()
        except Exception:
            pass
        ent.focus_set()
        ent.select_range(0, "end")
        w.wait_window()
        return res["v"]

    def edit_cell(self, e):
        iid, colid = self.tv.identify_row(e.y), self.tv.identify_column(e.x)
        if not iid:
            return
        ci = int(colid[1:]) - 1
        keys = {1: "line", 2: "date_t", 3: "shift", 4: "item", 5: "qty_t", 6: "prio"}   # 일자·교대조·순위는 비우면 기본값
        if ci not in keys:
            return
        r = self.rows[int(iid)]
        if ci == 1:            # 설비: 이 품번을 최근 2년 만든 호기 목록 (없으면 전체)
            allv, hist = self.line_choices(r.get("item", ""))
            v = self.ask_pick("설비", f"설비 고르기 (품번 {r.get('item', '')})", allv, hist,
                              r.get("line", ""), "이 품번을 만든 호기")
        elif ci == 4:          # 품번: 이 설비에서 최근 2년 만든 품번 목록 (없으면 전체)
            allv, hist = self.item_choices(r.get("line", ""))
            v = self.ask_pick("품번", f"품번 고르기 (설비 {r.get('line', '')})", allv, hist,
                              r.get("item", ""), "이 설비에서 만든 품번")
            if v is not None:
                v = v.upper()
        else:
            v = simpledialog.askstring(self.COLS[ci], f"{self.COLS[ci]} 값:", initialvalue=r.get(keys[ci], ""), parent=self.win)
        if v is None:
            return
        if v.strip() == str(r.get(keys[ci], "")).strip():
            return
        r[keys[ci]] = v.strip()
        for k in ("mc", "mname", "iname", "mold"):
            r.pop(k, None)
        r["st"], r["lv"] = "수정됨 - 검사 필요", ""
        r["no"] = ""
        self.checked = False
        self.show()

    # ---------- 표 헤드 클릭 정렬 ----------
    def set_sort_headers(self):
        """헤드 클릭: 같은 헤드는 오름차순 ↔ 내림차순 전환."""
        for i, title in enumerate(self.COLS):
            self.tv.heading(f"c{i}", text=title, command=lambda c=i: self.sort_by(c))

    def sort_by(self, col):
        if self.sort_col == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_col = col
            self.sort_reverse = False
        self.show()

    def sort_raw(self, r, col, original_index):
        """화면에 보이는 값 기준으로 정렬용 원값 반환."""
        if col == 0:
            return original_index + 1
        if col == 1:
            return r.get("line", "")
        if col == 2:
            return r.get("date") or r.get("date_t") or self.d_def.var.get()
        if col == 3:
            return r.get("shift_v") or r.get("shift") or self.v_shift.get()
        if col == 4:
            return r.get("item", "")
        if col == 5:
            return r.get("qty") if r.get("qty") is not None else r.get("qty_t", "")
        if col == 6:
            return r.get("prio") or r.get("prio_auto", "")
        if col == 7:
            return r.get("mc", "")
        if col == 8:
            return r.get("iname", "")
        if col == 9:
            return r.get("mold", "")
        if col == 10:
            return r.get("no", "")
        if col == 11:
            return r.get("st", "")
        if col == 12:
            return r.get("recent_date", "")
        if col == 13:
            return r.get("spm", "")
        return ""

    def sort_key(self, value, col):
        """숫자·날짜·문자를 각 형식에 맞게 정렬."""
        if value is None or str(value).strip() == "":
            return None
        if col in (0, 3, 5, 6, 13):
            try:
                return (0, float(str(value).replace(",", "").strip("() ")))
            except Exception:
                return (1, str(value).casefold())
        if col in (2, 12):
            if isinstance(value, dt.datetime):
                return (0, value.timestamp())
            if isinstance(value, dt.date):
                return (0, dt.datetime.combine(value, dt.time()).timestamp())
            try:
                d = self.parse_date(str(value).strip("() "))
                return (0, d.timestamp())
            except Exception:
                return (1, str(value).casefold())
        return str(value).casefold()

    def show(self):
        self.tv.delete(*self.tv.get_children())

        order = list(range(len(self.rows)))
        if self.sort_col is not None:
            filled, empty = [], []
            for i in order:
                key = self.sort_key(self.sort_raw(self.rows[i], self.sort_col, i), self.sort_col)
                (empty if key is None else filled).append((i, key))
            filled.sort(key=lambda x: x[1], reverse=self.sort_reverse)
            order = [i for i, _k in filled] + [i for i, _k in empty]

        # 현재 정렬 방향을 헤드에 표시
        for c, title in enumerate(self.COLS):
            arrow = ""
            if c == self.sort_col:
                arrow = " ▼" if self.sort_reverse else " ▲"
            self.tv.heading(f"c{c}", text=title + arrow, command=lambda cc=c: self.sort_by(cc))

        # iid는 원래 self.rows 인덱스를 유지해서 수정/삭제/MES 반영 대상이 꼬이지 않도록 함
        for i in order:
            r = self.rows[i]
            self.tv.insert("", "end", iid=str(i), tags=(r["lv"],) if r.get("lv") else (), values=[
                i + 1, r["line"], r["date_t"] or f"({self.d_def.var.get()})", r["shift"] or f"({self.v_shift.get()})",
                r["item"], r["qty_t"], r["prio"] or r.get("prio_auto", ""), r.get("mc", ""),
                r.get("iname", ""), r.get("mold", ""), r.get("no", ""), r.get("st", ""),
                (r.get("recent_date").strftime("%Y-%m-%d") if isinstance(r.get("recent_date"), (dt.datetime, dt.date)) else ""),
                (f"{r.get('spm'):.1f}" if r.get("spm") else "")])

    # ---------- 검사 ----------
    def check(self):
        if not self.rows:
            messagebox.showinfo(APP_TITLE, "붙여넣은 줄이 없습니다.", parent=self.win)
            return False
        tbl = self.g.table()
        link = self.app.mes.qualify(self.app.cfg["item"].get("table") or "ICOM_MOLD_ITEM")
        mst = self.app.mes.qualify(self.app.cfg.get("item_master_table") or "ICOM_ITEM_MASTER")
        bom = self.app.mes.qualify(self.app.cfg.get("bom_table") or "ICOM_ITEM_CHILD")
        try:
            d_def = dt.datetime.combine(self.d_def.get(), dt.time())
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return False
        for r in self.rows:                                   # 형식 검사
            r["lv"], r["st"] = "", ""
            try:
                r["date"] = self.parse_date(r["date_t"]) if r["date_t"] else d_def
                r["shift_v"] = str(r["shift"] or self.v_shift.get() or "1").strip()
            except Exception:
                r["date"], r["lv"], r["st"] = None, "err", "계획일자 형식 오류"
            try:
                q = float(str(r["qty_t"]).replace(",", ""))
                r["qty"] = int(q) if q.is_integer() else q
                if q <= 0:
                    raise ValueError
            except Exception:
                r["qty"], r["lv"], r["st"] = None, "err", (r["st"] + " / " if r["st"] else "") + "계획수량 오류"
            r["item"] = r["item"].strip().upper()

        def work(conn):
            cur = conn.cursor()
            cur.execute(f"SELECT ORGANIZATION_ID FROM (SELECT ORGANIZATION_ID, COUNT(*) N FROM {tbl} "
                        f"WHERE WORK_ORDER_DATE >= SYSDATE - 90 GROUP BY ORGANIZATION_ID ORDER BY N DESC) WHERE ROWNUM = 1")
            org = (cur.fetchone() or [1])[0]
            cur.execute(f"SELECT MACHINE_CODE, MAX(MACHINE_NAME) FROM {tbl} WHERE WORK_ORDER_DATE >= SYSDATE - 400 "
                        f"AND MACHINE_CODE IS NOT NULL GROUP BY MACHINE_CODE")
            machines = {show(a).strip(): show(b).strip() for a, b in cur.fetchall()}
            info = {}
            for it in {r["item"] for r in self.rows if r["item"]}:
                cur.execute(f"SELECT MAX(ITEM_NAME), COUNT(*) FROM {mst} WHERE ITEM_CODE = :i AND ORGANIZATION_ID = :o",
                            i=it, o=org)
                name, n = cur.fetchone()
                cur.execute(f"SELECT MAX(MOLD_CODE) FROM {link} WHERE ITEM_CODE = :i", i=it)
                mold = cur.fetchone()[0]
                cur.execute(f"SELECT COUNT(*) FROM {bom} WHERE ITEM_CODE = :i AND NVL(UNIT_PER_QTY, 0) > 0", i=it)
                b = cur.fetchone()[0]
                info[it] = (n, show(name), show(mold), b)
            dups = {}
            for r in self.rows:
                if r.get("date") and r["item"]:
                    cur.execute(f"SELECT MAX(WORK_ORDER_NO) FROM {tbl} WHERE WORK_ORDER_DATE >= :a AND WORK_ORDER_DATE < :b "
                                f"AND ITEM_CODE = :i", a=r["date"], b=r["date"] + dt.timedelta(days=1), i=r["item"])
                    dups[(r["date"], r["item"])] = show(cur.fetchone()[0])
            return org, machines, info, dups
        try:
            self.org, machines, info, dups = self.app.run_db(work)
        except Exception as e:
            self.app.err(e)
            return False
        machines = {**self.machines, **machines}
        norm = lambda x: re.sub(r"\s+", "", str(x or "")).upper()  # noqa: E731
        by_name = {}
        for code, name in machines.items():
            by_name.setdefault(norm(name), []).append(code)
        by_num = {}
        for code, name in machines.items():
            m = re.search(r"(\d+)\s*호기", name)
            if m:
                by_num.setdefault(int(m.group(1)), []).append(code)
        seen = set()
        for r in self.rows:
            if r["lv"] == "err":
                continue
            msgs, lv = [], "ok"
            ln = str(r["line"]).strip()
            if ln in machines:
                r["mc"] = ln
            elif len(by_name.get(norm(ln), [])) == 1:
                r["mc"] = by_name[norm(ln)][0]
            elif ln.isdigit() and len(by_num.get(int(ln), [])) == 1:
                r["mc"] = by_num[int(ln)][0]
            elif ln.isdigit() and [c for c in machines if c.isdigit() and int(c) == int(ln)]:
                r["mc"] = [c for c in machines if c.isdigit() and int(c) == int(ln)][0]
            else:
                r["mc"] = ""
                msgs.append("설비를 찾을 수 없음 (설비명·설비코드·호기 번호)")
                lv = "err"
            r["mname"] = machines.get(r["mc"], "")
            n, name, mold, b = info.get(r["item"], (0, "", "", 0))
            r["iname"], r["mold"] = name, mold or "*"
            if not n:
                msgs.append("품목 마스터에 없음 → MES 일괄처리에서 빠짐")
                lv = "err"
            if not mold:
                msgs.append("금형 연결 없음 (금형 * 로 들어감)")
                lv = "warn" if lv == "ok" else lv
            if n and not b:
                msgs.append("피치당 소재중량 없음 (소재 사용량 0)")
                lv = "warn" if lv == "ok" else lv
            d = dups.get((r.get("date"), r["item"]))
            if d:
                msgs.append(f"같은 날 이 품번 작업지시 있음 ({d})")
                lv = "dup" if lv == "ok" else lv
            k = (r.get("date"), r["mc"], r["item"])
            if k in seen:
                msgs.append("표 안에 같은 날·라인·품번 줄이 또 있음")
                lv = "warn" if lv == "ok" else lv
            seen.add(k)
            r["lv"], r["st"] = lv, " / ".join(msgs) or "정상"
        cnt = {}                                              # 순위: 비어 있으면 같은 날·설비 안에서 1, 2, 3...
        for r in self.rows:
            k = (r.get("date"), r.get("mc"))
            cnt[k] = cnt.get(k, 0) + 1
            r["prio_auto"] = str(cnt[k])
        self.checked = True
        self.show()
        ne = sum(1 for r in self.rows if r["lv"] == "err")
        nw = sum(1 for r in self.rows if r["lv"] in ("warn", "dup"))
        self.lbl.config(text=f"검사: {len(self.rows)}줄 - 오류 {ne} (빨강, 반영 안 됨) / 확인 {nw} (노랑·하늘)"
                             + ("  →  [W/O 번호생성]" if not ne else "  →  빨간 줄을 고치거나 빼세요"))
        return ne == 0

    # ---------- 번호 ----------
    def make_no(self):
        if not self.checked and not self.check():
            return False
        if any(r["lv"] == "err" for r in self.rows):
            messagebox.showwarning(APP_TITLE, "빨간 줄이 있습니다. 고치거나 뺀 뒤 하세요.", parent=self.win)
            return False
        pr = lambda r: int(r["prio"]) if str(r["prio"]).isdigit() else int(r.get("prio_auto") or 999)  # noqa: E731
        order = sorted(range(len(self.rows)), key=lambda i: (self.rows[i]["date"], self.rows[i]["mc"], pr(self.rows[i]), i))
        tbl = self.g.table()
        try:
            nos, note = self.app.run_db(lambda c: wo_numbers(c, tbl, [self.rows[i]["date"] for i in order]))
        except Exception as e:
            self.app.err(e)
            return False
        for i, no in zip(order, nos):
            self.rows[i]["no"] = no
        self.show()
        self.lbl.config(text=f"작업지시번호 {len(nos)}개 생성 (아직 저장 안 됨)  {('· 규칙: ' + note) if note else '· 같은 날 기존 번호 형식 사용'}"
                             "  →  번호 확인 후 [작업지시 반영]")
        return True

    # ---------- 반영 ----------
    def apply(self):
        if not self.rows:
            return
        if not all(r.get("no") for r in self.rows) and not self.make_no():
            return
        if any(r["lv"] == "err" for r in self.rows):
            return
        nw = sum(1 for r in self.rows if r["lv"] in ("warn", "dup"))
        days = sorted({r["date"].strftime("%Y-%m-%d") for r in self.rows})
        mode_line = (f"  근무기준: {hours_text(self.recent_hours)}\n"
                     if self.recent_hours else "")
        if not messagebox.askyesno("작업지시 반영", f"작업지시 {len(self.rows)}건을 MES에 반영합니다.\n\n"
                                               + mode_line
                                               + f"  계획일자: {', '.join(days[:5])}{' 외' if len(days) > 5 else ''}\n"
                                               f"  번호: {min((r['no'] for r in self.rows), key=lambda x: int(WO_RE.match(x).group(2)))} ~ "
                                               f"{max((r['no'] for r in self.rows), key=lambda x: int(WO_RE.match(x).group(2)))}\n"
                                               + (f"  ※ 확인 필요 줄 {nw}개 (노랑·하늘)\n" if nw else "")
                                               + f"\nMES '작업지시일괄처리'와 같은 방법(업로더 표 + {self.PROC})으로 넣습니다.\n"
                                                 f"기존 작업지시는 바뀌지 않습니다. 계속할까요?", parent=self.win,
                                   icon="warning" if nw else "question"):
            return
        tbl = self.g.table()
        up = self.app.mes.qualify(self.T_UP)
        user = (self.app.cfg.get("reg_user") or "").strip() or "PYBATCH"
        now = dt.datetime.now()

        def work(conn):
            cur = conn.cursor()
            nos = [r["no"] for r in self.rows]
            for no in nos:                                    # 번호가 그 사이 쓰였는지 다시 확인
                cur.execute(f"SELECT COUNT(*) FROM {tbl} WHERE WORK_ORDER_NO = :n", n=no)
                if cur.fetchone()[0]:
                    raise ValueError(f"작업지시번호 {no} 가 그 사이 생겼습니다. [W/O 번호생성]을 다시 하세요.")
            meta = table_meta(conn, up)
            cur.execute(f"SELECT NVL(MAX(SESSION_ID), 0) + 1 FROM {up}")
            sid = int(cur.fetchone()[0])
            try:
                for k, r in enumerate(self.rows, 1):
                    insert_dict(cur, up, meta, {
                        "SESSION_ID": sid, "UPLOAD_SEQ": k, "ORGANIZATION_ID": self.org,
                        "MACHINE_CODE": r["mc"], "MACHINE_NAME": r.get("mname") or None,
                        "WORK_ORDER_DATE": r["date"], "WORK_ORDER_NO": r["no"], "ITEM_CODE": r["item"],
                        "PLAN_QTY": r["qty"],
                        "PLAN_PRIORITY": int(r["prio"]) if str(r["prio"]).isdigit() else int(r.get("prio_auto") or k),
                        "WORK_SHIFT": r.get("shift_v") or "1", "STATUS_FLAG": "Y",
                        "COMMENTS": (f"python 최근30일 {self.recent_hours:g}h 계획" if self.recent_hours else "python 일괄등록"),
                        "ENTER_DATE": now, "ENTER_BY": user, "LAST_MODIFY_DATE": now, "LAST_MODIFY_BY": user})
                out, msg = cur.var(str), cur.var(str)
                cur.callproc(self.PROC, [self.org, sid, out, msg])   # MES 프로그램: 작업지시 생성 + COMMIT
                res = (out.getvalue(), msg.getvalue())
            except Exception:
                conn.rollback()
                raise
            if res[0] != "OK":
                conn.rollback()
                raise RuntimeError(f"MES 일괄처리 실패: {res[1]}")
            got = []
            for no in nos:
                cur.execute(f"SELECT MOLD_CODE, MACHINE_NAME FROM {tbl} WHERE WORK_ORDER_NO = :n", n=no)
                got.append(cur.fetchone())
            return sid, got
        try:
            sid, got = self.app.run_db(work)
        except ValueError as e:
            messagebox.showwarning(APP_TITLE, str(e), parent=self.win)
            return
        except Exception as e:
            messagebox.showerror(APP_TITLE, ora_hint(e), parent=self.win)
            return
        ok = 0
        for r, gt in zip(self.rows, got):
            if gt:
                ok += 1
                r["lv"], r["st"] = "ok", f"반영됨 (금형 {show(gt[0])})"
            else:
                r["lv"], r["st"] = "err", "반영 안 됨 (품목 마스터 조직 확인)"
        self.show()
        self.app.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", user, "작업지시일괄", tbl, f"session {sid}",
                              "; ".join(f"{r['no']} {r['mc']} {r['item']} {r['qty']}" for r in self.rows)]])
        self.app.load_log()
        self.lbl.config(text=f"반영 완료 {ok}/{len(self.rows)}건 (업로더 session {sid}).  현장 화면에 보이는지 확인하세요.")
        messagebox.showinfo(APP_TITLE, f"작업지시 {ok}/{len(self.rows)}건 반영했습니다.", parent=self.win)
        try:
            self.g.fetch(ask=False)
        except Exception:
            pass


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_cfg()
        self.mes = Mes(self)
        self.xl = ExcelIO(self.mes)
        self.path = None
        self.data = {p: {"rows": [], "info": None, "results": []} for p in PARTS}
        self.old_cols, self.old_rows, self.old_time = [], [], None
        self.item_cols, self.item_rows, self.item_code = [], [], None
        self.old_defs, self.item_defs = {}, {}
        self.edits = {}          # (p, 키) -> 수정 대기 내용
        self._editor = None

        root.title(APP_TITLE)
        root.geometry("1400x860")
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Treeview", rowheight=24)
        self.setup_styles(st)

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.nb = nb
        self.tabs, self.pages = {}, {}
        for key, title in self.TAB_TITLES:
            page = ttk.Frame(nb)
            nb.add(page, text=f" {title} ")
            self.make_banner(page, key)
            body = ttk.Frame(page)
            body.pack(fill="both", expand=True)
            self.pages[key], self.tabs[key] = page, body

        self.quick = QuickTab(self, self.tabs["quick"])
        self.build_new()
        self.build_old()
        self.grid_item = GridTab(self, self.tabs["itemmst"], "itemmst")
        self.grid_bom = GridTab(self, self.tabs["bom"], "bom")
        self.grid_wo = GridTab(self, self.tabs["wo"], "wo")
        self.rcv = ReceiptTab(self, self.tabs["rcv"])
        self.spm = SpmTab(self, self.tabs["spm"])
        self.inspect = InspectTabs(self, self.tabs["find"], self.tabs["diff"])
        self.build_map()
        self.build_set()
        self.build_log()
        self.build_help()
        self.apply_dev_tabs()
        self.show_tab("old")
        nb.bind("<<NotebookTabChanged>>", self.on_tab_changed, add="+")

        if openpyxl is None or oracledb is None:
            how = ("exe를 다시 만들어야 합니다 (빠진 부품 있음)" if getattr(sys, "frozen", False)
                   else "명령 프롬프트에서:  python -m pip install oracledb openpyxl")
            messagebox.showwarning(APP_TITLE, "필요한 부품을 불러오지 못했습니다.\n\n" + "\n".join(IMPORT_ERR)
                                   + f"\n\n{how}")
        pass

    # ---------- 화면 꾸미기 (탭 이름 / 버튼 색 / 페이지 안내) ----------
    TAB_TITLES = (("quick", "금형 간편등록"), ("new", "엑셀 대량등록"), ("old", "① 금형·품번 관리"),
                  ("itemmst", "품목 관리"), ("bom", "② 단위중량 관리"), ("wo", "③ 작업지시 관리"),
                  ("rcv", "④ 소재 관리"), ("spm", "⑤ SPM 관리"),
                  ("find", "MES 구조조사"), ("diff", "전후 비교"),
                  ("map", "엑셀칸 설정"), ("set", "환경설정"), ("log", "작업기록"), ("help", "사용법"))

    # key: (제목, 설명, 사용 순서)
    PAGE_INFO = {
        "quick": ("새 금형 1개를 화면에서 바로 등록",
                  "금형 + 품번 연결 + 단위중량을 한 번에 넣습니다. 기존 MES 내용은 절대 바뀌지 않고 새 줄만 추가됩니다.",
                  "① 비슷한 기준 금형 [찾기]  →  ② 새 금형코드·금형명 입력  →  ③ [+ 품번 추가]  →  ④ [검사]  →  ⑤ [MES에 등록]"),
        "new": ("엑셀로 금형·품번 여러 개를 한꺼번에 등록",
                "엑셀 양식에 적어서 불러온 뒤 검사하고 등록합니다. 기존 내용은 바뀌지 않습니다 (추가만).",
                "[엑셀 양식 만들기]  →  엑셀에 입력·저장  →  [엑셀 불러오기]  →  [① 검사] (빨강은 고치기)  →  [② MES 등록]"),
        "old": ("금형과 연결 품번 보기 · 고치기 · 새 금형으로 복사 · 잘못 등록한 것 삭제",
                "위 표 = 금형, 아래 표 = 선택한 금형의 품번.  고치기·삭제는 [수정 모드], 새 금형코드로 복사는 [다른 이름으로 저장 모드] (원본은 그대로).",
                "고치기: [MES에서 조회] → [수정 모드] → 칸 더블클릭 → [수정 저장]   /   복사: 금형 ☑ → [다른 이름으로 저장 모드] → "
                "금형코드·금형명·품번 고치기 → [다른 이름으로 저장]   /   삭제: ☑ 체크 → [선택 금형 삭제]·[선택 품번줄 삭제]"),
        "itemmst": ("품목(품번) 기본정보 보기 · 수정 · 오타 삭제",
                    "빨강 줄 = 품번이 날짜 모양(오타 의심). 작업지시·실적에 쓰인 품번은 삭제가 막힙니다.",
                    "[MES에서 조회]  →  [수정 모드] 체크  →  칸 더블클릭해서 고치기  →  [수정 저장]   /   새 품목: [새로 등록]"),
        "bom": ("피치당 소재중량(g) 관리 = 1샷에 코일에서 빠지는 무게 (스크랩 포함)",
                "코일 사용량 = 샷수 × 캐비티 × 이 값 으로 계산됩니다. 제품 순수무게가 아니라 두께×폭×피치×7.85÷1000 (g). "
                "캐비티 2 이상이면 ÷캐비티 값을 넣으세요. 0이면 소재가 차감되지 않습니다.",
                "탭을 열면 전체 조회 (품번 넣고 [MES에서 조회] = 그 품번만)  →  [수정 모드]  →  칸 더블클릭  →  [수정 저장]   /   "
                "없는 품번: [새로 등록]   /   많을 때: [엑셀로 저장] → [엑셀로 일괄 수정]"),
        "wo": ("작업지시 새로 만들기 · 고치기 · 확인 · 잘못 낸 것 삭제",
               "빨강 = 금형 없음(*)  ·  노랑 = 품번의 금형과 다름  ·  하늘 = 같은 날 중복. 실적이 있는 작업지시는 삭제할 수 없습니다.",
               "새로: [최근 작업지시 불러오기] 호기별 최근 품번 / [여러 건 일괄등록] 설비명·품번·수량만   /   고치기: [MES에서 조회] → 줄 더블클릭 → [수정 저장]"
               "   /   삭제: [수정 모드] → [선택 줄 삭제]"),
        "spm": ("품번별 SPM(분당 타수) 관리",
                "[③ 작업지시 관리] > 최근 작업지시 불러오기·일괄등록에서 계획수량 = SPM × 60 × 근무시간 × 캐비티 로 계산할 때 씁니다. "
                "프로그램 폴더의 spm.xlsx 에 저장되며 MES는 바뀌지 않습니다. 없는 품번은 최근 생산 기록으로 계산합니다.",
                "[+ 새로 등록] 품번·SPM 입력   /   고치기: 줄 더블클릭   /   빠진 것: [SPM 없는 품번 찾기] → 더블클릭해서 등록"),
        "rcv": ("소재(코일) 입고 조회 · 새 입고 · 고치기 · 잘못 입고한 것 삭제",
                "입고할 때 입고 + 재고LOT + 소재 재고가 한 번에 맞춰집니다. 생산에 사용·출고된 소재는 업체·담당자만 고칠 수 있고 삭제는 막힙니다.",
                "새 입고: [+ 새 입고] → 품번·HEAT·업체·중량 → [MES에 등록]   /   고치기: 줄 더블클릭 → [수정 저장]   /   "
                "삭제: [삭제 모드] → 초록 줄 선택 → [선택 입고 삭제]"),
        "find": ("MES 안쪽 구조 살펴보기 (조회 전용)",
                 "작업지시·금형과 관련된 테이블, 프로그램, 트리거를 찾아봅니다. MES 내용은 전혀 바뀌지 않습니다.",
                 "① 관련 테이블 → ② 작업지시 만드는 프로그램 → ③ 번호 생성기 → ④ 트리거 순서로 눌러 보기   /   프로그램 내용 검색: 글자 입력 → [검색]"),
        "diff": ("MES 화면에서 작업한 전과 후를 비교 (조회 전용)",
                 "MES가 실제로 어떤 테이블의 어떤 칸을 바꾸는지 확인합니다. MES 내용은 바뀌지 않습니다.",
                 "[① 찍기 전 저장]  →  MES 화면에서 작업(예: 작업지시 1건 등록)  →  [③ 찍은 후 비교]  →  바뀐 줄 확인"),
        "map": ("엑셀 대량등록 양식에 나올 칸 정하기",
                "엑셀 대량등록(숨김 탭)에서만 쓰는 설정이라 보통은 손댈 필요가 없습니다. 엑셀에 입력하고 싶은 칸만 '사용 Y'로 둡니다.",
                "[DB컬럼 가져오기]  →  사용 칸 더블클릭(Y/N)  →  [저장]"),
        "set": ("접속 정보와 기본값 설정",
                "접속 계정과 금형·품번 테이블은 고정되어 있습니다. 아래 체크칸으로 숨겨 둔 탭(금형 등록·품목 관리·고급)을 켤 수 있습니다.",
                "[연결 테스트]  →  [등록자]에 이름 입력  →  [저장]"),
        "log": ("이 프로그램으로 등록·수정·삭제한 기록",
                "모든 작업이 mold_register_log.csv 파일에도 남습니다. 삭제한 줄은 mold_deleted_backup.jsonl 에 백업됩니다.",
                "[새로고침]으로 최신 기록 보기   /   [파일 열기]로 기록 파일(엑셀) 열기"),
        "help": ("전체 사용법",
                 "버튼 색: 초록=MES 등록·저장, 파랑=조회·검사, 청록=엑셀, 빨강=삭제.",
                 ""),
    }

    BTN_COLORS = {  # 스타일: (평소색, 마우스 올렸을 때)
        "Go": ("#2E7D32", "#1B5E20"),       # MES에 등록·저장
        "Primary": ("#1565C0", "#0D47A1"),  # 조회·검사·찾기
        "Excel": ("#00838F", "#006064"),    # 엑셀
        "Danger": ("#C62828", "#8E0000"),   # 삭제
        "Recent": ("#F57C00", "#E65100"),   # 최근 작업지시 불러오기
    }
    BTN_RULES = (  # 버튼 글자에 이 말이 있으면 해당 색 (위에서부터 먼저 맞는 것)
        ("Recent", ("최근 작업지시 불러오기",)),
        ("Danger", ("삭제",)),
        ("Go", ("MES에 등록", "MES 등록", "수정 저장", "새로 등록", "새 작업지시", "새 입고", "다른 이름으로 저장", "작업지시 반영", "일괄등록", "+ 추가")),
        ("Excel", ("엑셀",)),
        ("Primary", ("조회", "검사", "찾기", "검색", "연결 테스트", "찍", "관련 테이블", "프로그램",
                     "번호 생성기", "트리거", "미리보기", "새로고침", "DB컬럼 가져오기", "고치기", "W/O 번호생성")),
    )

    def setup_styles(self, st):
        st.configure("Big.TButton", padding=(10, 5))
        st.configure("TNotebook.Tab", padding=(10, 5), font=("맑은 고딕", 10, "bold"))
        st.map("TNotebook.Tab", background=[("selected", "#1F4E79")], foreground=[("selected", "#FFFFFF")])
        for name, (bg, dark) in self.BTN_COLORS.items():
            for pre, pad in (("", (8, 3)), ("Big.", (12, 5))):
                sty = f"{name}.{pre}TButton" if pre else f"{name}.TButton"
                st.configure(sty, background=bg, foreground="#FFFFFF", bordercolor=dark,
                             lightcolor=bg, darkcolor=dark, padding=pad, font=("맑은 고딕", 9, "bold"))
                st.map(sty, background=[("disabled", "#B0BEC5"), ("pressed", dark), ("active", dark)],
                       foreground=[("disabled", "#ECEFF1")])
        # 화면에 나타나는 모든 버튼에 글자 보고 자동으로 색 입히기 (팝업창 포함)
        self.root.bind_class("TButton", "<Map>", self.color_button, add="+")

    def color_button(self, ev):
        w = ev.widget
        try:
            cur = str(w.cget("style"))
            if cur not in ("", "TButton", "Big.TButton"):
                return
            text = str(w.cget("text"))
        except Exception:
            return
        for name, words in self.BTN_RULES:
            if any(k in text for k in words):
                w.configure(style=f"{name}.Big.TButton" if cur == "Big.TButton" else f"{name}.TButton")
                return

    def make_banner(self, page, key):
        info = self.PAGE_INFO.get(key)
        if not info:
            return
        title, desc, steps = info
        bg = "#EEF4FB"
        box = tk.Frame(page, bg=bg, highlightbackground="#B7CCE4", highlightthickness=1)
        box.pack(fill="x", padx=6, pady=(6, 2))
        left = tk.Frame(box, bg=bg)
        left.pack(side="left", fill="x", expand=True, padx=10, pady=6)
        tk.Label(left, text=title, bg=bg, fg="#1F4E79", font=("맑은 고딕", 11, "bold"),
                 anchor="w").pack(fill="x")
        tk.Label(left, text=desc, bg=bg, fg="#333333", font=("맑은 고딕", 9),
                 anchor="w", justify="left", wraplength=1150).pack(fill="x")
        if steps:
            tk.Label(left, text="사용 순서:  " + steps, bg=bg, fg="#2E7D32", font=("맑은 고딕", 9, "bold"),
                     anchor="w", justify="left", wraplength=1150).pack(fill="x", pady=(2, 0))
        if key != "help":
            ttk.Button(box, text="자세한 사용법", command=lambda: self.show_tab("help")).pack(
                side="right", padx=10)

    def on_tab_changed(self, _e=None):
        """처음 여는 탭은 자동으로 전체 조회 (단위중량 등)"""
        try:
            cur = self.nb.select()
        except Exception:
            return
        for key, g in (("bom", getattr(self, "grid_bom", None)),):
            if g and str(self.pages.get(key)) == str(cur) and g.spec.get("auto_fetch") and g.fetched is None:
                self.root.after(50, lambda g=g: g.fetch(ask=False))

    def show_tab(self, key):
        self.nb.select(self.pages[key])

    DEV_TABS = ("find", "diff")
    REG_TABS = ("quick", "new", "map")
    ITEM_TABS = ("itemmst",)

    def apply_dev_tabs(self):
        """숨김 탭: 고급(MES 구조조사 / 전후 비교), 금형 등록(간편등록 / 엑셀 대량등록) - 환경설정에서 켜면 보임"""
        for keys, on in ((self.DEV_TABS, bool(self.cfg.get("show_dev_tabs"))),
                         (self.REG_TABS, bool(self.cfg.get("show_reg_tabs"))),
                         (self.ITEM_TABS, bool(self.cfg.get("show_item_tab")))):
            for k in keys:
                if on:
                    self.nb.add(self.pages[k])      # 숨겨진 탭은 원래 자리로 돌아옴
                else:
                    self.nb.hide(self.pages[k])
        if hasattr(self, "btn_copy_old"):
            if self.cfg.get("show_reg_tabs"):
                self.btn_copy_old.pack(side="right", padx=3)
            else:
                self.btn_copy_old.pack_forget()
        try:
            if self.nb.tab(self.nb.select(), "state") == "hidden":
                self.show_tab("old")
        except Exception:
            self.show_tab("old")

    # ---------- 공통 ----------
    def busy(self, on):
        self.root.config(cursor="watch" if on else "")
        self.root.update_idletasks()

    def run_db(self, fn):
        conn = None
        self.busy(True)
        try:
            conn = self.mes.connect()
            return fn(conn)
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
            self.busy(False)

    def run_bg(self, fn, title="확인 중"):
        """fn(conn, progress, cancelled) 를 뒤에서 실행하고, 그동안 진행 창을 띄움 (프로그램이 멈추지 않음)"""
        import threading
        import time
        box = {"msg": "MES 접속 중...", "res": None, "err": None, "stop": False}
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("520x140")
        win.transient(self.root)
        win.resizable(False, False)
        lbl = ttk.Label(win, text=box["msg"], wraplength=490, justify="left")
        lbl.pack(fill="x", padx=14, pady=(16, 8))
        pb = ttk.Progressbar(win, mode="indeterminate")
        pb.pack(fill="x", padx=14)
        pb.start(12)

        def stop():
            box["stop"] = True
            lbl.config(text="취소하는 중... (지금 확인 중인 표가 끝나면 멈춥니다)")
        ttk.Button(win, text="취소", command=stop).pack(pady=8)
        win.protocol("WM_DELETE_WINDOW", stop)

        def work():
            conn = None
            try:
                conn = self.mes.connect()
                box["res"] = fn(conn, lambda m: box.__setitem__("msg", m), lambda: box["stop"])
            except Exception as e:
                box["err"] = e
            finally:
                if conn is not None:
                    try:
                        conn.close()
                    except Exception:
                        pass
        th = threading.Thread(target=work, daemon=True)
        th.start()
        try:
            win.grab_set()
        except Exception:
            pass
        while th.is_alive():
            if not box["stop"]:
                lbl.config(text=box["msg"])
            self.root.update()
            time.sleep(0.05)
        try:
            win.grab_release()
        except Exception:
            pass
        win.destroy()
        if box["err"] is not None:
            raise box["err"]
        if box["stop"]:
            raise RuntimeError("취소했습니다. (MES는 그대로)")
        return box["res"]

    def err(self, e):
        self.busy(False)
        messagebox.showerror(APP_TITLE, ora_hint(e))

    @staticmethod
    def make_tree(parent, cols, widths=None):
        fr = ttk.Frame(parent)
        ids = [f"c{i}" for i in range(len(cols))]
        tv = ttk.Treeview(fr, columns=ids, show="headings", selectmode="extended")
        for i, (cid, c) in enumerate(zip(ids, cols)):
            tv.heading(cid, text=c)
            tv.column(cid, width=(widths[i] if widths else 110), anchor="w", stretch=False)
        ys = ttk.Scrollbar(fr, orient="vertical", command=tv.yview)
        xs = ttk.Scrollbar(fr, orient="horizontal", command=tv.xview)
        tv.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        tv.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        fr.rowconfigure(0, weight=1)
        fr.columnconfigure(0, weight=1)
        tv.tag_configure("err", background="#FFC7CE")
        tv.tag_configure("warn", background="#FFEB9C")
        tv.tag_configure("ok", background="#C6EFCE")
        tv.tag_configure("off", foreground="#999999")
        tv.tag_configure("edit", background="#FCE4D6")
        return fr, tv

    @staticmethod
    def reset_tree(tv, cols, widths=None):
        tv.delete(*tv.get_children())
        ids = [f"c{i}" for i in range(len(cols))]
        tv["columns"] = ids
        for i, (cid, c) in enumerate(zip(ids, cols)):
            tv.heading(cid, text=c)
            tv.column(cid, width=(widths[i] if widths else 110), anchor="w", stretch=False)

    @staticmethod
    def col_widths(names):
        return [max(80, min(220, blen(n, 2) * 9 + 24)) for n in names]

    def ask_save(self, name):
        return filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=name,
                                            filetypes=[("Excel", "*.xlsx")])

    def connect_optional(self):
        try:
            self.busy(True)
            return self.mes.connect()
        except Exception:
            return None
        finally:
            self.busy(False)

    # ==========================================================
    # 신규등록 탭
    # ==========================================================
    def build_new(self):
        t = self.tabs["new"]
        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=(6, 4))
        ttk.Button(bar, text="엑셀 양식 만들기", style="Big.TButton", command=self.on_template).pack(side="left", padx=3)
        ttk.Button(bar, text="엑셀 불러오기", style="Big.TButton", command=self.on_load).pack(side="left", padx=3)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(bar, text="① 검사", style="Big.TButton", command=self.on_check).pack(side="left", padx=3)
        ttk.Button(bar, text="② MES 등록", style="Big.TButton", command=self.on_register).pack(side="left", padx=3)
        ttk.Button(bar, text="엑셀 열기", command=lambda: self.path and open_file(self.path)).pack(side="right", padx=3)
        self.var_allow = tk.BooleanVar(value=bool(self.cfg.get("allow_item_to_existing")))

        def _allow():
            self.cfg["allow_item_to_existing"] = self.var_allow.get()
            save_cfg(self.cfg)
        ttk.Checkbutton(bar, text="기존 금형 품번추가 허용", variable=self.var_allow,
                        command=_allow).pack(side="right", padx=10)
        self.lbl_file = ttk.Label(t, text="불러온 파일 없음", foreground="#555")
        self.lbl_file.pack(fill="x", padx=4)

        pw = ttk.PanedWindow(t, orient="vertical")
        pw.pack(fill="both", expand=True, pady=4)
        self.tv_new, self.lbl_part = {}, {}
        for p in PARTS:
            f = ttk.Frame(pw)
            lbl = ttk.Label(f, text=f"[{PART_NAME[p]} 시트]", font=("", 10, "bold"))
            lbl.pack(fill="x", padx=2, pady=(2, 2))
            fr, tv = self.make_tree(f, ["행"])
            fr.pack(fill="both", expand=True)
            tv.bind("<Double-1>", lambda e, p=p: self.on_new_detail(p))
            pw.add(f, weight=3 if p == "mold" else 2)
            self.tv_new[p], self.lbl_part[p] = tv, lbl
        self.lbl_sum = ttk.Label(t, text="", font=("", 10, "bold"))
        self.lbl_sum.pack(fill="x", padx=4, pady=(0, 4))
        self.show_new()

    def show_new(self):
        total_err = 0
        parts_txt = []
        for p in PARTS:
            tv, d = self.tv_new[p], self.data[p]
            if not self.mes.enabled(p):
                self.reset_tree(tv, ["(품번 테이블 사용 안 함)"], [300])
                self.lbl_part[p].config(text=f"[{PART_NAME[p]} 시트] - 사용 안 함")
                continue
            used = self.mes.used(p)
            names = [c["name"] for c in used]
            self.reset_tree(tv, ["행"] + names + ["검사결과"], [45] + self.col_widths(names) + [520])
            res_by = {r["rno"]: r for r in d["results"]}
            n = {"OK": 0, "경고": 0, "오류": 0}
            for rno, row in d["rows"]:
                r = res_by.get(rno)
                vals = [rno] + [cell_text(row.get(c["col"])) for c in used] + [r["msg"] if r else ""]
                tag = {"오류": "err", "경고": "warn", "OK": "ok"}.get(r["status"]) if r else ""
                if r:
                    n[r["status"]] += 1
                tv.insert("", "end", iid=str(rno), values=vals, tags=(tag,) if tag else ())
            tbl = self.cfg[p].get("table", "")
            if d["results"]:
                txt = f"{PART_NAME[p]} {len(d['rows'])}건 (정상 {n['OK']} / 경고 {n['경고']} / 오류 {n['오류']})"
            else:
                txt = f"{PART_NAME[p]} {len(d['rows'])}건"
            self.lbl_part[p].config(text=f"[{PART_NAME[p]} 시트 -> {tbl}]   {txt}")
            parts_txt.append(txt)
            total_err += n["오류"]
        checked = any(self.data[p]["results"] for p in PARTS)
        self.lbl_sum.config(text=("   |   ".join(parts_txt) + ("" if checked else "   (검사 전)")),
                            foreground="#C00000" if total_err else ("#1F6E1F" if checked else "#333"))

    def on_template(self):
        try:
            self.mes.validate_map()
        except Exception as e:
            self.err(e)
            return
        path = self.ask_save(f"금형등록_{dt.datetime.now():%Y%m%d}.xlsx")
        if not path:
            return
        conn = self.connect_optional()
        if conn is None and not messagebox.askyesno(APP_TITLE, "MES에 접속하지 못했습니다.\n드롭다운 목록 없이 양식만 만들까요?"):
            return
        try:
            self.busy(True)
            self.xl.make_template(path, conn=conn)
            self.busy(False)
            open_file(path)
        except Exception as e:
            self.err(e)
        finally:
            if conn is not None:
                conn.close()

    def on_load(self, path=None, quiet=False):
        path = path or filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return False
        try:
            loaded = self.xl.load(path)
        except Exception as e:
            self.err(e)
            return False
        self.path = path
        for p in PARTS:
            rows, info = loaded[p]
            self.data[p] = {"rows": rows, "info": info, "results": []}
        sheets = ", ".join(self.data[p]["info"]["sheet"] for p in PARTS if self.data[p]["info"])
        self.lbl_file.config(text=f"{path}   ({sheets})")
        self.show_new()
        if not quiet and not any(self.data[p]["rows"] for p in PARTS):
            messagebox.showinfo(APP_TITLE, "입력된 내용이 없습니다.")
        return True

    def reload(self):
        if not self.path:
            messagebox.showinfo(APP_TITLE, "먼저 [엑셀 불러오기]를 하세요.")
            return False
        if not self.on_load(self.path, quiet=True):
            return False
        if not any(self.data[p]["rows"] for p in PARTS):
            messagebox.showinfo(APP_TITLE, "입력된 내용이 없습니다.")
            return False
        return True

    def run_check(self, conn):
        mres, ires = self.mes.check(conn, self.data["mold"]["rows"], self.data["item"]["rows"])
        self.data["mold"]["results"], self.data["item"]["results"] = mres, ires
        return mres, ires

    def write_back(self, done_text=None):
        return self.xl.write_results(self.path, [(self.data[p]["info"], self.data[p]["results"]) for p in PARTS],
                                     done_text)

    def counts(self):
        out = {}
        for p in PARTS:
            res = self.data[p]["results"]
            out[p] = {s: sum(r["status"] == s for r in res) for s in ("OK", "경고", "오류")}
        return out

    def on_check(self):
        if not self.reload():
            return
        try:
            self.run_db(self.run_check)
            self.show_new()
            saved = self.write_back()
        except Exception as e:
            self.err(e)
            return
        n = self.counts()
        nerr = sum(n[p]["오류"] for p in PARTS)
        msg = "검사 완료\n\n" + "\n".join(
            f"{PART_NAME[p]}: 정상 {n[p]['OK']} / 경고 {n[p]['경고']} / 오류 {n[p]['오류']}"
            for p in PARTS if self.mes.enabled(p))
        if saved != self.path:
            msg += f"\n\n엑셀이 열려 있어 결과를 다른 파일에 저장했습니다:\n{saved}"
        elif nerr:
            msg += "\n\n엑셀 파일에 빨간 칸으로 표시했습니다."
        (messagebox.showwarning if nerr else messagebox.showinfo)(APP_TITLE, msg)

    def on_register(self):
        if not self.reload():
            return
        conn = None
        try:
            self.busy(True)
            conn = self.mes.connect()
            mres, ires = self.run_check(conn)
            self.show_new()
            self.busy(False)
            n = self.counts()
            nerr = sum(n[p]["오류"] for p in PARTS)
            nwarn = sum(n[p]["경고"] for p in PARTS)
            if nerr:
                self.write_back()
                messagebox.showwarning(APP_TITLE, f"오류 {nerr}건이 있어 등록하지 않았습니다.\n빨간 칸을 고친 뒤 다시 하세요.")
                return
            msg = f"금형 {len(mres)}건 -> {self.mes.table('mold')}"
            if self.mes.enabled("item"):
                msg += f"\n품번 {len(ires)}건 -> {self.mes.table('item')}"
            ref = self.cfg.get("ref_mold") or ""
            for p in PARTS:
                auto = self.mes.last_auto.get(p) or {}
                if auto and (self.data[p]["results"]):
                    how = {"SYSDATE": "현재시각", "{USER}": "등록자", "{COPY}": f"기준금형 {ref} 값"}
                    msg += f"\n\n[{PART_NAME[p]}] 자동으로 채우는 필수 컬럼:\n  " + \
                           "\n  ".join(f"{c} = {how[h]}" for c, h in list(auto.items())[:10])
                    if len(auto) > 10:
                        msg += f"\n  ... 외 {len(auto) - 10}개"
            trg = self.mes.triggers(conn)
            tl = [f"{PART_NAME[p]}: " + ", ".join(v) for p, v in trg.items() if v]
            if tl:
                msg += "\n\n※ 이 테이블에는 MES 자체 자동동작(트리거)이 있습니다:\n  " + "\n  ".join(tl)
            msg += "\n\n새 줄만 추가(INSERT)합니다. 기존 금형·품번 줄은 바꾸지 않습니다."
            msg += "\n위 내용을 MES에 등록합니다."
            if nwarn:
                msg += f"\n※ 경고 {nwarn}건 포함 (노란 줄 확인)"
            msg += "\n\n계속할까요?"
            if not messagebox.askyesno("MES 등록", msg):
                return
            self.busy(True)
            done = self.mes.register(conn, mres, ires)
            self.busy(False)
        except Exception as e:
            self.err(e)
            return
        finally:
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass

        now = dt.datetime.now()
        self.write_log(done, now)
        try:
            saved = self.write_back(done_text=f"등록완료 {now:%Y-%m-%d %H:%M}")
        except Exception:
            saved = None
        for p in PARTS:
            for r in self.data[p]["results"]:
                r["status"], r["msg"] = "OK", f"등록완료 {now:%H:%M}"
        self.show_new()
        self.load_log()
        nm = sum(1 for d in done if d[0] == "mold")
        ni = len(done) - nm
        ok, miss = getattr(self.mes, "last_verify", (len(done), []))
        msg = f"등록 완료: 금형 {nm}건, 품번 {ni}건\n[작업기록] 탭에 기록했습니다.\n\nMES에서 다시 읽어 확인: {ok}/{len(done)}건 확인됨"
        if miss:
            msg += "\n※ 확인 안 됨: " + ", ".join(miss[:10])
        if saved and saved != self.path:
            msg += f"\n\n엑셀이 열려 있어 완료 표시는 다른 파일에 저장했습니다:\n{saved}"
        messagebox.showinfo(APP_TITLE, msg)
        for p in PARTS:
            self.data[p]["rows"], self.data[p]["results"] = [], []
        if messagebox.askyesno(APP_TITLE, "기존금형 목록을 새로 조회할까요?"):
            self.show_tab("old")
            self.on_fetch_old()

    def on_new_detail(self, p):
        sel = self.tv_new[p].selection()
        if not sel:
            return
        r = next((x for x in self.data[p]["results"] if str(x["rno"]) == sel[0]), None)
        if r:
            messagebox.showinfo(f"{PART_NAME[p]} 시트 {r['rno']}행", r["msg"].replace(", ", "\n"))

    # ==========================================================
    # 기존금형 탭
    # ==========================================================
    def build_old(self):
        t = self.tabs["old"]
        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=(6, 4))
        ttk.Button(bar, text="MES에서 조회", style="Big.TButton", command=self.on_fetch_old).pack(side="left", padx=3)
        ttk.Label(bar, text="  검색").pack(side="left")
        self.var_q = tk.StringVar()
        ttk.Entry(bar, textvariable=self.var_q, width=28).pack(side="left", padx=4)
        self.var_q.trace_add("write", lambda *_: self.show_old())
        self.btn_copy_old = ttk.Button(bar, text="선택행으로 등록양식 만들기", style="Big.TButton", command=self.on_copy_old)
        self.btn_copy_old.pack(side="right", padx=3)

        bar2 = ttk.Frame(t)
        bar2.pack(fill="x", pady=(0, 2))
        self.var_edit_mode = tk.BooleanVar(value=False)

        def _mode():
            if self.var_edit_mode.get():
                if not messagebox.askyesno(APP_TITLE, "수정 모드를 켜면 기존 금형·품번 내용을 바꿀 수 있습니다.\n"
                                                      "(신규등록과는 별개 기능입니다)\n\n켤까요?"):
                    self.var_edit_mode.set(False)
                    return
            else:
                self.end_edit(commit=False)
                if self.var_sa.get():
                    self.var_sa.set(True)
                    self.var_edit_mode.set(True)
                    messagebox.showinfo(APP_TITLE, "다른 이름으로 저장 모드를 먼저 끄세요.")
        ttk.Checkbutton(bar2, text="수정 모드", variable=self.var_edit_mode, command=_mode).pack(side="left", padx=(4, 10))
        self.var_sa = tk.BooleanVar(value=False)
        self.sa_code = None
        ttk.Checkbutton(bar2, text="다른 이름으로 저장 모드", variable=self.var_sa,
                        command=self.on_sa_mode).pack(side="left", padx=(0, 10))
        self.bar2_old = bar2
        self.lbl_sa = tk.Label(t, text="", fg="#FFFFFF", bg="#C55A11", font=("맑은 고딕", 9, "bold"), anchor="w")
        ttk.Button(bar2, text="수정 취소", command=self.on_cancel_edits).pack(side="right", padx=3)
        ttk.Button(bar2, text="선택 금형 삭제", command=self.on_delete_mold).pack(side="right", padx=3)
        ttk.Button(bar2, text="선택 금형 전체컬럼 수정", style="Big.TButton",
                   command=lambda: self.open_detail("mold")).pack(side="right", padx=3)
        ttk.Button(bar2, text="다른 이름으로 저장", style="Big.TButton", command=self.on_save_as_mold).pack(side="right", padx=3)
        self.btn_save_edit = ttk.Button(bar2, text="수정 저장 (0)", style="Big.TButton", command=self.on_save_edits)
        self.btn_save_edit.pack(side="right", padx=3)
        ttk.Label(t, text="맨 앞 ☐ 클릭 = 체크 (머리글 ☐ = 전체).  삭제는 체크한 줄(없으면 선택한 줄)이 대상.  "
                          "수정 모드에서 칸 더블클릭 = 수정 (Enter 확정 / Esc 취소).  주황 = 저장 대기.  "
                          "[다른 이름으로 저장 모드] = 금형 1개 골라 금형코드·금형명·품번까지 고친 뒤 새로 저장 (원본 그대로)",
                  foreground="#555").pack(fill="x", padx=8)

        pw = ttk.PanedWindow(t, orient="vertical")
        pw.pack(fill="both", expand=True, pady=4)
        f1 = ttk.Frame(pw)
        fr, self.tv_old = self.make_tree(f1, ["(조회 전)"])
        fr.pack(fill="both", expand=True)
        self.tv_old.bind("<<TreeviewSelect>>", self.on_old_select)
        self.tv_old.bind("<Double-1>", lambda e: self.begin_edit("mold", e))
        self.chk_old, self.chk_items = set(), set()
        self.setup_check(self.tv_old, "old")
        pw.add(f1, weight=3)
        f2 = ttk.Frame(pw)
        ibar = ttk.Frame(f2)
        ibar.pack(fill="x", padx=2, pady=2)
        self.lbl_old_items = ttk.Label(ibar, text="선택한 금형의 품번", font=("", 10, "bold"))
        self.lbl_old_items.pack(side="left")
        ttk.Button(ibar, text="선택 품번줄 삭제", command=self.on_delete_items).pack(side="left", padx=(16, 3))
        ttk.Button(ibar, text="선택 품번 전체컬럼 수정", command=lambda: self.open_detail("item")).pack(side="left", padx=3)
        ttk.Button(ibar, text="품번 테이블에서 찾기", command=self.on_find_items).pack(side="right", padx=3)
        self.var_iq = tk.StringVar()
        ent = ttk.Entry(ibar, textvariable=self.var_iq, width=20)
        ent.pack(side="right", padx=3)
        ent.bind("<Return>", lambda _e: self.on_find_items())
        ttk.Label(ibar, text="금형코드/품번 일부").pack(side="right")
        fr, self.tv_old_items = self.make_tree(f2, ["-"])
        fr.pack(fill="both", expand=True)
        self.tv_old_items.bind("<Double-1>", lambda e: self.begin_edit("item", e))
        self.setup_check(self.tv_old_items, "items")
        pw.add(f2, weight=1)
        self.lbl_old = ttk.Label(t, text="")
        self.lbl_old.pack(fill="x", padx=4, pady=(0, 4))

    def on_fetch_old(self):
        if self.edits and not messagebox.askyesno(APP_TITLE, f"저장 안 한 수정 {len(self.edits)}건이 있습니다.\n버리고 다시 조회할까요?"):
            return
        self.edits.clear()
        self.update_edit_btn()
        try:
            self.old_cols, self.old_rows, self.old_defs = self.run_db(self.mes.fetch_old)
            self.old_time = dt.datetime.now()
            self.chk_old.clear()
            self.item_cols, self.item_rows, self.item_code = [], [], None
            self.show_old()
        except Exception as e:
            self.err(e)

    def show_old(self):
        if not self.old_time:
            return
        names = [self.head_name(self.old_defs, c) for c in self.old_cols]
        disp = [self.row_display("mold", self.old_cols, row) for row in self.old_rows]
        self.reset_tree(self.tv_old, names, self.fit_widths(names, [v for v, _ in disp]))
        q = self.var_q.get().strip().lower()
        n = 0
        ref = (self.cfg.get("ref_mold") or "").upper()
        ci = self.old_cols.index(self.mes.code_col()) if self.mes.code_col() in self.old_cols else None
        for i, (vals, edited) in enumerate(disp):
            if q and not any(q in v.lower() for v in vals):
                continue
            tag = ("edit",) if edited else (("ok",) if (ci is not None and ref and vals[ci].upper() == ref) else ())
            self.tv_old.insert("", "end", iid=str(i), values=vals, tags=tag,
                               text="☑" if i in self.chk_old else "☐")
            n += 1
        self.lbl_old.config(text=f"조회 {self.old_time:%Y-%m-%d %H:%M}   전체 {len(self.old_rows)}건"
                                 f"{f'   검색결과 {n}건' if q else ''}"
                                 f"   (초록 줄 = 기준 금형 / 여러 줄 선택: Ctrl·Shift+클릭)")

    # ---------- 체크박스 ----------
    def setup_check(self, tv, which):
        tv.configure(show="tree headings")
        tv.column("#0", width=40, minwidth=40, stretch=False, anchor="center")
        tv.heading("#0", text="☐", command=lambda: self.check_all(tv, which))
        tv.bind("<Button-1>", lambda e: self.on_check_click(e, tv, which), add="+")

    def chk_set(self, which):
        return self.chk_old if which == "old" else self.chk_items

    def on_check_click(self, e, tv, which):
        if tv.identify_region(e.x, e.y) != "tree" and tv.identify_column(e.x) != "#0":
            return
        iid = tv.identify_row(e.y)
        if not iid:
            return
        st = self.chk_set(which)
        i = int(iid)
        st.symmetric_difference_update({i})
        tv.item(iid, text="☑" if i in st else "☐")
        self.update_check_label(tv, which)
        return "break" if which == "items" else None

    def check_all(self, tv, which):
        st = self.chk_set(which)
        ids = [int(i) for i in tv.get_children()]
        on = not all(i in st for i in ids)
        for i in ids:
            (st.add if on else st.discard)(i)
            tv.item(str(i), text="☑" if on else "☐")
        self.update_check_label(tv, which)

    def update_check_label(self, tv, which):
        n = len(self.chk_set(which))
        tv.heading("#0", text=f"☑{n}" if n else "☐")

    def picked(self, tv, which):
        """체크한 줄 (없으면 선택한 줄)"""
        st = self.chk_set(which)
        return sorted(st) if st else sorted(int(i) for i in tv.selection())

    def selected_codes(self):
        cc = self.mes.code_col()
        if cc not in self.old_cols:
            return []
        ci = self.old_cols.index(cc)
        return [self.old_rows[int(i)][ci] for i in self.tv_old.selection()]

    def on_old_select(self, _e=None):
        codes = self.selected_codes()
        if len(codes) != 1 or not self.mes.enabled("item"):
            self.lbl_old_items.config(text="선택한 금형의 품번 (금형 한 개를 선택하세요)")
            self.reset_tree(self.tv_old_items, ["-"])
            return
        code = codes[0]
        try:
            cols, rows, defs = self.run_db(lambda c: self.mes.fetch_items_of(c, code))
        except Exception as e:
            self.lbl_old_items.config(text=f"품번 조회 실패: {e}")
            return
        self.item_cols, self.item_rows, self.item_code, self.item_defs = cols, rows, code, defs
        self.item_search = None
        self.chk_items.clear()
        self.show_items()

    def show_items(self):
        cols, rows, code = self.item_cols, self.item_rows, self.item_code
        names = [self.head_name(self.item_defs, c) for c in cols]
        disp = [self.row_display("item", cols, r) for r in rows]
        self.reset_tree(self.tv_old_items, names, self.fit_widths(names, [v for v, _ in disp]))
        for i, (vals, edited) in enumerate(disp):
            self.tv_old_items.insert("", "end", iid=str(i), values=vals, tags=("edit",) if edited else (),
                                     text="☑" if i in self.chk_items else "☐")
        if getattr(self, "item_search", None):
            self.lbl_old_items.config(text=f"품번 테이블 검색 '{self.item_search}' : {len(rows)}건", foreground="#1F4E79")
            return
        self.lbl_old_items.config(text=f"금형 {code} 의 품번 {len(rows)}건"
                                       + ("   ← 품번이 없으면 작업지시가 안 될 수 있음" if not rows else ""),
                                  foreground="#C00000" if not rows else "#1F4E79")
        if not rows and code is not None:
            try:
                cols2, rows2, _ = self.run_db(lambda c: self.mes.find_items(c, show(code).strip()))
                if rows2:
                    li = cols2.index(self.mes.link_col())
                    found = sorted({repr(r[li]) for r in rows2})[:3]
                    self.lbl_old_items.config(text=f"금형 {code} 의 품번 0건  ※ 비슷한 금형코드로 {len(rows2)}건 있음: "
                                                   f"{', '.join(found)} (공백/대소문자 차이?)", foreground="#C00000")
            except Exception:
                pass

    def on_find_items(self):
        q = self.var_iq.get().strip()
        if not q:
            messagebox.showinfo(APP_TITLE, "찾을 금형코드나 품번 일부를 입력하세요.")
            return
        try:
            cols, rows, defs = self.run_db(lambda c: self.mes.find_items(c, q))
        except Exception as e:
            self.err(e)
            return
        self.item_cols, self.item_rows, self.item_defs, self.item_code = cols, rows, defs, None
        self.item_search = q
        self.chk_items.clear()
        self.show_items()

    # ---------- 잘못 등록한 줄 삭제 ----------
    def can_delete(self):
        if self.var_sa.get():
            messagebox.showinfo(APP_TITLE, "다른 이름으로 저장 모드에서는 삭제할 수 없습니다. 모드를 끄고 하세요.")
            return False
        if not self.var_edit_mode.get():
            messagebox.showinfo(APP_TITLE, "삭제는 [수정 모드]를 체크해야 할 수 있습니다.")
            return False
        self.end_edit(commit=True)
        if self.edits:
            messagebox.showinfo(APP_TITLE, "저장 안 한 수정이 있습니다. 먼저 [수정 저장] 또는 [수정 취소]를 하세요.")
            return False
        return True

    def backup(self, targets, now):
        with open(BACKUP_PATH, "a", encoding="utf-8") as f:
            for t in targets:
                f.write(json.dumps({"time": f"{now:%Y-%m-%d %H:%M:%S}", "user": self.cfg.get("reg_user", ""),
                                    "table": self.mes.table(t["p"]),
                                    "row": {k: show(v) for k, v in t["row"].items()}},
                                   ensure_ascii=False) + "\n")

    def after_delete(self, done, now):
        rows = [[f"{now:%Y-%m-%d %H:%M:%S}", self.cfg.get("reg_user", ""), f"{PART_NAME[t['p']]}삭제",
                 self.mes.table(t["p"]), "/".join(show(v).strip() for v in t["key"].values()),
                 "; ".join(f"{k}={cell_text(v)}" for k, v in t["row"].items() if v is not None)] for t in done]
        self.append_log(rows)
        self.load_log()

    def item_targets(self, idxs):
        out = []
        for i in idxs:
            row = self.item_rows[i]
            key = self.row_key("item", self.item_cols, row)
            if key is None:
                raise ValueError("품번 매핑에 중복검사키가 없어 어느 줄인지 특정할 수 없습니다.")
            out.append({"p": "item", "key": dict(key), "row": dict(zip(self.item_cols, row)),
                        "label": "품번 " + "/".join(show(v).strip() for _, v in key)})
        return out

    def rid_rows(self, conn, p, col, value):
        """해당 테이블에서 col = value 인 줄들을 ROWID와 함께 읽기"""
        cur = conn.cursor()
        cur.execute(f"SELECT ROWIDTOCHAR(T.ROWID) RID__, T.* FROM {self.mes.table(p)} T "
                    f"WHERE {ident(col, '컬럼')} = :v", v=value)
        names = [d[0] for d in cur.description]
        return [dict(zip(names, r)) for r in cur.fetchall()]

    def delete_plan(self, conn, plan):
        """plan: [(p, rid, label)] -> 한 묶음으로 삭제 (1줄씩 확인, 하나라도 이상하면 전부 취소)"""
        cur = conn.cursor()
        try:
            for p, rid, label in plan:
                cur.execute(f"DELETE FROM {self.mes.table(p)} WHERE ROWID = CHARTOROWID(:r)", r=rid)
                if cur.rowcount != 1:
                    raise RuntimeError(f"{label}: 이미 없어졌거나 바뀐 줄입니다. 다시 조회하세요.")
            conn.commit()
        except Exception as ex:
            conn.rollback()
            raise RuntimeError(f"삭제 중 오류 -> 전체 취소 (MES는 그대로)\n\n{ora_hint(ex)}")
        return len(plan)

    def log_deleted(self, rows, now, kind):
        """rows: [(p, row dict)] -> 백업 파일 + 작업기록"""
        with open(BACKUP_PATH, "a", encoding="utf-8") as f:
            for p, row in rows:
                f.write(json.dumps({"time": f"{now:%Y-%m-%d %H:%M:%S}", "user": self.cfg.get("reg_user", ""),
                                    "table": self.mes.table(p),
                                    "row": {k: show(v) for k, v in row.items() if k != "RID__"}},
                                   ensure_ascii=False) + "\n")
        self.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", self.cfg.get("reg_user", ""), kind, self.mes.table(p),
                          f"{show(row.get(self.mes.code_col())).strip()}/{show(row.get('ITEM_CODE')).strip()}",
                          "; ".join(f"{k}={cell_text(v)}" for k, v in row.items() if v is not None and k != "RID__")]
                         for p, row in rows])
        self.load_log()

    def on_delete_items(self):
        if not self.can_delete():
            return
        idxs = self.picked(self.tv_old_items, "items")
        if not idxs:
            messagebox.showinfo(APP_TITLE, "아래 품번 표에서 지울 줄을 체크(☑)하거나 선택하세요.")
            return
        lc = self.mes.link_col()
        if lc not in self.item_cols:
            messagebox.showerror(APP_TITLE, f"품번 표에 연결 컬럼({lc})이 없습니다.")
            return
        li = self.item_cols.index(lc)
        picked = [self.item_rows[i] for i in idxs]

        def find(conn):
            """화면의 줄과 모든 칸이 같은 DB 줄을 ROWID로 찾기"""
            out, used = [], set()
            cache = {}
            for row in picked:
                v = row[li]
                if v not in cache:
                    cache[v] = self.rid_rows(conn, "item", lc, v)
                want = [cell_text(x) for x in row]
                hit = next((r for r in cache[v] if r["RID__"] not in used
                            and [cell_text(r.get(c)) for c in self.item_cols] == want), None)
                if not hit:
                    raise RuntimeError(f"품번 줄을 MES에서 찾지 못했습니다 (그 사이 바뀜?). 다시 조회하세요.\n"
                                       f"{show(v)} / {show(row[self.item_cols.index('ITEM_CODE')]) if 'ITEM_CODE' in self.item_cols else ''}")
                used.add(hit["RID__"])
                out.append(hit)
            return out
        try:
            hits = self.run_db(find)
        except Exception as e:
            self.err(e)
            return
        label = lambda r: f"금형 {show(r.get(lc)).strip()} - 품번 {show(r.get('ITEM_CODE')).strip()}"  # noqa: E731
        lines = "\n".join(f"  {label(r)}" for r in hits[:20]) + (f"\n  ... 외 {len(hits) - 20}줄" if len(hits) > 20 else "")
        if not messagebox.askyesno("품번 연결 삭제", f"아래 {len(hits)}줄(금형-품번 연결)을 MES에서 삭제합니다.\n\n{lines}\n\n"
                                                f"금형과 품번 자체는 지워지지 않고 연결만 끊깁니다.\n"
                                                f"(삭제한 내용은 백업 파일에 저장됩니다)\n계속할까요?", icon="warning"):
            return
        now = dt.datetime.now()
        try:
            n = self.run_db(lambda c: self.delete_plan(c, [("item", r["RID__"], label(r)) for r in hits]))
        except Exception as e:
            self.err(e)
            return
        self.log_deleted([("item", r) for r in hits], now, "품번연결삭제")
        messagebox.showinfo(APP_TITLE, f"{n}줄 삭제 완료.\n백업: {BACKUP_PATH}")
        self.chk_items.clear()
        self.update_check_label(self.tv_old_items, "items")
        if getattr(self, "item_search", None):
            self.on_find_items()
        else:
            self.on_old_select()

    def on_delete_mold(self):
        if not self.can_delete():
            return
        cc = self.mes.code_col()
        if cc not in self.old_cols:
            messagebox.showerror(APP_TITLE, f"금형 표에 금형코드 컬럼({cc})이 없습니다.")
            return
        ci = self.old_cols.index(cc)
        codes = []
        for i in self.picked(self.tv_old, "old"):
            c = self.old_rows[i][ci]
            if c not in codes:
                codes.append(c)
        if not codes:
            messagebox.showinfo(APP_TITLE, "위 금형 표에서 지울 금형을 체크(☑)하거나 선택하세요.")
            return
        ref = (self.cfg.get("ref_mold") or "").strip().upper()
        if any(show(c).strip().upper() == ref for c in codes):
            messagebox.showwarning(APP_TITLE, f"기준 금형({ref})은 삭제할 수 없습니다. 체크를 풀거나 다른 금형을 기준으로 지정하세요.")
            return
        lc = self.mes.link_col()

        def work(conn, progress, cancelled):
            h, out = {"c": conn}, []
            try:
                for c in codes:
                    used, failed = self.mes.usage_of_code(h, c, progress, cancelled)
                    if cancelled():
                        break
                    out.append({"code": c, "used": used, "failed": failed,
                                "mold": self.rid_rows(h["c"], "mold", cc, c),
                                "items": self.rid_rows(h["c"], "item", lc, c) if self.mes.enabled("item") else []})
            finally:
                if h["c"] is not conn:
                    try:
                        h["c"].close()
                    except Exception:
                        pass
            return out
        try:
            info = self.run_bg(work, "금형 사용 이력 확인 중")
        except Exception as e:
            self.err(e)
            return
        blocked = [x for x in info if x["used"]]
        ok = [x for x in info if not x["used"] and x["mold"]]
        if blocked:
            msg = "아래 금형은 작업지시·실적 등 이력에서 쓰이고 있어 삭제하지 않습니다:\n\n"
            for x in blocked[:10]:
                msg += f"■ {show(x['code']).strip()}\n" + "\n".join(
                    f"     {t}.{c}: 기록 있음" for t, c, n in x["used"][:6]) + "\n"
            msg += ("\n※ 이력이 있는 금형을 지우면 작업지시·실적 기록이 깨집니다.\n"
                    "   잘못 연결된 품번만 끊으려면 아래 품번 표에서 [선택 품번줄 삭제]를 쓰세요.")
            if not ok:
                messagebox.showwarning(APP_TITLE, msg)
                return
            if not messagebox.askyesno(APP_TITLE, msg + f"\n\n나머지 {len(ok)}개만 삭제를 계속할까요?", icon="warning"):
                return
        if not ok:
            messagebox.showinfo(APP_TITLE, "삭제할 금형을 MES에서 찾지 못했습니다. 다시 조회하세요.")
            return
        lines = "\n".join(f"  {show(x['code']).strip()}  (품번 연결 {len(x['items'])}줄)" for x in ok[:20])
        fails = sorted({t for x in ok for t in x["failed"]})
        warn = (f"\n※ 확인하지 못한 테이블 {len(fails)}개 (너무 크거나 오류): {', '.join(fails[:5])}\n"
                f"   이 표에 이 금형 기록이 있을 수도 있으니 확실할 때만 진행하세요." if fails else "")
        one = show(ok[0]["code"]).strip() if len(ok) == 1 else None
        typed = simpledialog.askstring(
            "금형 삭제", f"아래 금형 {len(ok)}개를 삭제합니다 (금형 + 연결된 품번 줄).\n\n{lines}\n\n"
                        f"작업지시·샷수 등 다른 이력에서는 쓰이지 않는 것을 확인했습니다.{warn}\n"
                        f"삭제한 내용은 백업 파일에 저장됩니다.\n\n"
                        + (f"확인을 위해 금형코드 {one} 를 그대로 입력하세요:" if one else "확인을 위해 '삭제'라고 입력하세요:"),
            parent=self.root)
        if typed is None:
            return
        if (typed.strip().upper() != one.upper()) if one else (typed.strip() != "삭제"):
            messagebox.showinfo(APP_TITLE, "입력한 글자가 달라서 삭제하지 않았습니다.")
            return
        plan, rows = [], []
        for x in ok:
            cs = show(x["code"]).strip()
            for r in x["items"]:
                plan.append(("item", r["RID__"], f"금형 {cs} - 품번 {show(r.get('ITEM_CODE')).strip()}"))
                rows.append(("item", r))
            for r in x["mold"]:
                plan.append(("mold", r["RID__"], f"금형 {cs}"))
                rows.append(("mold", r))
        now = dt.datetime.now()
        try:
            self.run_db(lambda c: self.delete_plan(c, plan))
        except Exception as e:
            self.err(e)
            return
        self.log_deleted(rows, now, "금형삭제")
        ni = sum(len(x["items"]) for x in ok)
        messagebox.showinfo(APP_TITLE, f"금형 {len(ok)}개 삭제 완료 (품번 연결 {ni}줄 포함).\n백업: {BACKUP_PATH}")
        self.edits.clear()
        self.old_time = None
        self.on_fetch_old()

    # ---------- 기존 내용 바로 수정 ----------
    def row_key(self, p, cols, row):
        keys = [c["col"] for c in self.mes.keys(p)]
        if not keys or any(k not in cols for k in keys):
            if p == "mold":
                keys = [self.mes.code_col()]
            else:
                keys = [k for k in (self.mes.link_col(), "ITEM_CODE", "ORGANIZATION_ID") if k in cols]
            if not keys or any(k not in cols for k in keys):
                return None
        return tuple((k, row[cols.index(k)]) for k in keys)

    def row_display(self, p, cols, row):
        vals = [cell_text(v) for v in row]
        k = self.row_key(p, cols, row)
        e = self.edits.get((p, k)) if k else None
        if not e:
            return vals, False
        for col, (old, new) in e["changes"].items():
            if col in cols:
                vals[cols.index(col)] = cell_text(new)
        return vals, True

    def update_edit_btn(self):
        n = sum(len(e["changes"]) for e in self.edits.values())
        self.btn_save_edit.config(text=f"수정 저장 ({n})")

    @staticmethod
    def head_name(defs, col):
        c = defs.get(col)
        return col if (not c or c["name"] == col) else f"{c['name']} ({col})"

    @staticmethod
    def fit_widths(names, rows_vals):
        out = []
        for i, n in enumerate(names):
            w = blen(n, 2)
            for vals in rows_vals[:300]:
                if i < len(vals):
                    w = max(w, blen(vals[i][:40], 2))
            out.append(max(60, min(340, w * 8 + 20)))
        return out

    def src(self, p):
        if p == "mold":
            return self.tv_old, self.old_cols, self.old_rows, self.old_defs
        return self.tv_old_items, self.item_cols, self.item_rows, self.item_defs

    def edit_block(self, p, col, row, cols, defs):
        """수정 못 하는 이유 (되면 None)"""
        if not self.var_edit_mode.get():
            return "지금은 보기 전용입니다.\n고치려면 위쪽 [수정 모드]를 체크하세요."
        c = defs.get(col)
        if c is None:
            return "컬럼 정보를 찾을 수 없습니다."
        if self.row_key(p, cols, row) is None:
            return (f"{PART_NAME[p]} 매핑에 중복검사키가 없어서 어느 줄인지 특정할 수 없습니다.\n"
                    f"컬럼매핑에서 중복검사키를 지정하세요.")
        if self.var_sa.get():
            lk = self.mes.code_col() if p == "mold" else self.mes.link_col()
            if lk in cols and show(row[cols.index(lk)]).strip() != self.sa_code:
                return f"다른 이름으로 저장 모드에서는 고른 금형 {self.sa_code} 과(와) 그 품번만 고칠 수 있습니다."
            if p == "item" and col == self.mes.link_col():
                return "품번의 금형코드 칸은 저장할 때 새 금형코드로 자동으로 바뀝니다. 위 금형 표의 금형코드를 고치세요."
        elif (p == "mold" and col == self.mes.code_col()) or (p == "item" and col == self.mes.link_col()):
            return ("금형코드는 작업지시·샷수 이력 등 여러 곳에 연결되어 있어 여기서 수정할 수 없습니다.\n"
                    "새 코드가 필요하면 [다른 이름으로 저장 모드]를 쓰세요.")
        if not c.get("editable", True):
            return f"{col} ({c.get('dbtype', '')}) 형식은 여기서 수정할 수 없습니다."
        return None

    def confirm_key(self, c):
        return (not c["key"]) or messagebox.askyesno(
            APP_TITLE, f"'{c['name']}'은(는) 줄을 구분하는 키 값입니다.\n"
                       f"다른 곳에서 쓰고 있지 않은지 확인하셨나요?\n\n수정할까요?")

    def set_edit(self, p, cols, row, col, text, defs):
        """수정 대기 목록에 반영. 오류면 메시지 반환"""
        c = defs[col]
        old = row[cols.index(col)]
        try:
            new = conv(c, text)
            if new is None and c["req"]:
                raise ValueError("필수 칸이라 비울 수 없음")
            if c["type"] == "TEXT" and c["maxlen"] and new is not None:
                n = blen(new, int(self.cfg.get("kor_bytes") or 3))
                if n > int(c["maxlen"]):
                    raise ValueError(f"너무 김 ({n}/{c['maxlen']}byte)")
        except ValueError as e:
            return f"{c['name']}: {e}\n수정하지 않았습니다."
        key = self.row_key(p, cols, row)
        k = (p, key)
        same = (cell_text(new) == cell_text(old) and (new is None) == (old is None)) or \
               (isinstance(old, str) and isinstance(new, str) and old.rstrip() == new)
        if same:
            if k in self.edits:
                self.edits[k]["changes"].pop(col, None)
                if not self.edits[k]["changes"]:
                    del self.edits[k]
        else:
            label = f"{PART_NAME[p]} " + "/".join(show(v) for _, v in key)
            e = self.edits.setdefault(k, {"p": p, "key": dict(key), "changes": {}, "label": label})
            e["changes"][col] = (old, new)
        self.update_edit_btn()
        return None

    def refresh_row(self, p, idx):
        tv, cols, rows, _ = self.src(p)
        if tv.exists(str(idx)):
            vals, edited = self.row_display(p, cols, rows[idx])
            tv.item(str(idx), values=vals, tags=("edit",) if edited else ())

    def begin_edit(self, p, event):
        self.end_edit(commit=True)
        tv, cols, rows, defs = self.src(p)
        if tv.identify_region(event.x, event.y) != "cell":
            return
        if not self.var_edit_mode.get():          # 보기 전용: 세로 상세 보기만
            if tv.identify_row(event.y):
                tv.selection_set(tv.identify_row(event.y))
                self.open_detail(p)
            return
        iid, colid = tv.identify_row(event.y), tv.identify_column(event.x)
        if not iid or not colid:
            return
        ci = int(colid[1:]) - 1
        if ci < 0 or ci >= len(cols):
            return
        col, row = cols[ci], rows[int(iid)]
        why = self.edit_block(p, col, row, cols, defs)
        if why:
            messagebox.showwarning(APP_TITLE, why)
            return
        if not self.confirm_key(defs[col]):
            return
        bbox = tv.bbox(iid, colid)
        if not bbox:
            return
        x, y, w, h = bbox
        ent = ttk.Entry(tv)
        ent.place(x=x, y=y, width=max(w, 160), height=h)
        ent.insert(0, tv.item(iid, "values")[ci])
        ent.select_range(0, "end")
        ent.focus_set()
        self._editor = {"ent": ent, "p": p, "idx": int(iid), "col": col}
        ent.bind("<Return>", lambda _e: self.end_edit(commit=True))
        ent.bind("<KP_Enter>", lambda _e: self.end_edit(commit=True))
        ent.bind("<Escape>", lambda _e: self.end_edit(commit=False))
        ent.bind("<FocusOut>", lambda _e: self.end_edit(commit=True))
        tv.bind("<MouseWheel>", lambda _e: self.end_edit(commit=True), add="+")

    def end_edit(self, commit):
        ed = self._editor
        if not ed:
            return
        self._editor = None
        text = ed["ent"].get()
        ed["ent"].destroy()
        if not commit:
            return
        p, idx = ed["p"], ed["idx"]
        _, cols, rows, defs = self.src(p)
        msg = self.set_edit(p, cols, rows[idx], ed["col"], text, defs)
        if msg:
            messagebox.showerror(APP_TITLE, msg)
        self.refresh_row(p, idx)

    # ---------- 한 줄 전체 컬럼 세로 보기/수정 ----------
    def open_detail(self, p):
        self.end_edit(commit=True)
        tv, cols, rows, defs = self.src(p)
        sel = tv.selection()
        if len(sel) != 1:
            messagebox.showinfo(APP_TITLE, f"{PART_NAME[p]} 표에서 한 줄을 선택하세요.")
            return
        idx = int(sel[0])
        row = rows[idx]
        key = self.row_key(p, cols, row)
        title = f"{PART_NAME[p]} 상세 수정 - " + ("/".join(show(v) for _, v in key) if key else f"{idx + 1}번째 줄")
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("760x720")
        win.transient(self.root)
        ttk.Label(win, text=("값 칸을 더블클릭해서 수정 → 주황색 = 저장 대기.  [수정 저장]을 눌러야 MES에 반영됩니다."
                            if self.var_edit_mode.get() else "보기 전용 (고치려면 [금형·품번 관리] 탭의 [수정 모드] 체크)"),
                  foreground="#555").pack(fill="x", padx=8, pady=(8, 2))
        qv = tk.StringVar()
        top = ttk.Frame(win)
        top.pack(fill="x", padx=8)
        ttk.Label(top, text="컬럼 찾기").pack(side="left")
        ttk.Entry(top, textvariable=qv, width=24).pack(side="left", padx=4)
        fr, dtv = self.make_tree(win, ["DB컬럼", "이름", "값", "형식"], [190, 150, 300, 90])
        fr.pack(fill="both", expand=True, padx=8, pady=6)

        def refill(*_):
            dtv.delete(*dtv.get_children())
            vals, _ = self.row_display(p, cols, row)
            e = self.edits.get((p, key)) if key else None
            q = qv.get().strip().lower()
            for j, col in enumerate(cols):
                c = defs.get(col, {})
                nm = c.get("name", col)
                if q and q not in col.lower() and q not in nm.lower() and q not in vals[j].lower():
                    continue
                tag = ()
                if e and col in e["changes"]:
                    tag = ("edit",)
                elif self.edit_block(p, col, row, cols, defs):
                    tag = ("off",)
                typ = c.get("dbtype", "") + (" 필수" if c.get("req") else "")
                dtv.insert("", "end", iid=str(j), values=[col, "" if nm == col else nm, vals[j], typ], tags=tag)

        def edit(_e=None):
            s_ = dtv.selection()
            if not s_:
                return
            j = int(s_[0])
            col = cols[j]
            why = self.edit_block(p, col, row, cols, defs)
            if why:
                messagebox.showwarning(APP_TITLE, why, parent=win)
                return
            if not self.confirm_key(defs[col]):
                return
            cur_txt = dtv.item(s_[0], "values")[2]
            v = simpledialog.askstring(col, f"{defs[col]['name']} ({col})\n새 값을 입력하세요:",
                                       initialvalue=cur_txt, parent=win)
            if v is None:
                return
            msg = self.set_edit(p, cols, row, col, v, defs)
            if msg:
                messagebox.showerror(APP_TITLE, msg, parent=win)
            refill()
            self.refresh_row(p, idx)
            dtv.selection_set(str(j))
            dtv.see(str(j))

        def save():
            if self.on_save_edits():
                win.destroy()

        dtv.bind("<Double-1>", edit)
        dtv.bind("<Return>", edit)
        qv.trace_add("write", refill)
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(bar, text="수정 저장", style="Big.TButton", command=save).pack(side="left", padx=3)
        ttk.Button(bar, text="닫기", command=win.destroy).pack(side="right", padx=3)
        refill()

    def on_cancel_edits(self):
        self.end_edit(commit=False)
        if not self.edits:
            return
        if not messagebox.askyesno(APP_TITLE, f"저장 안 한 수정 {len(self.edits)}줄을 모두 되돌릴까요?"):
            return
        self.edits.clear()
        self.update_edit_btn()
        self.show_old()
        if self.item_code is not None:
            self.show_items()

    # ---------- 다른 이름으로 저장 모드 ----------
    def on_sa_mode(self):
        """켜면: 금형 1개만 골라서 그 금형·품번 칸(금형코드·금형명·품번 포함)을 표에서 고친 뒤 [다른 이름으로 저장]"""
        self.end_edit(commit=True)
        if self.var_sa.get():
            idxs = self.picked(self.tv_old, "old")
            cc = self.mes.code_col()
            if len(idxs) != 1 or cc not in self.old_cols:
                self.var_sa.set(False)
                messagebox.showinfo(APP_TITLE, "다른 이름으로 저장 모드는 금형 1개를 먼저 골라야 합니다.\n"
                                               "[MES에서 조회] → 금형 1개 체크(☑) 또는 클릭 → 다시 체크하세요.")
                return
            if self.edits:
                if not messagebox.askyesno(APP_TITLE, f"저장 안 한 수정 {len(self.edits)}건이 있습니다.\n버리고 시작할까요?"):
                    self.var_sa.set(False)
                    return
                self.edits.clear()
                self.update_edit_btn()
            i = idxs[0]
            self.sa_code = show(self.old_rows[i][self.old_cols.index(cc)]).strip()
            self.var_edit_mode.set(True)
            self.tv_old.selection_set(str(i))
            self.tv_old.see(str(i))
            self.on_old_select()
            self.lbl_sa.config(text=f"  ▶ 다른 이름으로 저장 모드: {self.sa_code} 복사본 만드는 중 - 위·아래 칸(금형코드·금형명·품번 포함)을 "
                                    f"더블클릭해 고친 뒤 [다른 이름으로 저장]  (원본은 안 바뀜)")
            self.lbl_sa.pack(fill="x", padx=4, pady=(0, 2), after=self.bar2_old)
            self.show_old()
        else:
            if self.edits and not messagebox.askyesno(APP_TITLE, "고쳐 놓은 내용은 저장되지 않고 사라집니다.\n"
                                                                 "다른 이름으로 저장 모드를 끌까요?"):
                self.var_sa.set(True)
                return
            self.edits.clear()
            self.update_edit_btn()
            self.sa_code = None
            self.lbl_sa.pack_forget()
            self.show_old()
            if self.item_cols:
                self.show_items()

    # ---------- 다른 이름으로 저장 (원본 그대로, 금형+품번을 새로 복사) ----------
    def pending_vals(self, p, cols, row):
        """표에서 고쳐 놓고 저장 안 한 값까지 반영한 줄 (dict)"""
        vals = dict(zip(cols, row))
        k = self.row_key(p, cols, row)
        e = self.edits.get((p, k)) if k else None
        for c, (_o, new) in (e or {}).get("changes", {}).items():
            vals[c] = new
        return vals, (p, k) if e else None

    def on_save_as_mold(self):
        self.end_edit(commit=True)
        cc, lc = self.mes.code_col(), self.mes.link_col()
        if self.var_sa.get() and self.sa_code and cc in self.old_cols:
            ci0 = self.old_cols.index(cc)
            idxs = [i for i, r in enumerate(self.old_rows) if show(r[ci0]).strip() == self.sa_code][:1]
        else:
            idxs = self.picked(self.tv_old, "old")
        if len(idxs) != 1:
            messagebox.showinfo(APP_TITLE, "복사할 금형 1개를 체크(☑)하거나 선택하세요.")
            return
        cols, defs = self.old_cols, self.old_defs
        if cc not in cols:
            messagebox.showerror(APP_TITLE, f"금형 표에 금형코드 컬럼({cc})이 없습니다.")
            return
        mrow = self.old_rows[idxs[0]]
        mvals, mkey = self.pending_vals("mold", cols, mrow)
        old_code = show(mrow[cols.index(cc)]).strip()
        nc = "MOLD_NAME" if "MOLD_NAME" in cols else None

        # 아래 품번 표 (이 금형의 품번 + 저장 안 한 수정값)
        icols, irows = self.item_cols, self.item_rows
        if self.item_code is None or show(self.item_code).strip() != old_code or getattr(self, "item_search", None):
            try:
                icols, irows, _d = self.run_db(lambda c: self.mes.fetch_items_of(c, mrow[cols.index(cc)]))
            except Exception as e:
                self.err(e)
                return
        items, ikeys = [], []
        for r in irows:
            v, k = self.pending_vals("item", icols, r)
            items.append({"on": True, "vals": v, "orig": show(r[icols.index("ITEM_CODE")]).strip()
                          if "ITEM_CODE" in icols else ""})
            if k:
                ikeys.append(k)
        ITEM_EDIT = [c for c in ("ITEM_CODE", "CAVITY_QTY", "ST_VALUE") if c in icols]

        win = tk.Toplevel(self.root)
        win.title("다른 이름으로 저장")
        win.geometry("800x620")
        win.transient(self.root)
        bar = ttk.Frame(win)                         # 버튼은 맨 아래 고정 (창이 작아도 보이게)
        bar.pack(side="bottom", fill="x", padx=8, pady=8)
        bg = "#EEF4FB"
        head = tk.Frame(win, bg=bg, highlightbackground="#B7CCE4", highlightthickness=1)
        head.pack(fill="x", padx=8, pady=8)
        tk.Label(head, bg=bg, fg="#1F4E79", font=("맑은 고딕", 11, "bold"), anchor="w",
                 text=f"금형 {old_code} 을(를) 새 금형으로 복사 (금형 + 품번)").pack(fill="x", padx=10, pady=(6, 0))
        tk.Label(head, bg=bg, fg="#333", anchor="w", justify="left", font=("맑은 고딕", 9),
                 text="원본 금형·품번은 바뀌지 않고 새 줄만 추가됩니다. 나머지 칸은 그대로 복사, 샷수 같은 누적값은 비웁니다.\n"
                      "품번 하나는 금형 하나에만 연결되므로, 아래 품번은 새 품번으로 바꿔야 합니다 (칸 더블클릭).").pack(
            fill="x", padx=10, pady=(0, 6))

        top = ttk.Frame(win, padding=(16, 2))
        top.pack(fill="x")
        v_code = tk.StringVar(value=show(mvals.get(cc)).strip())
        v_name = tk.StringVar(value=show(mvals.get(nc)).strip() if nc else "")
        ttk.Label(top, text="새 금형코드 *", font=("", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        e1 = ttk.Entry(top, textvariable=v_code, width=30)
        e1.grid(row=0, column=1, sticky="w", padx=8)
        if nc:
            ttk.Label(top, text="새 금형명 *", font=("", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
            ttk.Entry(top, textvariable=v_name, width=50).grid(row=1, column=1, sticky="w", padx=8)
        v_dw = tk.BooleanVar(value="DRAWING_NO" in cols and show(mvals.get("DRAWING_NO")).strip() == old_code)
        if "DRAWING_NO" in cols:
            ttk.Checkbutton(top, text="도면번호도 새 금형코드로 바꾸기", variable=v_dw).grid(row=2, column=1, sticky="w", padx=8)

        mid = ttk.LabelFrame(win, text=" 같이 복사할 품번 (☑ = 포함, 칸 더블클릭 = 고치기) ", padding=6)
        mid.pack(fill="both", expand=True, padx=8, pady=6)
        heads = ["포함", "원래 품번"] + [{"ITEM_CODE": "새 품번 *", "CAVITY_QTY": "캐비티", "ST_VALUE": "ST"}[c] for c in ITEM_EDIT] + ["상태"]
        fr, tv = self.make_tree(mid, heads, [50, 170] + [170 if c == "ITEM_CODE" else 70 for c in ITEM_EDIT] + [230])
        tv.configure(height=6)
        fr.pack(fill="both", expand=True)
        st = {}
        ok_missing = set()

        def refill():
            tv.delete(*tv.get_children())
            for i, it in enumerate(items):
                v = it["vals"]
                tag = st.get(i, ("", ""))
                tv.insert("", "end", iid=str(i), tags=(tag[0],) if tag[0] else (),
                          values=["☑" if it["on"] else "☐", it["orig"] or "(새로 추가)"]
                          + [show(v.get(c)) for c in ITEM_EDIT] + [tag[1]])

        def on_dbl(e):
            iid, colid = tv.identify_row(e.y), tv.identify_column(e.x)
            if not iid:
                return
            i, ci = int(iid), int(colid[1:]) - 1
            if ci <= 1:
                items[i]["on"] = not items[i]["on"]
                refill()
                return
            c = ITEM_EDIT[ci - 2]
            t = simpledialog.askstring(c, f"{heads[ci]} 값을 넣으세요:", initialvalue=show(items[i]["vals"].get(c)),
                                       parent=win)
            if t is None:
                return
            t = t.strip()
            if c == "ITEM_CODE":
                items[i]["vals"][c] = t or None
            else:
                try:
                    items[i]["vals"][c] = (int(float(t)) if float(t).is_integer() else float(t)) if t else None
                except ValueError:
                    messagebox.showwarning(APP_TITLE, "숫자를 넣으세요.", parent=win)
                    return
            st.pop(i, None)
            refill()

        def on_click(e):
            iid, colid = tv.identify_row(e.y), tv.identify_column(e.x)
            if iid and colid == "#1":
                items[int(iid)]["on"] = not items[int(iid)]["on"]
                refill()
                return "break"
        tv.bind("<Double-1>", on_dbl)
        tv.bind("<Button-1>", on_click, add="+")

        def add_item():
            base = dict(items[0]["vals"]) if items else {c: None for c in icols}
            base["ITEM_CODE"] = None
            items.append({"on": True, "vals": base, "orig": ""})
            refill()
        ib = ttk.Frame(mid)
        ib.pack(fill="x", pady=(4, 0))
        ttk.Button(ib, text="+ 품번 추가", command=add_item).pack(side="left")
        lbl = ttk.Label(win, text="", foreground="#C62828", wraplength=720)
        lbl.pack(fill="x", padx=12)
        refill()
        e1.focus_set()
        e1.select_range(0, "end")

        def save():
            code, name = v_code.get().strip(), v_name.get().strip()
            if not code or code.upper() == old_code.upper():
                lbl.config(text="새 금형코드를 넣으세요 (원본과 달라야 합니다).")
                return
            if nc and not name:
                lbl.config(text="새 금형명을 넣으세요.")
                return
            use = [(i, it) for i, it in enumerate(items) if it["on"]]
            codes = [show(it["vals"].get("ITEM_CODE")).strip() for _, it in use]
            if any(not c for c in codes):
                lbl.config(text="포함(☑)한 품번 중 비어 있는 품번이 있습니다. 넣거나 체크를 푸세요.")
                return
            if len(set(c.upper() for c in codes)) != len(codes):
                lbl.config(text="같은 품번이 두 번 들어 있습니다.")
                return
            reg = (self.cfg.get("reg_user") or "").strip()
            now = dt.datetime.now()
            new_m = dict(mvals)
            new_m[cc] = code
            if nc:
                new_m[nc] = name
            if v_dw.get() and "DRAWING_NO" in new_m:
                new_m["DRAWING_NO"] = code

            def prep(row):
                for c in list(row):
                    if any(w in c for w in RESET_WORDS):
                        row[c] = None
                    elif RE_AUTO_DATE.search(c) or RE_UPD_DATE.search(c):
                        if isinstance(row[c], dt.datetime) or row[c] is None:
                            row[c] = now
                    elif reg and (RE_AUTO_USER.search(c) or RE_UPD_USER.search(c)):
                        row[c] = reg
                return row
            new_m = prep(new_m)
            new_items = []
            for _, it in use:
                r = prep(dict(it["vals"]))
                r[lc] = code
                new_items.append(r)
            tm, ti = self.mes.table("mold"), self.mes.table("item")

            def work(conn):
                cur = conn.cursor()
                cur.execute(f"SELECT COUNT(*) FROM {tm} WHERE UPPER(TRIM({ident(cc, '컬럼')})) = :c", c=code.upper())
                if cur.fetchone()[0]:
                    raise ValueError(f"금형코드 {code} 는 이미 있습니다. 다른 코드를 넣으세요.")
                bad = {}
                for i, (k, it) in enumerate(use):
                    ic = codes[i]
                    cur.execute("SELECT COUNT(*) FROM ICOM_ITEM_MASTER WHERE ITEM_CODE = :i", i=ic)
                    if not cur.fetchone()[0] and ic not in ok_missing:
                        bad[k] = ("warn", "품목 마스터에 없음 - 한 번 더 누르면 그대로 저장")
                        continue
                    cur.execute(f"SELECT MAX({ident(lc, '컬럼')}) FROM {ti} WHERE ITEM_CODE = :i", i=ic)
                    m = cur.fetchone()[0]
                    if m:
                        bad[k] = ("err", f"이미 금형 {show(m).strip()}에 연결됨 → 새 품번으로")
                if bad:
                    return bad
                mm, mi = table_meta(conn, tm), table_meta(conn, ti)
                try:
                    insert_dict(cur, tm, mm, new_m)
                    for r in new_items:
                        insert_dict(cur, ti, mi, r)
                    conn.commit()
                except Exception:
                    conn.rollback()
                    raise
                return None
            try:
                bad = self.run_db(work)
            except ValueError as e:
                lbl.config(text=str(e))
                return
            except Exception as e:
                messagebox.showerror(APP_TITLE, "등록 중 오류 -> 전부 취소 (MES는 그대로)\n\n" + ora_hint(e), parent=win)
                return
            if bad:
                st.clear()
                st.update(bad)
                refill()
                errs = [k for k, (t, _) in bad.items() if t == "err"]
                warns = [k for k, (t, _) in bad.items() if t == "warn"]
                for k in warns:                      # 품목 마스터에 없는 품번: 확인했으면 다음엔 통과
                    ok_missing.add(show(items[k]["vals"].get("ITEM_CODE")).strip())
                lbl.config(text=(f"빨간 줄 {len(errs)}개를 고치세요. " if errs else "")
                           + (f"노란 줄 {len(warns)}개는 품목 마스터(ERP)에 없는 품번입니다 - 맞으면 [MES에 등록]을 한 번 더 누르세요. "
                              if warns else "") + "(아직 아무것도 등록하지 않았습니다)")
                return
            self.append_log([[f"{now:%Y-%m-%d %H:%M:%S}", reg, "금형 다른이름저장", tm, f"{old_code} -> {code}",
                              f"품번 {len(new_items)}개: {', '.join(codes)}"]])
            self.load_log()
            for k in [mkey] + ikeys:                 # 원본에 걸려 있던 '저장 대기' 수정은 버림 (원본은 그대로)
                if k:
                    self.edits.pop(k, None)
            self.update_edit_btn()
            if self.var_sa.get():
                self.var_sa.set(False)
                self.sa_code = None
                self.lbl_sa.pack_forget()
            win.destroy()
            messagebox.showinfo(APP_TITLE, f"새 금형 {code} 를 등록했습니다 (품번 연결 {len(new_items)}개).\n"
                                           f"원본 {old_code} 는 그대로입니다.")
            self.var_q.set(code)
            self.on_fetch_old()

        ttk.Button(bar, text="MES에 등록", style="Big.TButton", command=save).pack(side="left", padx=3)
        ttk.Button(bar, text="닫기", command=win.destroy).pack(side="right", padx=3)

    def on_save_edits(self):
        self.end_edit(commit=True)
        if self.var_sa.get():
            messagebox.showinfo(APP_TITLE, "다른 이름으로 저장 모드입니다. 원본을 덮어쓰지 않도록 [수정 저장]은 막혀 있습니다.\n"
                                           "[다른 이름으로 저장]을 누르세요.")
            return False
        if not self.edits:
            messagebox.showinfo(APP_TITLE, "수정한 내용이 없습니다.\n칸을 더블클릭해서 고치세요.")
            return False
        lines = []
        for e in self.edits.values():
            by = {c["col"]: c for c in self.cfg[e["p"]]["columns"]}
            for col, (old, new) in e["changes"].items():
                nm = by[col]["name"] if col in by else col
                lines.append(f"{e['label']} · {nm}: '{cell_text(old)}' → '{cell_text(new)}'")
        more = f"\n... 외 {len(lines) - 25}건" if len(lines) > 25 else ""
        if not messagebox.askyesno("MES 수정", "아래 내용을 MES에 바로 반영합니다.\n\n" + "\n".join(lines[:25]) + more
                                   + "\n\n계속할까요?"):
            return False
        edits = list(self.edits.values())
        try:
            done = self.run_db(lambda c: self.mes.apply_updates(c, edits))
        except Exception as ex:
            self.err(ex)
            return False
        now = dt.datetime.now()
        rows = []
        for e in done:
            by = {c["col"]: c for c in self.cfg[e["p"]]["columns"]}
            txt = "; ".join(f"{by[col]['name'] if col in by else col}: {cell_text(o)} -> {cell_text(n)}"
                            for col, (o, n) in e["changes"].items())
            rows.append([f"{now:%Y-%m-%d %H:%M:%S}", self.cfg.get("reg_user", ""), f"{PART_NAME[e['p']]}수정",
                         self.mes.table(e["p"]), "/".join(show(v) for v in e["key"].values()), txt])
        self.append_log(rows)
        self.load_log()
        self.edits.clear()
        self.update_edit_btn()
        messagebox.showinfo(APP_TITLE, f"{len(done)}줄 수정 완료.\n[작업기록] 탭에 기록했습니다.")
        code = self.item_code
        try:
            self.old_cols, self.old_rows, self.old_defs = self.run_db(self.mes.fetch_old)
            self.old_time = dt.datetime.now()
            self.chk_old.clear()
            self.show_old()
            if code is not None:
                self.item_cols, self.item_rows, self.item_defs = self.run_db(lambda c: self.mes.fetch_items_of(c, code))
                self.show_items()
        except Exception as ex:
            self.err(ex)
        return True

    def on_set_ref(self):
        codes = self.selected_codes()
        if len(codes) != 1:
            messagebox.showinfo(APP_TITLE, "기준으로 쓸 금형 한 개를 선택하세요.\n"
                                           "(정상적으로 작업지시가 되는 금형을 고르세요)")
            return
        self.cfg["ref_mold"] = show(codes[0])
        self.set_vars["ref_mold"].set(self.cfg["ref_mold"])
        save_cfg(self.cfg)
        self.show_old()
        messagebox.showinfo(APP_TITLE, f"기준 금형: {self.cfg['ref_mold']}\n\n"
                                       "기본값이 {COPY}인 칸(회사코드 등)은 이 금형의 값을 그대로 씁니다.")

    def on_copy_old(self):
        codes = self.selected_codes()
        if not codes:
            messagebox.showinfo(APP_TITLE, "복사할 금형을 선택하세요.")
            return
        molds = [dict(zip(self.old_cols, self.old_rows[int(i)])) for i in self.tv_old.selection()]
        path = self.ask_save(f"금형등록_{dt.datetime.now():%Y%m%d_%H%M}.xlsx")
        if not path:
            return
        conn = self.connect_optional()
        try:
            items = []
            if conn is not None and self.mes.enabled("item"):
                for code in codes:
                    cols, rows, _ = self.mes.fetch_items_of(conn, code)
                    items += [dict(zip(cols, r)) for r in rows]
            self.busy(True)
            self.xl.make_template(path, prefill={"mold": molds, "item": items}, conn=conn)
            self.busy(False)
        except Exception as e:
            self.err(e)
            return
        finally:
            if conn is not None:
                conn.close()
        open_file(path)
        self.on_load(path, quiet=True)
        self.show_tab("new")
        messagebox.showinfo(APP_TITLE, f"금형 {len(molds)}건, 품번 {len(items)}건을 양식에 복사했습니다.\n\n"
                                       "엑셀에서 금형 시트와 품번 시트의 금형코드(주황)를 새 코드로 입력하고\n"
                                       "저장한 뒤 [① 검사]를 누르세요.")

    # ==========================================================
    # 컬럼매핑 탭
    # ==========================================================
    def build_map(self):
        t = self.tabs["map"]
        self.map_part = tk.StringVar(value="mold")
        self.map_data = {p: copy.deepcopy(self.cfg[p]["columns"]) for p in PARTS}
        self.map_meta = {p: {"table": self.cfg[p].get("table", ""),
                             "codecol": self.cfg[p].get("code_col" if p == "mold" else "link_col", ""),
                             "enabled": self.cfg[p].get("enabled", True)} for p in PARTS}
        self.cur_part = "mold"

        top = ttk.Frame(t)
        top.pack(fill="x", pady=(8, 2))
        for p in PARTS:
            ttk.Radiobutton(top, text=f"{PART_NAME[p]} 테이블", value=p, variable=self.map_part,
                            command=self.on_map_part).pack(side="left", padx=6)
        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Label(top, text="테이블").pack(side="left")
        self.var_tbl = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_tbl, width=24, state="readonly").pack(side="left", padx=4)
        self.lbl_code = ttk.Label(top, text="금형코드 컬럼")
        self.lbl_code.pack(side="left", padx=(10, 0))
        self.var_codecol = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_codecol, width=16, state="readonly").pack(side="left", padx=4)
        self.var_item_on = tk.BooleanVar()
        self.chk_item = ttk.Checkbutton(top, text="품번 테이블도 등록", variable=self.var_item_on)
        self.chk_item.pack_forget()

        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=(4, 4))
        ttk.Button(bar, text="DB컬럼 가져오기", style="Big.TButton", command=self.on_fetch_cols).pack(side="left", padx=3)
        ttk.Button(bar, text="▲", width=3, command=lambda: self.move_map(-1)).pack(side="left", padx=(12, 2))
        ttk.Button(bar, text="▼", width=3, command=lambda: self.move_map(1)).pack(side="left", padx=2)
        ttk.Button(bar, text="저장", style="Big.TButton", command=self.on_save_map).pack(side="right", padx=3)
        ttk.Label(t, text="칸 더블클릭으로 수정 (Y/N·형식은 더블클릭하면 바뀜).  기본값: SYSDATE / {USER}=등록자 / "
                          "{COPY}=기준 금형 값 복사 / 고정값.   존재확인 예: ICOM_ITEM.ITEM_CODE",
                  foreground="#555").pack(fill="x", padx=4)
        fr, self.tv_map = self.make_tree(t, [f[1] for f in MAP_FIELDS], [f[2] for f in MAP_FIELDS])
        fr.pack(fill="both", expand=True, pady=4)
        self.tv_map.bind("<Double-1>", self.on_map_edit)
        self.load_map_part("mold")

    def stash_map_part(self):
        p = self.cur_part
        self.map_meta[p]["table"] = self.var_tbl.get().strip().upper()
        self.map_meta[p]["codecol"] = self.var_codecol.get().strip().upper()
        if p == "item":
            self.map_meta[p]["enabled"] = self.var_item_on.get()

    def load_map_part(self, p):
        self.cur_part = p
        m = self.map_meta[p]
        self.var_tbl.set(m["table"])
        self.var_codecol.set(m["codecol"])
        self.lbl_code.config(text="금형코드 컬럼" if p == "mold" else "금형코드 연결 컬럼")
        self.var_item_on.set(bool(self.map_meta["item"]["enabled"]))
        self.chk_item.state(["!disabled"] if p == "item" else ["disabled"])
        self.show_map()

    def on_map_part(self):
        self.stash_map_part()
        self.load_map_part(self.map_part.get())

    def show_map(self, select=None):
        tv = self.tv_map
        tv.delete(*tv.get_children())
        cc = self.var_codecol.get().strip().upper()
        for i, c in enumerate(self.map_data[self.cur_part]):
            vals = []
            for key, _, _, kind in MAP_FIELDS:
                v = c.get(key)
                if kind == "yn":
                    vals.append("Y" if v else "N")
                elif kind == "int":
                    vals.append(str(v) if v else "")
                else:
                    vals.append(v or "")
            tags = ("ok",) if c["col"] == cc else (() if c["use"] else ("off",))
            tv.insert("", "end", iid=str(i), values=vals, tags=tags)
        if select is not None:
            tv.selection_set(str(select))
            tv.see(str(select))

    def on_map_edit(self, e):
        tv = self.tv_map
        iid = tv.identify_row(e.y)
        colid = tv.identify_column(e.x)
        if not iid or not colid:
            return
        i, j = int(iid), int(colid[1:]) - 1
        if j < 0 or j >= len(MAP_FIELDS):
            return
        key, title, _, kind = MAP_FIELDS[j]
        c = self.map_data[self.cur_part][i]
        if kind == "yn":
            c[key] = not c.get(key)
        elif kind == "type":
            order = ["TEXT", "NUM", "DATE"]
            c[key] = order[(order.index(c.get(key, "TEXT")) + 1) % 3] if c.get(key) in order else "TEXT"
        else:
            v = simpledialog.askstring(title, f"{c['col']} - {title}", initialvalue=str(c.get(key) or ""),
                                       parent=self.root)
            if v is None:
                return
            v = v.strip()
            if kind == "int":
                c[key] = int(v) if v.isdigit() else 0
            elif key in ("col", "ref"):
                c[key] = v.upper()
            else:
                c[key] = v
        self.show_map(select=i)

    def move_map(self, d):
        sel = self.tv_map.selection()
        if not sel:
            return
        lst = self.map_data[self.cur_part]
        i = int(sel[0])
        k = i + d
        if 0 <= k < len(lst):
            lst[i], lst[k] = lst[k], lst[i]
            self.show_map(select=k)

    def on_fetch_cols(self, table=None):
        p = self.cur_part
        table = (table or self.var_tbl.get()).strip().upper()
        if not table:
            messagebox.showinfo(APP_TITLE, f"{PART_NAME[p]} 테이블명을 입력하거나 [테이블 찾기]를 하세요.")
            return
        if not messagebox.askyesno(APP_TITLE, f"{PART_NAME[p]} 매핑을 {table} 구조로 덮어씁니다. 계속할까요?"):
            return
        try:
            cols, code = self.run_db(lambda c: self.mes.fetch_columns(c, table, p))
        except Exception as e:
            self.err(e)
            return
        if not cols:
            messagebox.showwarning(APP_TITLE, "테이블을 찾지 못했습니다. 스키마/테이블명을 확인하세요.")
            return
        self.var_tbl.set(table)
        self.map_data[p] = cols
        if code:
            self.var_codecol.set(code)
        self.show_map()
        auto = [f"{c['col']}={c['default']}" for c in cols if not c["use"] and c["default"]]
        msg = f"{table}: {len(cols)}개 컬럼을 불러왔습니다.\n\n"
        msg += f"금형코드 컬럼: {code or '못 찾음 - 직접 입력하세요'}\n"
        if auto:
            msg += "\n자동으로 기본값 처리한 컬럼 (입력칸에서 뺌):\n - " + "\n - ".join(auto[:12])
            if len(auto) > 12:
                msg += f"\n - 외 {len(auto) - 12}개"
        msg += ("\n\n정리할 것:\n- 입력 안 할 컬럼은 사용 N (NOT NULL이면 기본값 필요)\n"
                "- {COPY}는 기준 금형 값을 복사합니다 ([금형·품번 관리] 탭에서 지정)\n"
                "- 끝나면 [저장]")
        messagebox.showinfo(APP_TITLE, msg)

    def on_save_map(self):
        self.stash_map_part()
        warns = []
        for p in PARTS:
            m = self.map_meta[p]
            lst = self.map_data[p]
            if p == "item" and not m["enabled"]:
                continue
            try:
                ident(m["table"], f"{PART_NAME[p]} 테이블")
                for c in lst:
                    ident(c["col"], "DB컬럼")
                    if c["ref"]:
                        ident(c["ref"], "존재확인")
            except ValueError as e:
                messagebox.showerror(APP_TITLE, f"[{PART_NAME[p]}] {e}")
                return
            cc = m["codecol"]
            c_code = next((c for c in lst if c["col"] == cc), None)
            if not c_code or not c_code["use"]:
                messagebox.showerror(APP_TITLE, f"[{PART_NAME[p]}] 금형코드 컬럼({cc or '미지정'})이 목록에 없거나 사용 N입니다.")
                return
            miss = [c["col"] for c in lst if not c["use"] and c["req"] and not c["default"]]
            if miss:
                warns.append(f"[{PART_NAME[p]}] 필수인데 사용 N·기본값 없음: " + ", ".join(miss))
            if not any(c["use"] and c["key"] for c in lst):
                warns.append(f"[{PART_NAME[p]}] 중복검사키가 없습니다.")
        for p in PARTS:
            self.cfg[p]["columns"] = copy.deepcopy(self.map_data[p])
            self.cfg[p]["table"] = self.map_meta[p]["table"]
        self.cfg["mold"]["code_col"] = self.map_meta["mold"]["codecol"]
        self.cfg["item"]["link_col"] = self.map_meta["item"]["codecol"]
        self.cfg["item"]["enabled"] = bool(self.map_meta["item"]["enabled"])
        save_cfg(self.cfg)
        msg = "저장했습니다.\n예전에 만든 엑셀 양식은 [엑셀 양식 만들기]로 새로 만드세요."
        if warns:
            msg += "\n\n※ 확인 필요 (DB에 기본값이 없으면 등록 실패):\n" + "\n".join(warns)
        messagebox.showinfo(APP_TITLE, msg)
        self.path = None
        self.data = {p: {"rows": [], "info": None, "results": []} for p in PARTS}
        self.show_new()

    # ==========================================================
    # 설정 탭
    # ==========================================================
    SET_FIELDS = [
        ("oracle.dsn", "접속 주소(DSN)", "예: 192.168.1.250:1521/XE"),
        ("oracle.lib_dir", "Instant Client 폴더", "11g는 필수 (예: C:\\instantclient_19_30)"),
        ("schema", "스키마", "MES 스키마"),
        ("reg_user", "등록자", "기본값 {USER}에 들어갈 이름"),
        ("ref_mold", "기준 금형코드", "엑셀 대량등록에서 모르는 칸을 이 금형 값으로 채움 (보통 비워 둠)"),
        ("mold.order_by", "기존금형 정렬", "예: MOLD_CODE"),
        ("kor_bytes", "한글 1자 byte", "DB 문자셋 UTF-8=3, KO16MSWIN949=2"),
        ("limit", "조회 최대건수", "기존금형·품목·작업지시 한 번에 가져올 건수"),
        ("item_master_table", "품목 테이블", "품목 탭 (기본 ICOM_ITEM_MASTER)"),
        ("bom_table", "단위중량 테이블", "단위중량 탭 (기본 ICOM_ITEM_CHILD)"),
        ("keywords", "구조 조사 찾을 단어", "쉼표로 구분. 테이블/프로그램 이름에 이 단어가 있으면 찾음"),
        ("small_limit", "전후 비교 줄 단위 한도", "이 줄 수 이하 테이블은 새 줄·바뀐 줄 내용까지 비교"),
        ("wo_table", "작업지시 테이블", "작업지시 탭 (기본 IPLN_WORK_ORDER_MASTER)"),
        ("rcv_weight_unit", "소재 순중량 단위", "auto / kg / g   (auto = 최근 입고 값 크기로 자동 판단)"),
    ]

    def build_set(self):
        box = ttk.Frame(self.tabs["set"], padding=16)
        box.pack(fill="both", expand=True)
        self.set_vars = {}
        for r, (path, label, hint) in enumerate(self.SET_FIELDS):
            ttk.Label(box, text=label, font=("", 10, "bold")).grid(row=r, column=0, sticky="w", pady=5)
            v = tk.StringVar(value=str(self.get_cfg(path)))
            ttk.Entry(box, textvariable=v, width=46).grid(row=r, column=1, sticky="w", padx=8)
            ttk.Label(box, text=hint, foreground="#777").grid(row=r, column=2, sticky="w")
            self.set_vars[path] = v
        n = len(self.SET_FIELDS)
        ttk.Label(box, text=f"접속 계정: {DB_USER} (고정)   금형 테이블: ICOM_MOLD / 품번 테이블: ICOM_MOLD_ITEM (고정)",
                  foreground="#777").grid(row=n, column=1, sticky="w", pady=(10, 4))
        self.var_dev = tk.BooleanVar(value=bool(self.cfg.get("show_dev_tabs")))

        def _dev():
            self.cfg["show_dev_tabs"] = self.var_dev.get()
            save_cfg(self.cfg)
            self.apply_dev_tabs()
        ttk.Checkbutton(box, text="고급 탭 보이기 (MES 구조조사 · 전후 비교 - 보통은 필요 없음)",
                        variable=self.var_dev, command=_dev).grid(row=n + 2, column=1, columnspan=2, sticky="w", pady=4)
        self.var_regtab = tk.BooleanVar(value=bool(self.cfg.get("show_reg_tabs")))

        def _reg():
            self.cfg["show_reg_tabs"] = self.var_regtab.get()
            save_cfg(self.cfg)
            self.apply_dev_tabs()
        ttk.Checkbutton(box, text="금형 등록 탭 보이기 (금형 간편등록 · 엑셀 대량등록 · 엑셀칸 설정)",
                        variable=self.var_regtab, command=_reg).grid(row=n + 3, column=1, columnspan=2, sticky="w", pady=4)
        self.var_itemtab = tk.BooleanVar(value=bool(self.cfg.get("show_item_tab")))

        def _itm():
            self.cfg["show_item_tab"] = self.var_itemtab.get()
            save_cfg(self.cfg)
            self.apply_dev_tabs()
        ttk.Checkbutton(box, text="품목 관리 탭 보이기 (품목 마스터 보기·수정 - 보통은 ERP에서 관리)",
                        variable=self.var_itemtab, command=_itm).grid(row=n + 4, column=1, columnspan=2, sticky="w", pady=4)
        bar = ttk.Frame(box)
        bar.grid(row=n + 1, column=1, columnspan=2, sticky="w", pady=8)
        ttk.Button(bar, text="저장", style="Big.TButton", command=self.on_save_set).pack(side="left", padx=(8, 4))
        ttk.Button(bar, text="연결 테스트", style="Big.TButton", command=self.on_test).pack(side="left", padx=4)

    def get_cfg(self, path):
        d = self.cfg
        for k in path.split("."):
            d = d.get(k, "") if isinstance(d, dict) else ""
        return d

    def apply_set(self):
        for path, v in self.set_vars.items():
            val = v.get().strip()
            if path in ("kor_bytes", "limit", "small_limit"):
                val = int(val) if val.isdigit() else DEFAULT_CFG[path]
            keys = path.split(".")
            if len(keys) == 2:
                self.cfg[keys[0]][keys[1]] = val
            else:
                self.cfg[path] = val

    def on_save_set(self):
        self.apply_set()
        save_cfg(self.cfg)
        messagebox.showinfo(APP_TITLE, "저장했습니다.")

    def on_test(self):
        self.apply_set()

        def fn(conn):
            cur = conn.cursor()
            lines = ["접속 성공."]
            for p in PARTS:
                if not self.mes.enabled(p) or not self.cfg[p].get("table"):
                    continue
                try:
                    cur.execute(f"SELECT COUNT(*) FROM {self.mes.table(p)}")
                    lines.append(f"{PART_NAME[p]}: {self.mes.table(p)} {cur.fetchone()[0]:,}건")
                except Exception as e:
                    lines.append(f"{PART_NAME[p]}: {self.mes.table(p)} 조회 실패 - {e}")
            return "\n".join(lines)
        try:
            messagebox.showinfo(APP_TITLE, self.run_db(fn))
        except Exception as e:
            self.err(e)

    # ---------- 테이블 찾기 ----------
    def on_find_table(self):
        self.apply_set()
        try:
            found = self.run_db(self.mes.find_mold_tables)
        except Exception as e:
            self.err(e)
            return
        if not found:
            messagebox.showwarning(APP_TITLE, "금형 관련 테이블을 찾지 못했습니다.\n스키마 이름을 확인하세요.")
            return
        win = tk.Toplevel(self.root)
        win.title("금형 테이블 찾기")
        win.geometry("1000x540")
        win.transient(self.root)
        ttk.Label(win, text="★ = 유력.  행수는 실제 개수.  더블클릭 = 미리보기.  "
                            "[금형 테이블로] / [품번 테이블로] 로 지정하세요.",
                  foreground="#555").pack(fill="x", padx=8, pady=(8, 4))
        cols = ["추천", "테이블", "설명", "기본키(PK)", "참조수", "컬럼수", "행수"]
        fr, tv = self.make_tree(win, cols, [50, 220, 220, 160, 60, 60, 110])
        fr.pack(fill="both", expand=True, padx=8)
        top = found[0]["score"]
        cur_t = {self.cfg[p].get("table", "").upper(): PART_NAME[p] for p in PARTS}
        for i, f in enumerate(found):
            star = "★" if f["score"] >= 50 or (f["score"] == top and top > 0) else ""
            if f["table"].upper() in cur_t:
                star = f"[{cur_t[f['table'].upper()]}]"
            tv.insert("", "end", iid=str(i), tags=("ok",) if star else (),
                      values=[star, f["table"], f["cmt"], f["pk"], f["refcnt"], f["colcnt"],
                              "" if f["rows"] is None else f"{f['rows']:,}" + ("" if f["exact"] else " (추정)")])
        tv.selection_set("0")

        def picked():
            sel = tv.selection()
            return found[int(sel[0])]["table"] if sel else None

        def do_preview():
            t = picked()
            if not t:
                return
            try:
                hdr, rows = self.run_db(lambda c: self.mes.preview(c, t))
                self.show_preview(win, t, hdr, rows)
            except Exception as e:
                messagebox.showerror(APP_TITLE, f"미리보기 오류\n{type(e).__name__}: {e}", parent=win)

        def do_use(p):
            t = picked()
            if not t:
                return
            win.destroy()
            self.cfg[p]["table"] = t
            if p == "item":
                self.cfg["item"]["enabled"] = True
                self.map_meta["item"]["enabled"] = True
            self.map_meta[p]["table"] = t
            save_cfg(self.cfg)
            self.show_tab("map")
            self.stash_map_part()
            self.map_meta[p]["table"] = t
            self.map_part.set(p)
            self.load_map_part(p)
            self.on_fetch_cols(t)

        tv.bind("<Double-1>", lambda _e: do_preview())
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=8)
        ttk.Button(bar, text=f"미리보기 ({PREVIEW_ROWS}행)", style="Big.TButton", command=do_preview).pack(side="left", padx=3)
        ttk.Button(bar, text="금형 테이블로", style="Big.TButton", command=lambda: do_use("mold")).pack(side="left", padx=3)
        ttk.Button(bar, text="품번 테이블로", style="Big.TButton", command=lambda: do_use("item")).pack(side="left", padx=3)
        ttk.Button(bar, text="닫기", command=win.destroy).pack(side="right", padx=3)

    def show_preview(self, parent, t, hdr, data):
        pw = tk.Toplevel(parent)
        pw.title(f"미리보기 - {t}")
        pw.geometry("1150x620")
        pw.transient(parent)
        bar = ttk.Frame(pw)
        bar.pack(fill="x", padx=8, pady=(8, 0))
        ttk.Label(bar, text="검색").pack(side="left")
        qv = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=qv, width=30)
        ent.pack(side="left", padx=4)
        lbl = ttk.Label(bar, text="", font=("", 10, "bold"))
        lbl.pack(side="left", padx=10)
        ttk.Label(bar, text="줄 더블클릭 = 세로로 자세히", foreground="#777").pack(side="right")
        fr, tv = self.make_tree(pw, [str(h) for h in hdr], [max(90, min(220, len(str(h)) * 9 + 20)) for h in hdr])
        fr.pack(fill="both", expand=True, padx=8, pady=8)
        more = "  (최대 표시 수 도달 - 더 있을 수 있음)" if len(data) >= PREVIEW_ROWS else ""

        def refill(*_):
            tv.delete(*tv.get_children())
            q = qv.get().strip().lower()
            n = bad = 0
            for i, vals in enumerate(data):
                if q and not any(q in v.lower() for v in vals):
                    continue
                try:
                    tv.insert("", "end", iid=str(i), values=vals)
                    n += 1
                except Exception:
                    bad += 1
            txt = f"{len(data):,}행 불러옴" if data else "0행 - 이 계정으로 읽을 수 있는 데이터가 없습니다"
            if q:
                txt += f" / 검색결과 {n:,}행"
            if bad:
                txt += f" / 표시 실패 {bad}행"
            lbl.config(text=txt + more, foreground="#C00000" if (not data or bad) else "#1F4E79")

        def detail(_e):
            sel = tv.selection()
            if not sel:
                return
            vals = data[int(sel[0])]
            dw = tk.Toplevel(pw)
            dw.title(f"{t} - {int(sel[0]) + 1}번째 줄")
            dw.geometry("620x640")
            dw.transient(pw)
            dfr, dtv = self.make_tree(dw, ["컬럼", "값"], [200, 390])
            dfr.pack(fill="both", expand=True, padx=8, pady=8)
            for h, v in zip(hdr, vals):
                dtv.insert("", "end", values=[str(h), v])

        tv.bind("<Double-1>", detail)
        qv.trace_add("write", refill)
        refill()
        pw.lift()
        pw.focus_force()
        ent.focus_set()

    # ==========================================================
    # 사용법 탭
    # ==========================================================
    HELP = [
        ("h1", "태진다이텍 MES 관리 도구 사용법"),
        ("p", "MES(오라클)에 있는 금형·품번 연결, 피치당 소재중량, 작업지시, 소재 입고를 보고 고치는 프로그램입니다.\n"
              "새로 만들 때는 기존 줄을 바꾸지 않고 새 줄만 추가합니다. 모든 작업은 [작업기록]에 남습니다."),

        ("h2", "0. 처음 한 번"),
        ("li", "[환경설정] 탭 → [연결 테스트]. 접속 계정과 금형(ICOM_MOLD)·품번(ICOM_MOLD_ITEM) 테이블은 고정되어 있습니다."),
        ("li", "[등록자]에 이름을 넣고 [저장]하면 수정·등록한 사람으로 기록됩니다."),

        ("h2", "[① 금형·품번 관리] 탭 - 위 = 금형, 아래 = 선택한 금형의 품번"),
        ("li", "[MES에서 조회]. 검색칸에 글자를 넣으면 모든 칸에서 찾습니다. 금형을 클릭하면 아래에 연결 품번이 나옵니다."),
        ("li", "맨 앞 ☐ 을 누르면 ☑ 체크 (머리글 ☐ = 보이는 줄 전체). 삭제·복사는 체크한 줄(없으면 선택한 줄)이 대상입니다."),
        ("li", "고치기: [수정 모드] 체크 → 위·아래 표의 칸 더블클릭 → 입력 → Enter → [수정 저장].  주황 = 아직 저장 안 됨, [수정 취소] = 되돌리기."),
        ("li", "   칸이 많으면 [선택 금형 전체컬럼 수정] / [선택 품번 전체컬럼 수정]으로 세로 목록에서 고치면 편합니다."),
        ("li", "   금형코드는 작업지시·샷수 이력과 이어져 있어 여기서는 고칠 수 없습니다 → 아래 '다른 이름으로 저장 모드'를 쓰세요."),
        ("li", "새 금형 만들기(복사): 금형 1개 체크 → [다른 이름으로 저장 모드] 체크 → 위·아래 칸(금형코드·금형명·품번 포함) 더블클릭해 고치기"),
        ("li", "   → [다른 이름으로 저장] → 창에서 확인 → [MES에 등록].  원본 금형·품번은 그대로이고 새 금형+품번이 한꺼번에 추가됩니다."),
        ("li", "   품번 하나는 금형 하나에만 연결되므로 복사하는 품번은 새 품번이어야 합니다 (빨강 = 이미 다른 금형에 연결됨)."),
        ("li", "   노랑 = 품목 마스터(ERP)에 없는 품번. 맞으면 [MES에 등록]을 한 번 더 누르면 저장됩니다."),
        ("li", "   이 모드에서는 원본을 지키기 위해 [수정 저장]과 삭제가 막힙니다. 모드를 끄면 고쳐 놓은 내용은 버려집니다."),
        ("li", "품번 연결 끊기: [수정 모드] → 아래 표에서 ☑ 체크 → [선택 품번줄 삭제]. 금형·품목 자체는 남고 연결만 지워집니다."),
        ("li", "   연결을 끊은 품번으로 새로 내는 작업지시는 금형이 '*'(없음)로 들어갑니다. 이미 낸 작업지시는 그대로입니다."),
        ("li", "금형 삭제: [수정 모드] → 위 표에서 ☑ 체크 → [선택 금형 삭제]. 진행 창에서 사용 이력을 확인합니다 (취소 가능)."),
        ("li", "   작업지시·샷수 등 이력이 있는 금형은 지워지지 않습니다 (이력 보호). 확인 못한 큰 표가 있으면 알려줍니다."),

        ("h2", "[② 단위중량 관리] 탭 - 피치당 소재중량(g)"),
        ("li", "탭을 열면 전체 목록이 자동으로 조회됩니다. ◉ 모두 / ○ 단위중량 0/없음만 중에서 고를 수 있습니다."),
        ("li", "피치당 소재중량(g) = 두께 × 폭 × 피치 × 7.85 ÷ 1000.  예) 두께 1.0, 폭 126, 피치 122 → 약 120.7g."),
        ("li", "   제품 순수무게가 아닙니다 (가운데 구멍·스크랩 포함). 캐비티가 2 이상이면 이 값 ÷ 캐비티를 넣습니다."),
        ("li", "현장에서 1샷마다 '샷수 × 캐비티 × 이 값'으로 코일 사용량이 계산됩니다. 없거나 0이면 소재가 빠지지 않습니다 (노랑)."),
        ("li", "원소재(코일) 코드는 소재 입고에 쓰는 코일 품번과 글자가 똑같아야 그 코일에서 빠집니다."),
        ("li", "고치기: [수정 모드] → 칸 더블클릭 → [수정 저장]. 이전 값은 MES가 이력에 자동 보관합니다."),
        ("li", "새로 넣기: [수정 모드] → (비슷한 줄 선택) → [새로 등록] → 제품 품번·원소재 코드·피치당 소재중량 입력."),
        ("li", "검색 결과가 0건이면 금형 연결·품목 마스터·최근 생산 기록을 확인해서 알려줍니다."),
        ("li", "한꺼번에: [엑셀로 저장] → 엑셀에서 고치기 → [엑셀로 일괄 수정].  [최근 60일 단위중량 0 생산 점검] = 빠진 조합 찾기."),

        ("h2", "[③ 작업지시 관리] 탭 - 새로 만들기 / 고치기 / 확인"),
        ("li", "[최근 작업지시 불러오기]: 적용할 작업일자 선택 → 정상 8h / 잔업까지 10h / 시간 직접 입력 → 최근 30일에서 호기별 최종 작업지시 1건만 불러와 [⑤ SPM 관리]의 SPM(없는 품번은 최근 실적 SPM)×시간×캐비티로 계획수량 자동 계산 → 팝업에서 수정·삭제 → 검사·번호생성 후 최종 반영."),
        ("li", "고치기: 줄 더블클릭 또는 [선택 작업지시 고치기] → 지시일·근무조·설비·품번·수량·상태 → [수정 저장]. 실적이 있으면 경고합니다."),
        ("li", "조회: 지시일(달력)·설비·품번으로 [MES에서 조회].  빨강 = 금형 없음(*), 노랑 = 품번 연결과 금형 다름, 하늘 = 같은 날 중복."),
        ("li", "삭제: [수정 모드] → 줄 선택 → [선택 줄 삭제]. 실적이 있는 작업지시는 지워지지 않습니다."),
        ("li", "[여러 건 일괄등록]: 설비명·품번·계획수량만 넣으면 됩니다 (빠른 입력 줄에서 Enter, 또는 엑셀 3칸 복사 → 표에 Ctrl+V). 근무는 정상 8h / 잔업 10h / 직접 h 입력."),
        ("li", "   계획일자·교대조는 창 위 기본값, 순위는 설비별 순서대로 자동. MES 양식(라인번호·일자·교대조·품목·수량·순위)도 붙여넣기 됩니다."),
        ("li", "   → [검사] → [W/O 번호생성] → [작업지시 반영]. MES '작업지시일괄처리'와 같은 방법(업로더 + 같은 MES 프로그램)으로 넣습니다."),
        ("li", "   설비는 설비명(예: 28_250)·설비코드(예: 005)·호기 번호 모두 됩니다. 품목 마스터에 없는 품번은 MES도 빼므로 빨강으로 막습니다."),

        ("h2", "[④ 소재 관리] 탭 - 소재(코일) 입고 조회 / 새 입고 / 고치기 / 삭제"),
        ("li", "조회: 입고일 [달력] 또는 [오늘]·[최근 7일]·[이번 달] 버튼. 품번·HEAT_NO·입고번호·업체는 일부만 넣어도 찾습니다."),
        ("li", "[+ 새 입고]: 입고일 → 품번 → HEAT_NO → 업체 → 중량(kg) → 수량(롤) → [MES에 등록]. 입고 + 재고LOT 추가, 소재 재고에 더하기."),
        ("li", "   코일 여러 개: '등록 후 창 유지'를 켜면 HEAT_NO·중량만 바꿔 계속 입력할 수 있습니다."),
        ("li", "고치기: 줄 더블클릭 → [수정 저장]. 입고·재고LOT를 같이 고치고 재고도 차이만큼 맞춥니다. 사용·출고된 입고는 업체·담당자만."),
        ("li", "삭제: [삭제 모드] → 초록 줄 선택 → [선택 입고 삭제] → '삭제' 입력. 재고에서도 뺍니다 (취소된 입고는 다시 빼지 않음)."),
        ("li", "초록 = 수정·삭제 가능, 회색 = 생산에 사용·출고됨, 노랑 = 취소(반품)된 입고.  중량은 kg로 넣고 kg로 봅니다."),

        ("h2", "[⑤ SPM 관리] 탭 - 품번별 SPM(분당 타수)"),
        ("li", "작업지시 계획수량 = SPM × 60 × 근무시간 × 캐비티. 여기 없는 품번은 최근 생산 기록의 SPM을 씁니다."),
        ("li", "[+ 새로 등록] / 줄 더블클릭 = 고치기 / [선택 줄 삭제].  [SPM 없는 품번 찾기] = 최근 30일 작업 품번 중 빠진 것."),
        ("li", "프로그램 폴더의 spm.xlsx 에 저장됩니다 (MES는 안 바뀜). [엑셀에서 가져오기] = A열 품번·B열 SPM 파일로 한꺼번에 추가·변경."),

        ("h2", "색깔 뜻"),
        ("li", "표: 주황 = 고쳤지만 아직 저장 안 됨,  빨강 = 오류·확인 필요,  노랑 = 경고,  초록 = 정상"),
        ("li", "버튼: 초록 = MES에 등록·저장 (실제로 MES가 바뀜),  파랑 = 조회·찾기 (보기만),  청록 = 엑셀,  빨강 = 삭제"),

        ("h2", "숨겨 둔 탭 ([환경설정]에서 체크하면 나타남)"),
        ("li", "금형 등록 탭: 금형 간편등록 · 엑셀 대량등록 · 엑셀칸 설정 (엑셀로 여러 금형을 한꺼번에 넣을 때)."),
        ("li", "품목 관리 탭: 품목 마스터 보기·수정 (보통은 ERP에서 관리)."),
        ("li", "고급 탭: MES 구조조사 · 전후 비교 (MES 안쪽 구조 확인용, 조회만)."),

        ("h2", "자주 나오는 메시지"),
        ("li", "이미 있습니다 / 중복 불가 [UIX_...] → 같은 금형코드나 같은 품번 연결이 이미 MES에 있습니다."),
        ("li", "그 사이 다른 곳에서 바뀌었거나 없어진 줄 → MES 화면·현장에서 먼저 바뀌었습니다. 다시 조회한 뒤 고치세요."),
        ("li", "오류 → 전체 취소 → MES에는 아무것도 바뀌지 않았습니다. 메시지를 보고 고친 뒤 다시 하세요."),

        ("h2", "안전장치"),
        ("li", "새로 만들기는 새 줄 추가(INSERT)만 합니다. 수정은 바꾼 칸만, 그 사이 바뀌었으면 저장하지 않습니다."),
        ("li", "한 묶음 중 하나라도 실패하면 전부 취소합니다. 지운 내용은 mold_deleted_backup.jsonl 에 백업됩니다."),
        ("li", "파일 위치 (프로그램과 같은 폴더): mold_config.json = 설정,  mold_register_log.csv = 작업기록,  mold_deleted_backup.jsonl = 삭제 백업"),
    ]

    def build_help(self):
        t = self.tabs["help"]
        fr = ttk.Frame(t)
        fr.pack(fill="both", expand=True, padx=6, pady=6)
        txt = tk.Text(fr, wrap="word", padx=18, pady=14, relief="flat", background="#FFFFFF",
                      font=("맑은 고딕", 10), spacing1=2, spacing3=2)
        ys = ttk.Scrollbar(fr, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=ys.set)
        txt.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        txt.tag_configure("h1", font=("맑은 고딕", 15, "bold"), foreground="#1F4E79", spacing3=8)
        txt.tag_configure("h2", font=("맑은 고딕", 11, "bold"), foreground="#1F4E79", spacing1=14, spacing3=4)
        txt.tag_configure("p", foreground="#333333", spacing3=4)
        txt.tag_configure("li", lmargin1=14, lmargin2=30)
        for kind, text in self.HELP:
            if kind == "li":
                txt.insert("end", "•  " + text + "\n", "li")
            else:
                txt.insert("end", text + "\n", kind)
        txt.configure(state="disabled")

    # ==========================================================
    # 등록이력 탭
    # ==========================================================
    LOG_COLS = ["일시", "등록자", "구분", "테이블", "키", "입력내용"]

    def build_log(self):
        t = self.tabs["log"]
        bar = ttk.Frame(t)
        bar.pack(fill="x", pady=(6, 4))
        ttk.Button(bar, text="새로고침", command=self.load_log).pack(side="left", padx=3)
        ttk.Button(bar, text="파일 열기", command=lambda: open_file(LOG_PATH)).pack(side="left", padx=3)
        fr, self.tv_log = self.make_tree(t, self.LOG_COLS, [140, 80, 70, 220, 160, 800])
        fr.pack(fill="both", expand=True, pady=4)
        self.load_log()

    def append_log(self, rows):
        if os.path.exists(LOG_PATH):
            with open(LOG_PATH, encoding="utf-8-sig") as f:
                head = next(csv.reader(f), [])
            if head != self.LOG_COLS:
                os.replace(LOG_PATH, LOG_PATH.replace(".csv", "_v1.csv"))
        new = not os.path.exists(LOG_PATH)
        with open(LOG_PATH, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new:
                w.writerow(self.LOG_COLS)
            w.writerows(rows)

    def write_log(self, done, now):
        rows = []
        for p, r, log in done:
            k = "/".join(show(r["vals"].get(c["col"])) for c in self.mes.keys(p))
            rows.append([f"{now:%Y-%m-%d %H:%M:%S}", self.cfg.get("reg_user", ""), PART_NAME[p] + "등록",
                         self.mes.table(p), k, log])
        self.append_log(rows)

    def load_log(self):
        self.tv_log.delete(*self.tv_log.get_children())
        if not os.path.exists(LOG_PATH):
            return
        with open(LOG_PATH, encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))[1:]
        for row in reversed(rows):
            self.tv_log.insert("", "end", values=row)


def main():
    root = tk.Tk()

    def on_error(exc, val, tb):          # 화면 동작 중 오류를 숨기지 않고 보여줌
        import traceback
        detail = "".join(traceback.format_exception(exc, val, tb))[-1500:]
        messagebox.showerror(APP_TITLE, f"프로그램 오류가 났습니다. 이 내용을 알려주세요.\n\n{detail}")
    root.report_callback_exception = on_error
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
