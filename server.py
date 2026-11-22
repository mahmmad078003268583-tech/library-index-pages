#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
خادم مكتبة السبزواري العامة الداخلي
- يقرأ ملف الإكسل من نفس المجلد ويعرضه كبطاقات عبر الشبكة المحلية
- يدمج الكتب المكررة (نفس الاسم + المؤلف) ويعرض المجموع الكلي
- يولّد رمز QR لرابط الدخول — بدون إنترنت
"""
from __future__ import annotations

import argparse
import io
import json
import math
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import openpyxl
    import qrcode
except ImportError as exc:  # pragma: no cover
    sys.exit(
        "مكتبة ناقصة: %s\nثبّتها بأمر واحد:\n  pip install openpyxl qrcode pillow"
        % getattr(exc, "name", exc)
    )

# الاسم الثابت عبر mDNS — يبقى صالحاً مهما تغيّرت الشبكة (اختياري)
try:
    from zeroconf import ServiceInfo, Zeroconf
except ImportError:
    ServiceInfo = None
    Zeroconf = None

BASE_DIR = Path(__file__).resolve().parent
NCOLS = 12
SHEET_NAME = "الفهرس الشامل"
NAMELESS = "(بدون اسم الكتاب)"
UNSET_CAT = "(غير مصنف)"
PLACEHOLDERS = {"فارغ", "فارغ العطف موجود", "؟", "مستمر", "استمرارية", "نقص"}
STATUS_NAMES = {"مستمر", "استمرارية", "نقص"}
UPLOAD_PASSWORD = "0000"       # كلمة سر زر تحديث ملف الإكسل
MAX_UPLOAD = 60 * 1024 * 1024  # حد أقصى لحجم الملف المرفوع (60MB)
# تصنيفات هي ملاحظات وليس تصنيفات حقيقية
JUNK_CATS = {"فارغ", "استمرار", "مستمرة", "العطف موجود بالمخزن", "استمرارية"}

_WS = re.compile(r"\s+")
_TASHKEEL = re.compile(r"[ً-ْٰـ]")
_MAP = str.maketrans(
    {
        "أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا",
        "ى": "ي", "ة": "ه", "ؤ": "و", "ئ": "ي",
        **{chr(0x0660 + i): str(i) for i in range(10)},
        **{chr(0x06F0 + i): str(i) for i in range(10)},
    }
)


def cell_str(v) -> str:
    """تحويل قيمة خلية إلى نص نظيف."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%Y/%m/%d")
        except Exception:
            return str(v)
    return _WS.sub(" ", str(v)).strip()


def norm(v) -> str:
    """نص عادي: مسافات موحّدة بدون تشكيل."""
    s = cell_str(v)
    s = _TASHKEEL.sub("", s)
    s = s.translate(_MAP)
    return _WS.sub(" ", s).strip()


def key_norm(s) -> str:
    """تهيئة للبحث والدمج (أ/إ/آ→ا، ة→ه، ى→ي، أرقام موحّدة)."""
    return norm(s)


def merge_key(s) -> str:
    """مفتاح الدمج: بدون أي مسافات (يجمع عبدالرحيم و عبد الرحيم معاً)."""
    return key_norm(s).replace(" ", "")


_PUB_JUNK = re.compile(r"^[\s\-–—._,/\\،:;]+|[\s\-–—._,/\\،:;]+$")
# قيم شائعة في عمود دار النشر وليست أسماء دور حقيقية
_PUB_BAD = PLACEHOLDERS | {"/", "\\", ",", "،", "-", "؟", "?", "…"}


def pub_clean(v) -> str:
    """اسم دار النشر نظيفاً للعرض (بلا شارات في الطرفين وبلا مسافات مكررة)."""
    s = _PUB_JUNK.sub("", cell_str(v)).replace("ـ", "")
    return _WS.sub(" ", s).strip()


def pub_key(v) -> str:
    """مفتاح دار النشر: موحّد وبلا مسافات وبلا شارات زائدة (يجمع «...-» و«...»)."""
    return merge_key(pub_clean(v))


def year_num(v) -> int | None:
    """رقم السنة من خلية (ميلادي أو هجري): أكبر رقم بين 1000 و2200 وإلا فلا."""
    nums = re.findall(r"\d+", norm(v))
    if not nums:
        return None
    y = max(int(n) for n in nums)
    return y if 1000 <= y <= 2200 else None


# ---------------------------------------------------------------------------
# حالة البيانات (تُعاد قراءتها تلقائياً عند تغيير ملف الإكسل)
# ---------------------------------------------------------------------------
STATE = {
    "items": [],      # [{t: العنوان, n: العدد, f: [[تسمية, قيمة], ...]}]
    "search": [],     # نص البحث المهيّأ لكل كتاب (بترتيب الأ/items)
    "author": [],     # مفتاح اسم المؤلف لكل كتاب (للبحث المتقدم)
    "cat_of": [],     # التصنيف المعروض لكل كتاب
    "pub_of": [],     # مفتاح دار النشر لكل كتاب (للفرز)
    "greg": [],       # أحدث سنة ميلادية لكل كتاب (للفرز)
    "hijri": [],      # أحدث سنة هجرية لكل كتاب (للفرز)
    "publishers": [], # [(اسم الدار, عدد الكتب, المفتاح)]
    "suggest": [],    # مرشّحات الإكمال التلقائي: (أساس, الشكل, المفتاح)
    "headers": [],
    "categories": [],  # [(الاسم, العدد)]
    "stats": {},
    "path": None,
    "mtime": 0.0,
}
_LOCK = threading.RLock()
_LOAD_LOCK = threading.Lock()
_RETRY = {"until": 0.0}
SITE = {"port": None, "url": None, "lan_url": None}
# يمكن فرض عنوان الشبكة الداخلية (راوتر الطلبة) عبر المتغير أو الخيار --lan-ip
FORCE_INTERNAL_IP = os.environ.get("LIBRARY_INTERNAL_IP", "").strip()


def find_excel() -> Path | None:
    files = [f for f in BASE_DIR.glob("*.xlsx") if not f.name.startswith("~$")]
    if not files:
        return None
    for f in files:
        if "فهرس" in f.name:
            return f
    return max(files, key=lambda p: p.stat().st_mtime)


def load_data(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.worksheets[0]
        rows_iter = ws.iter_rows(values_only=True)
        header_row = next(rows_iter, None)
        headers = [str(h).strip() if h is not None else "" for h in list(header_row or ())[:NCOLS]]
        while len(headers) < NCOLS:
            headers.append("عمود %d" % (len(headers) + 1))

        # أعمدة الفرز: دار النشر وسنتا الطبع — تُكتشف من رؤوس الأعمدة
        def _col(*needles: str):
            for i, h in enumerate(headers):
                if all(n in h for n in needles):
                    return i
            return None

        col_pub = _col("دار النشر")
        col_g = _col("ميلادي")
        col_h = _col("هـجري")
        if col_h is None:
            col_h = _col("هجري")

        groups: dict = {}
        order: list = []
        kept_rows = 0

        for raw in rows_iter:
            vals = [cell_str(raw[i]) if i < len(raw) else "" for i in range(NCOLS)]
            if not any(vals):
                continue
            if vals[0] in STATUS_NAMES:
                continue
            filled = [v for v in vals if v]
            if all(v in PLACEHOLDERS for v in filled):
                continue
            name = vals[0]
            if name.startswith("فهرس مكتبة"):
                continue  # صف عنوان داخلي وليس كتاباً
            if name in ("", "فارغ", "فارغ العطف موجود"):
                # صف بلا اسم كتاب: نحتفظ به فقط إذا كان غنياً بالبيانات
                rich = sum(1 for v in vals[1:] if v and v not in PLACEHOLDERS)
                if rich < 3:
                    continue
                title = NAMELESS
            else:
                title = name

            kept_rows += 1
            k = merge_key(title) + "\x00" + merge_key(vals[1])
            it = groups.get(k)
            if it is None:
                it = {"t": title, "n": 0, "v": list(vals), "rows": []}
                groups[k] = it
                order.append(k)
            it["n"] += 1
            it["rows"].append(list(vals))
            # إثراء الحقول الفارغة من بقية الصفوف المكررة
            v = it["v"]
            for i in range(NCOLS):
                if not v[i] and vals[i]:
                    v[i] = vals[i]
    finally:
        wb.close()

    items, search, cat_of = [], [], []
    author_of = []
    pub_of, greg_of, hijri_of = [], [], []
    pub_labels: dict = {}   # مفتاح الدار -> الشكل الأنصح للعرض
    pub_counts: dict = {}   # مفتاح الدار -> عدد الكتب
    cat_labels: dict = {}   # المفتاح المهيّأ -> العدّاد
    cat_display: dict = {}  # المفتاح المهيّأ -> الشكل الأشيع للعرض
    suggest: list = []      # مرشّحات الإكمال: (0=كتاب, 1=مؤلف, الشكل, المفتاح)
    seen_titles: set = set()
    seen_authors: set = set()
    dup_groups = 0
    for k in order:
        it = groups[k]
        vals = it["v"]
        fields = [(headers[i], vals[i]) for i in range(1, NCOLS) if vals[i]]
        items.append(
            {"t": it["t"], "a": vals[1], "n": it["n"], "f": fields, "r": it["rows"]}
        )
        search.append(merge_key(" ".join(vals)))
        author_of.append(merge_key(vals[1]))
        # مرشّحات الإكمال التلقائي (عناوين + مؤلفون فريدون)
        tmk = merge_key(it["t"])
        if tmk and tmk not in seen_titles:
            seen_titles.add(tmk)
            suggest.append((0, it["t"], tmk))
        amk = merge_key(vals[1])
        if vals[1] and amk not in seen_authors:
            seen_authors.add(amk)
            suggest.append((1, vals[1], amk))
        # تهيئة التصنيف: دمج اختلاف الهمزات وحذف # وتصفية الملاحظات
        raw_cat = vals[7]
        cat_key = key_norm(raw_cat).lstrip("#").strip()
        if not cat_key or cat_key in {key_norm(j) for j in JUNK_CATS}:
            cat_key = UNSET_CAT
            label = UNSET_CAT
        else:
            label = raw_cat or cat_key
        cat_of.append(cat_key)
        cat_labels[cat_key] = cat_labels.get(cat_key, 0) + 1
        if cat_key not in cat_display:
            cat_display[cat_key] = label
        # دار النشر وسنتا الطبع (للفرز): القيم من الصف المُثرى نفسه المعروض في البطاقة
        # كي يوافق ترتيب الفرز تماماً ما يراه المستخدم على الشاشة
        raw_pub = vals[col_pub] if col_pub is not None else ""
        label_pub = pub_clean(raw_pub)
        # القيم الشائعة (؟، /، فارغ…) ليست أسماء دور → لا تدخل القائمة
        pk = "" if (not label_pub or label_pub in _PUB_BAD) else pub_key(raw_pub)
        pub_of.append(pk)
        if pk:
            pub_counts[pk] = pub_counts.get(pk, 0) + 1
            if pk not in pub_labels:
                pub_labels[pk] = label_pub
        gy = year_num(vals[col_g]) if col_g is not None else None
        hy = year_num(vals[col_h]) if col_h is not None else None
        greg_of.append(gy)
        hijri_of.append(hy)
        if it["n"] > 1:
            dup_groups += 1

    try:
        mtime = path.stat().st_mtime
    except OSError:
        mtime = 0.0
    stats = {
        "rows": kept_rows,
        "books": len(items),
        "merged": kept_rows - len(items),
        "dup_groups": dup_groups,
        "excel": path.name,
        "sheet": ws.title,
        "updated": time.strftime("%Y-%m-%d %H:%M", time.localtime(mtime)),
    }
    return {
        "items": items,
        "search": search,
        "author": author_of,
        "cat_of": cat_of,
        "pub_of": pub_of,
        "greg": greg_of,
        "hijri": hijri_of,
        "publishers": [
            (pub_labels[k], pub_counts[k], k)
            for k in sorted(pub_counts, key=lambda k: pub_labels[k])
        ],
        "suggest": suggest,
        "headers": headers,
        "categories": [
            (cat_display[key], cat_labels[key], key) for key in sorted(cat_labels)
        ],
        "stats": stats,
    }


def apply_data(data: dict, path: Path) -> None:
    with _LOCK:
        STATE.update(data)
        try:
            STATE["mtime"] = path.stat().st_mtime
        except OSError:
            STATE["mtime"] = 0.0


def ensure_fresh() -> None:
    """إعادة قراءة الإكسل تلقائياً إذا تغيّر الملف."""
    path = STATE["path"]
    if path is None:
        return
    try:
        m = path.stat().st_mtime
    except OSError:
        return
    if m == STATE["mtime"]:
        return
    if time.time() < _RETRY["until"]:
        return
    if not _LOAD_LOCK.acquire(blocking=False):
        return  # عملية تحميل جارية الآن
    try:
        try:
            data = load_data(path)
        except Exception as e:
            _RETRY["until"] = time.time() + 10
            print("[تحذير] تعذّر قراءة الإكسل (%s) — تبقى البيانات القديمة" % e, flush=True)
            return
        apply_data(data, path)
        print(
            "[تحديث] تم تحميل %d كتاباً من الإكسل (%s)"
            % (data["stats"]["books"], data["stats"]["updated"]),
            flush=True,
        )
    finally:
        _LOAD_LOCK.release()


def initial_load(path: Path) -> None:
    for attempt in range(1, 7):
        try:
            data = load_data(path)
        except Exception as e:
            print("محاولة %d: تعذّر قراءة الملف (%s)" % (attempt, e), flush=True)
            time.sleep(2)
            continue
        apply_data(data, path)
        s = data["stats"]
        print(
            "تمت القراءة: %d كتاباً فريداً من %d صفاً (دُمج %d صف مكرر في %d كتاباً)"
            % (s["books"], s["rows"], s["merged"], s["dup_groups"]),
            flush=True,
        )
        return
    print("تنبيه: تعذّر قراءة الملف — ستُعاد المحاولة تلقائياً عند أول طلب.", flush=True)


# ---------------------------------------------------------------------------
# الشبكة ورمز QR
# ---------------------------------------------------------------------------
def lan_ip() -> str:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 53))  # لا يُرسل أي حزم (اتصال UDP تخيلي)
        ip = s.getsockname()[0]
    except OSError:
        ip = ""
    finally:
        s.close()
    if not ip or ip.startswith(("127.", "169.254.")):
        try:
            infos = socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET)
            for _, _, _, _, sa in infos:
                if not sa[0].startswith(("127.", "169.254.")):
                    ip = sa[0]
                    break
        except OSError:
            pass
    return ip or "localhost"


def _private_v4(ip: str) -> bool:
    """عنوان IPv4 خاص بالشبكات الداخلية (ليس Tailscale ولا APIPA)."""
    try:
        a, b, _c, _d = (int(x) for x in ip.split("."))
    except (ValueError, AttributeError):
        return False
    return a == 10 or (a == 172 and 16 <= b <= 31) or (a == 192 and b == 168)


def private_lan_ips() -> list[str]:
    """كل عناوين الشبكة الداخلية للجهاز — قد يكون أكثر من كرت/راوتر."""
    found: set[str] = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if _private_v4(ip):
                found.add(ip)
    except OSError:
        pass
    return sorted(found)


def offline_lan_url() -> str:
    """رابط الطلبة: يفضّل كرت الشبكة غير كرت الإنترنت إن وُجد."""
    port = SITE["port"] or 80
    suffix = "" if port == 80 else ":%d" % port
    if FORCE_INTERNAL_IP:
        return "http://%s%s/" % (FORCE_INTERNAL_IP, suffix)
    current = current_lan_ip()
    for ip in private_lan_ips():
        if ip != current:
            return "http://%s%s/" % (ip, suffix)
    return "http://%s%s/" % (current, suffix)


_IP_CACHE = {"at": 0.0, "ip": ""}


def current_lan_ip() -> str:
    """عنوان الشبكة الحالي مع تخزين قصير (يُعاد الحساب كل 15 ثانية)."""
    now = time.time()
    if now - _IP_CACHE["at"] < 15 and _IP_CACHE["ip"]:
        return _IP_CACHE["ip"]
    ip = lan_ip()
    _IP_CACHE["at"] = now
    _IP_CACHE["ip"] = ip
    return ip


def current_lan_url() -> str:
    """رابط الأجهزة — يتحدّث تلقائياً إذا تغيّرت شبكة الجهاز."""
    port = SITE["port"] or 80
    suffix = "" if port == 80 else ":%d" % port
    ip = current_lan_ip()
    url = "http://%s%s/" % (ip, suffix)
    mdns_sync(ip)
    if url == SITE["lan_url"]:
        return url
    old = SITE["lan_url"]
    SITE["lan_url"] = url
    if old:
        print("[شبكة] تغيّر عنوان الجهاز:  %s  ←→  %s" % (old, url), flush=True)
    try:
        lines = []
        if external_url():
            lines.append(
                "الرابط العام الثابت (اطبعه مرة واحدة — يعمل من أي مكان): "
                + external_url()
            )
        if stable_url():
            lines.append("الرابط الثابت الداخلي: " + STABLE_URL)
        lines.append("الرابط الحالي: " + url)
        lan = offline_lan_url()
        (BASE_DIR / "QR-المكتبة.png").write_bytes(qr_png(best_url() or url))
        (BASE_DIR / "QR-المكتبة-داخلي.png").write_bytes(qr_png(lan))
        lines.append("رابط الشبكة الداخلية (بلا إنترنت): " + lan)
        (BASE_DIR / "رابط-المكتبة.txt").write_text(
            "\n".join(lines) + "\n", encoding="utf-8"
        )
        write_desktop_files()
    except Exception as e:
        print("[تحذير] تعذّر تحديث ملفات الرابط (%s)" % e, flush=True)
    return url


# ---------------------------------------------------------------------------
# الاسم الثابت library.local (mDNS) — يعمل على أي شبكة بلا إنترنت
# ---------------------------------------------------------------------------
STABLE_HOST = "library.local."
STABLE_URL = "http://library.local/"
_MDNS = {"zc": None, "info": None, "ip": None}


def mdns_sync(ip: str) -> bool:
    """يدوّن الاسم library.local ويجدّده تلقائياً عند تغيّر عنوان الشبكة."""
    if Zeroconf is None or not ip or ip == "localhost":
        return False
    if _MDNS["ip"] == ip and _MDNS["info"] is not None:
        return True
    try:
        if _MDNS["info"] is not None:
            try:
                _MDNS["zc"].unregister_service(_MDNS["info"])
            except Exception:
                pass
            _MDNS["info"] = None
        info = ServiceInfo(
            "_http._tcp.local.",
            "Library Index._http._tcp.local.",
            addresses=[socket.inet_aton(ip)],
            port=SITE["port"] or 80,
            properties={"path": "/"},
            server=STABLE_HOST,
        )
        if _MDNS["zc"] is None:
            _MDNS["zc"] = Zeroconf()
        _MDNS["zc"].register_service(info)
        _MDNS["info"] = info
        _MDNS["ip"] = ip
        print("[اسم ثابت] %s  ←→  %s" % (STABLE_URL, ip), flush=True)
        return True
    except Exception as e:
        print("[تنبيه] تعذّر تشغيل الاسم الثابت (%s)" % e, flush=True)
        return False


def stable_url() -> str:
    return STABLE_URL if _MDNS["info"] is not None else ""


def mdns_stop() -> None:
    try:
        if _MDNS["zc"] is not None:
            if _MDNS["info"] is not None:
                _MDNS["zc"].unregister_service(_MDNS["info"])
            _MDNS["zc"].close()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# الرابط العام الثابت عبر النفق (Tailscale Funnel) — يعمل من أي شبكة
# ---------------------------------------------------------------------------
EXTERNAL_URL_FILE = BASE_DIR / "الرابط-الخارجي.txt"
TUNNEL_CLI = Path(r"C:\Program Files\Tailscale\tailscale.exe")


def external_url() -> str:
    """الرابط العام الثابت — يُقرأ من ملف يُنشأ عند تفعيل النفق."""
    try:
        if EXTERNAL_URL_FILE.is_file():
            u = EXTERNAL_URL_FILE.read_text(encoding="utf-8").strip()
            if u.startswith("http"):
                return u
    except Exception:
        pass
    return ""


def best_url() -> str:
    """أفضل رابط للطباعة: العام (يعمل من أي مكان) ثم الداخلي الثابت ثم الحالي."""
    return external_url() or stable_url() or SITE.get("lan_url") or ""


def funnel_sync(port: int) -> None:
    """يطابق النفق العام مع منفذ الخادم الحالي (best-effort)."""
    if not external_url() or not TUNNEL_CLI.is_file():
        return
    try:
        subprocess.run(
            [str(TUNNEL_CLI), "funnel", "--bg", str(port)],
            creationflags=subprocess.CREATE_NO_WINDOW,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=20,
        )
        print("[نفق عام] %s  ←→  منفذ %d" % (external_url(), port), flush=True)
    except Exception as e:
        print("[تنبيه] تعذّر مزامنة النفق العام (%s)" % e, flush=True)


# ---------------------------------------------------------------------------
# ملفات سطح المكتب: رمز QR + الرابط + اختصار يفتح الموقع
# ---------------------------------------------------------------------------
def desktop_dir() -> Path | None:
    home = Path.home()
    cands = [home / "Desktop"]
    onedrive = home / "OneDrive"
    if onedrive.is_dir():
        cands.extend(sorted(onedrive.glob("Desktop")))
    for c in cands:
        if c.is_dir():
            return c
    return None


def write_desktop_files() -> None:
    """ينسخ رمز QR الثابت ورابط الدخول واختصاراً إلى سطح المكتب."""
    d = desktop_dir()
    if d is None:
        return
    link = best_url()
    if not link:
        return
    try:
        lan = offline_lan_url()
        (d / "QR-المكتبة.png").write_bytes(qr_png(link))
        (d / "QR-المكتبة-داخلي.png").write_bytes(qr_png(lan))
        lines = []
        if external_url():
            lines.append(
                "الرابط العام الثابت (اطبعه مرة واحدة — يعمل من البيت والمع): "
                + external_url()
            )
        if stable_url():
            lines.append("الرابط الثابت الداخلي: " + STABLE_URL)
        if SITE.get("lan_url"):
            lines.append("الرابط الحالي: " + SITE["lan_url"])
        if lan:
            lines.append("رابط الشبكة الداخلية (بلا إنترنت): " + lan)
        if not lines:
            lines = ["الرابط: " + link]
        (d / "رابط-المكتبة.txt").write_text(
            "\n".join(lines) + "\n", encoding="utf-8"
        )
        (d / "مكتبة السبزواري العامة.url").write_text(
            "[InternetShortcut]\r\nURL=%s\r\n" % link, encoding="utf-8"
        )
        print(
            "[سطح المكتب] أُنشئت: QR-المكتبة.png + QR-المكتبة-داخلي.png + رابط-المكتبة.txt + مكتبة السبزواري العامة.url",
            flush=True,
        )
    except Exception as e:
        print("[تنبيه] تعذّر تحديث ملفات سطح المكتب (%s)" % e, flush=True)


def qr_png(url: str) -> bytes:
    qr = qrcode.QRCode(
        error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=2
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------
def _first(params: dict, key: str, default: str = "") -> str:
    v = params.get(key)
    return (v[0] if v else default).strip()


def _int(v: str, default: int) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


# ملفات PWA — التطبيق القابل للتثبيت على شاشة الجوال (أيقونات وmanifest)
_STATIC = {
    "/manifest.json": ("manifest.json", "application/manifest+json; charset=utf-8"),
    "/sw.js": ("sw.js", "application/javascript; charset=utf-8"),
    "/icon-192.png": ("icon-192.png", "image/png"),
    "/icon-512.png": ("icon-512.png", "image/png"),
    "/apple-touch-icon.png": ("apple-touch-icon.png", "image/png"),
}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print("[%s] %s" % (time.strftime("%H:%M:%S"), fmt % args), flush=True)

    def _send(self, code: int, body: bytes, ctype: str, cache: str = "no-store") -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", cache)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj: dict, code: int = 200) -> None:
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self._send(code, body, "application/json; charset=utf-8")

    def do_GET(self):  # noqa: N802
        u = urlparse(self.path)
        try:
            if u.path == "/":
                page = BASE_DIR / "index.html"
                data = (
                    page.read_bytes()
                    if page.exists()
                    else "<h1>index.html غير موجود</h1>".encode("utf-8")
                )
                self._send(200, data, "text/html; charset=utf-8", "no-cache")
            elif u.path == "/api/meta":
                ensure_fresh()
                with _LOCK:
                    stats = dict(STATE["stats"])
                    cats = [list(c) for c in STATE["categories"]]
                self._json(
                    {
                        "books": stats.get("books", 0),
                        "rows": stats.get("rows", 0),
                        "merged": stats.get("merged", 0),
                        "dup_groups": stats.get("dup_groups", 0),
                        "excel": stats.get("excel", ""),
                        "sheet": stats.get("sheet", ""),
                        "updated": stats.get("updated", ""),
                        "categories": cats,
                        "url": current_lan_url(),
                        "stable": stable_url(),
                        "external": external_url(),
                        "lan_internal": offline_lan_url(),
                        "local_url": SITE["url"],
                    }
                )
            elif u.path == "/api/books":
                ensure_fresh()
                params = parse_qs(u.query)
                self._json(self._books(params))
            elif u.path == "/api/suggest":
                ensure_fresh()
                params = parse_qs(u.query)
                self._json(self._suggest(params))
            elif u.path == "/api/parts":
                ensure_fresh()
                params = parse_qs(u.query)
                self._json(self._parts(params))
            elif u.path == "/qr.png":
                self._send(200, qr_png(best_url() or current_lan_url()), "image/png")
            elif u.path == "/qr-lan.png":
                self._send(200, qr_png(offline_lan_url()), "image/png")
            elif u.path in _STATIC:
                fname, ctype = _STATIC[u.path]
                fpath = BASE_DIR / fname
                if fpath.is_file():
                    self._send(200, fpath.read_bytes(), ctype, "no-cache")
                else:
                    self._send(404, b"404", "text/plain; charset=utf-8")
            else:
                self._send(404, b"404", "text/plain; charset=utf-8")
        except (BrokenPipeError, ConnectionResetError):
            pass

    def do_POST(self):  # noqa: N802
        u = urlparse(self.path)
        try:
            if u.path == "/api/upload":
                self._upload(parse_qs(u.query))
            else:
                self._send(404, b"404", "text/plain; charset=utf-8")
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _books(self, params: dict) -> dict:
        raw_q = _first(params, "q")
        cat = _first(params, "cat")
        sort = _first(params, "sort")       # "" | "greg" | "hijri"
        raw_pub = _first(params, "pub")     # اسم دار النشر (كامل أو جزء منه)
        raw_author = _first(params, "au")   # اسم المؤلف (بحث متقدم)
        y_from = _int(_first(params, "yf"), 0) or None
        y_to = _int(_first(params, "yt"), 0) or None
        y_kind = _first(params, "yk") or "greg"   # greg | hijri
        page = max(1, _int(_first(params, "page", "1"), 1))
        per = min(100, max(10, _int(_first(params, "per", "60"), 60)))
        # كلمات البحث بدون مسافات كي تطابق "عبدالرحيم" و "عبد الرحيم" معاً
        tokens = [key_norm(t) for t in raw_q.split() if t.strip()]
        joined = merge_key(raw_q)
        author_tokens = [key_norm(t) for t in raw_author.split() if t.strip()]

        with _LOCK:
            items = STATE["items"]
            search = STATE["search"]
            authors = STATE["author"]
            cat_of = STATE["cat_of"]
            pub_of = STATE["pub_of"]
            greg = STATE["greg"]
            hijri = STATE["hijri"]
            pub_keys = [p[2] for p in STATE["publishers"]]

        # دار النشر: تطابق تام بالمفتاح، وإلا تطابق جزئي بكل الدور المطابقة
        pub_set = None
        if raw_pub:
            pk = pub_key(raw_pub)
            if pk:
                pub_set = {pk} | {k for k in pub_keys if pk in k}

        year_arr = hijri if y_kind == "hijri" else greg

        sel = []
        for i, it in enumerate(items):
            # البحث النصي يشمل كل الكتب مهما كان التصنيف المحدد
            if cat and not tokens and cat_of[i] != cat:
                continue
            if pub_set is not None and pub_of[i] not in pub_set:
                continue
            if author_tokens and not all(t in authors[i] for t in author_tokens):
                continue
            if y_from is not None or y_to is not None:
                y = year_arr[i]
                if y is None:
                    continue  # الكتب بلا سنة لا تدخل نطاق السنة
                if y_from is not None and y < y_from:
                    continue
                if y_to is not None and y > y_to:
                    continue
            if tokens and not all(t in search[i] for t in tokens):
                continue
            sel.append(i)

        # الفرز الصريح يغلب ترتيب البحث، وفرز دار النشر يبدأ بالميلادي الأحدث
        key_arr = None
        if sort == "hijri":
            key_arr = hijri
        elif sort == "greg" or (pub_set is not None and not sort):
            key_arr = greg

        if key_arr is not None:
            # من الأحدث إلى الأقدم، والأكتب بلا سنة في النهاية (مع ثبات ترتيب الفهرس)
            sel.sort(key=lambda i: (key_arr[i] is None, -(key_arr[i] or 0)))
        elif tokens:
            # ترتيب: تطابق تام مع الاسم، ثم الاسم يتضمنه، ثم المؤلف، ثم الباقي
            titles = [merge_key(it["t"]) for it in items]
            authors = [merge_key(it.get("a", "")) for it in items]

            def rank(i: int) -> int:
                t = titles[i]
                if joined and t == joined:
                    return 0
                if joined and joined in t:
                    return 1
                if all(x in t for x in tokens):
                    return 2
                if all(x in authors[i] for x in tokens):
                    return 3
                return 4
            sel.sort(key=rank)

        total = len(sel)
        pages = max(1, math.ceil(total / per))
        start = (page - 1) * per
        # نُرجع المؤشر الأصلي (i) لفتح نافذة الأجزاء، ونبقي الصفوف الكبيرة خارج القائمة
        page_items = [
            {
                "t": items[i]["t"],
                "a": items[i]["a"],
                "n": items[i]["n"],
                "f": items[i]["f"],
                "i": i,
            }
            for i in sel[start : start + per]
        ]
        return {
            "total": total,
            "page": page,
            "pages": pages,
            "per": per,
            "items": page_items,
        }

    def _parts(self, params: dict) -> dict:
        """أجزاء كتاب واحد (الصفوف المدمجة فيه) لنافذة مستقلة."""
        idx = _int(_first(params, "i"), -1)
        with _LOCK:
            items = STATE["items"]
            headers = STATE["headers"]
        if not (0 <= idx < len(items)):
            self._json({"ok": False, "error": "هذا الكتاب غير موجود"}, 404)
            return
        it = items[idx]
        self._json(
            {
                "ok": True,
                "t": it["t"],
                "a": it.get("a", ""),
                "n": it["n"],
                "headers": headers,
                "parts": it.get("r", []),
            }
        )

    def _suggest(self, params: dict) -> dict:
        """إكمال تلقائي شبيه بجوجل: عناوين ومؤلفون يبدؤون بما كتبه الطالب."""
        raw_q = _first(params, "q")
        limit = min(10, max(3, _int(_first(params, "n", "8"), 8)))
        joined = merge_key(raw_q)
        if not joined:
            return {"suggestions": []}
        tokens = [key_norm(t) for t in raw_q.split() if t.strip()]

        with _LOCK:
            src = STATE["suggest"]

        scored = []
        for base, disp, mk in src:
            if mk.startswith(joined):
                sc = base            # بداية النص: الأقوى (كتاب ثم مؤلف)
            elif joined in mk:
                sc = 2 + base        # وسط النص
            elif all(t in mk for t in tokens):
                sc = 4 + base        # كل الكلمات موجودة
            else:
                continue
            scored.append((sc, len(mk), disp))
        scored.sort()

        out, seen = [], set()
        for _, _, disp in scored:
            if disp not in seen:
                seen.add(disp)
                out.append(disp)
            if len(out) >= limit:
                break
        return {"suggestions": out}

    def _upload(self, params: dict) -> None:
        """استبدال ملف الإكسل برفع ملف جديد (محمي بكلمة سر)."""
        if _first(params, "pw") != UPLOAD_PASSWORD:
            self._json({"ok": False, "error": "كلمة السر غير صحيحة"}, 401)
            return

        name = _first(params, "name") or "الملف الجديد.xlsx"
        if not name.lower().endswith(".xlsx"):
            self._json({"ok": False, "error": "يجب أن يكون الملف بصيغة xlsx"}, 400)
            return

        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        if length <= 0:
            self._json({"ok": False, "error": "لم يصل أي ملف"}, 400)
            return
        if length > MAX_UPLOAD:
            self._json({"ok": False, "error": "الملف كبير جداً (الحد 60MB)"}, 400)
            return

        body = self.rfile.read(length)
        # ملف مؤقت يبدأ بـ ~$ حتى لا يلتقطه find_excel أثناء الرفع
        tmp = BASE_DIR / ("~$upload-%d.xlsx" % int(time.time() * 1000))
        try:
            tmp.write_bytes(body)
        except OSError as e:
            self._json({"ok": False, "error": "تعذّر حفظ الملف: %s" % e}, 500)
            return

        if not _LOAD_LOCK.acquire(blocking=False):
            tmp.unlink(missing_ok=True)
            self._json({"ok": False, "error": "جارٍ تحديث الآن — أعد المحاولة بعد لحظات"}, 503)
            return
        try:
            path = STATE.get("path") or find_excel()
            if path is None:
                self._json({"ok": False, "error": "لم يُعثر على ملف الإكسل الأصلي"}, 500)
                return

            # 1) التحقق من صلاحية الملف قبل أي استبدال
            try:
                data = load_data(tmp)
            except Exception as e:
                self._json(
                    {"ok": False, "error": "الملف غير صالح أو تالف: %s" % e}, 400
                )
                return

            # 2) نسخة احتياطية من الملف الحالي ثم استبداله
            try:
                bak_dir = BASE_DIR / "نسخ احتياطية"
                bak_dir.mkdir(exist_ok=True)
                if path.exists():
                    stamp = time.strftime("%Y%m%d-%H%M%S")
                    shutil.copy2(path, bak_dir / ("%s-%s" % (stamp, path.name)))
                shutil.move(str(tmp), str(path))
            except OSError as e:
                self._json({"ok": False, "error": "تعذّر استبدال الملف: %s" % e}, 500)
                return

            # 3) تفعيل البيانات الجديدة فوراً (بالاسم والتاريخ الحقيقيين للملف)
            s = data["stats"]
            s["excel"] = path.name
            try:
                s["updated"] = time.strftime(
                    "%Y-%m-%d %H:%M", time.localtime(path.stat().st_mtime)
                )
            except OSError:
                pass
            apply_data(data, path)
            print(
                "[تحديث] تم استبدال ملف الإكسل (%s) — %d كتاباً"
                % (s.get("excel", path.name), s.get("books", 0)),
                flush=True,
            )
            self._json(
                {
                    "ok": True,
                    "books": s.get("books", 0),
                    "rows": s.get("rows", 0),
                    "merged": s.get("merged", 0),
                    "updated": s.get("updated", ""),
                    "excel": s.get("excel", path.name),
                }
            )
        finally:
            tmp.unlink(missing_ok=True)
            _LOAD_LOCK.release()


# ---------------------------------------------------------------------------
# التشغيل
# ---------------------------------------------------------------------------
def print_banner() -> None:
    with _LOCK:
        s = dict(STATE["stats"])
    line = "═" * 60
    print()
    print(line)
    print("  مكتبة السبزواري العامة — الخادم الداخلي يعمل الآن")
    print(line)
    print("  على هذا الجهاز : %s" % SITE["url"])
    print("  للأجهزة (هاتف) : %s" % SITE["lan_url"])
    print("  للشبكة الداخلية: %s  (بلا إنترنت)" % offline_lan_url())
    if stable_url():
        print("  الرابط الثابت  : %s  (يبقى مهما تغيّرت الشبكة)" % STABLE_URL)
    if external_url():
        print("  الرابط العام   : %s  (يعمل من أي مكان)" % external_url())
    print("  امسح رمز QR أو اكتب الرابط في المتصفح")
    if s:
        print(
            "  الكتب: %d | الصفوف: %d | المكرر المدمج: %d"
            % (s.get("books", 0), s.get("rows", 0), s.get("merged", 0))
        )
    print(line)
    print("  لإيقاف الخادم: أغلق النافذة أو اضغط Ctrl+C")
    print(line, "\n", flush=True)


def main() -> None:
    global FORCE_INTERNAL_IP
    ap = argparse.ArgumentParser(description="خادم مكتبة السبزواري العامة الداخلي")
    ap.add_argument("--port", type=int, default=None, help="رقم المنفذ (افتراضي: 80 ثم 8000)")
    ap.add_argument("--no-browser", action="store_true", help="لا تفتح المتصفح تلقائياً")
    ap.add_argument(
        "--lan-ip",
        default=FORCE_INTERNAL_IP,
        help="عنوان الحاسبة على راوتر الطلبة (داخلي بلا إنترنت)",
    )
    args = ap.parse_args()
    FORCE_INTERNAL_IP = args.lan_ip.strip()

    ports = [args.port] if args.port else [80, 8000]

    # أولاً: هل كان الخادم يعمل مسبقاً؟ (يفتح المتصفح وينتهي بدون تشغيل ثانٍ)
    import urllib.request

    for p in ports:
        try:
            with urllib.request.urlopen(
                "http://127.0.0.1:%d/api/meta" % p, timeout=1
            ) as resp:
                if resp.status == 200:
                    print("الخادم يعمل مسبقاً على المنفذ %d — يتم فتح المتصفح." % p)
                    if not args.no_browser:
                        webbrowser.open(
                            "http://localhost%s/" % ("" if p == 80 else ":%d" % p)
                        )
                    return
        except Exception:
            continue

    excel = find_excel()
    if excel is None:
        sys.exit("لم يُعثر على ملف Excel (*.xlsx) في: %s" % BASE_DIR)
    print("جارٍ قراءة: %s ..." % excel.name, flush=True)
    STATE["path"] = excel
    initial_load(excel)

    httpd = None
    for p in ports:
        try:
            httpd = ThreadingHTTPServer(("0.0.0.0", p), Handler)
            SITE["port"] = p
            break
        except OSError as e:
            print("المنفذ %d غير متاح (%s)" % (p, e), flush=True)
    if httpd is None:
        sys.exit("تعذّر فتح أي منفذ.")

    port = SITE["port"]
    suffix = "" if port == 80 else ":%d" % port
    SITE["url"] = "http://localhost%s/" % suffix
    current_lan_url()  # يحسب العنوان ويحفظ QR والملف النصي، وينعك تغيّره تلقائياً
    write_desktop_files()  # رمز QR والرابط واختصار الموقع على سطح المكتب
    funnel_sync(port)  # يطابق النفق العام (إن وُجد) مع منفذ الخادم

    print_banner()
    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(SITE["url"])).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nتم إيقاف الخادم.", flush=True)
    finally:
        httpd.server_close()
        mdns_stop()


if __name__ == "__main__":
    main()
