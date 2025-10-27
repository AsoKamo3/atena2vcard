import io, re, csv
from flask import Flask, render_template, request, send_file, abort

app = Flask(__name__)

# ---------- vCardユーティリティ ----------
def esc(s: str) -> str:
    """vCardエスケープ：バックスラッシュ・改行・セミコロン・カンマ"""
    s = (s or "")
    return s.replace("\\", "\\\\").replace("\n", "\\n").replace(";", "\\;").replace(",", "\\,")

def split_multi(value):
    if not isinstance(value, str):
        return []
    return [v.strip() for v in value.split(";") if v.strip()]

def normalize_phone(num: str) -> str:
    if not isinstance(num, str) or not num.strip():
        return ""
    lead_plus = num.strip().startsWith("+") if hasattr(str, "startsWith") else num.strip().startswith("+")
    digits = "".join(re.findall(r"\d", num))
    return ("+" + digits) if (lead_plus and digits) else digits

def detect_phone_type(num: str) -> str:
    # 携帯自動判定（+81 形式にも対応）
    return "CELL" if re.match(r"^(\+81)?0?(90|80|70)", num) else "WORK"

def adr_from_row(row: dict) -> str:
    """Apple vCard 3.0 ADR 7区切り: PO Box; extended; street; locality; region; postal; country"""
    addr_type = "HOME" if (row.get("宛先", "").strip() == "自宅") else "WORK"

    # 入力列（会社/自宅を自動選択）
    postal = (row.get(f"{'自宅' if addr_type=='HOME' else '会社'}〒", "")
              or row.get("会社〒", "") or row.get("自宅〒", "")).strip()
    a1 = (row.get(f"{addr_type}住所1", "") or row.get("会社住所1", "")).strip()
    a2 = (row.get(f"{addr_type}住所2", "") or row.get("会社住所2", "")).strip()
    a3 = (row.get(f"{addr_type}住所3", "") or row.get("会社住所3", "")).strip()

    # 都道府県/市区町村の簡易抽出（例：東京都千代田区神田駿河台４－３）
    pref, city, street = "", "", ""
    if re.search(r"(都|道|府|県)", a1):
        m = re.match(r"^(.*?)(都|道|府|県)(.*)$", a1)
        if m:
            pref = m.group(1) + m.group(2)
            rest = m.group(3)
            if "区" in rest:
                p = rest.find("区")
                city = rest[: p + 1]
                street = rest[p + 1 :]
            else:
                street = rest
    else:
        street = a1

    street_full = " ".join(x for x in [street, a2] if x)
    if not any([street_full, city, pref, postal]):
        return ""

    return f"ADR;TYPE={addr_type}:;;{esc(street_full)};{esc(city)};{esc(pref)};{esc(postal)};日本"

def row_to_vcard(row: dict) -> str:
    """1件の行データ → vCard(3.0) 文字列"""
    L = ["BEGIN:VCARD", "VERSION:3.0"]

    # --- 名前（FN は「名 姓」） ---
    family = (row.get("姓") or "").strip()
    given  = (row.get("名") or "").strip()
    if family or given:
        L.append(f"N:{esc(family)};{esc(given)};;;")
        L.append(f"FN:{esc(given)} {esc(family)}")
    nick = (row.get("ニックネーム") or "").strip()
    if nick:
        L.append(f"NICKNAME:{esc(nick)}")
    fk = (row.get("姓かな") or "").strip()
    gk = (row.get("名かな") or "").strip()
    if fk: L.append(f"X-PHONETIC-LAST-NAME:{esc(fk)}")
    if gk: L.append(f"X-PHONETIC-FIRST-NAME:{esc(gk)}")

    # --- 組織 ---
    company = (row.get("会社名") or "").strip()
    dept1 = (row.get("部署名1") or "").strip()
    dept2 = (row.get("部署名2") or "").strip()
    dept = " ".join([d for d in [dept1, dept2] if d])
    if company or dept:
        L.append(f"ORG:{esc(company)};{esc(dept)}")
    org_kana = (row.get("会社名かな") or "").strip()
    if org_kana:
        L.append(f"X-PHONETIC-ORG:{esc(org_kana)}")
    title = (row.get("役職名") or "").strip()
    if title:
        L.append(f"TITLE:{esc(title)}")

    # --- 電話（複数可、携帯自動判定）---
    for label, key in [("WORK", "会社電話"), ("HOME", "自宅電話"), ("CELL", "携帯番号")]:
        for raw in split_multi(row.get(key, "") or ""):
            num = normalize_phone(raw)
            if not num:
                continue
            t = "CELL" if (label == "CELL" or detect_phone_type(num) == "CELL") else label
            L.append(f"TEL;TYPE={t},VOICE:{num}")

    # --- メール（複数可）---
    for label, key in [("WORK", "会社E-mail"), ("HOME", "自宅E-mail"), ("OTHER", "その他E-mail")]:
        for mail in split_multi(row.get(key, "") or ""):
            if mail:
                L.append(f"EMAIL;TYPE=INTERNET,{label}:{esc(mail)}")

    # --- 住所 ---
    adr = adr_from_row(row)
    if adr: L.append(adr)

    # --- 備考 NOTE（備考1〜3 を改行結合）---
    notes = [(row.get(f"備考{i}") or "").strip() for i in range(1, 4)]
    notes = [n for n in notes if n]
    if notes:
        L.append(f"NOTE:{esc('\\n'.join(notes))}")

    # --- メモ item5〜9（Apple拡張）---
    for i in range(1, 6):
        val = (row.get(f"メモ{i}") or "").strip()
        if val:
            idx = i + 4
            L.append(f"item{idx}.X-ABRELATEDNAMES:{esc(val)}")
            L.append(f"item{idx}.X-ABLabel:メモ{i}")

    L.append("END:VCARD")
    return "\n".join(L)

# ---------- CSV→行辞書の読み取り ----------
def read_csv_rows(file_bytes: bytes):
    # 代表的な文字コードを順に試す
    encodings = ["utf-8", "utf-8-sig", "cp932", "shift_jis"]
    last_err = None
    for enc in encodings:
        try:
            text = file_bytes.decode(enc)
            f = io.StringIO(text)
            reader = csv.DictReader(f)
            rows = []
            for row in reader:
                # None → "" に統一（NaN対策）
                clean = { (k or "").strip(): (v if v is not None else "") for k,v in row.items() }
                rows.append(clean)
            return rows
        except Exception as e:
            last_err = e
            continue
    raise ValueError(f"CSVの読み込みに失敗しました（文字コードの可能性）。最後のエラー: {last_err}")

# ---------- ルーティング ----------
@app.get("/")
def index():
    return render_template("index.html")

@app.post("/convert")
def convert():
    if "file" not in request.files:
        abort(400, "file is required")
    f = request.files["file"]
    data = f.read()

    try:
        rows = read_csv_rows(data)
    except Exception as e:
        abort(400, f"CSVの読み込みに失敗しました: {e}")

    vcards = [row_to_vcard(r) for r in rows]
    payload = ("\n".join(vcards)).encode("utf-8")

    return send_file(
        io.BytesIO(payload),
        mimetype="text/vcard; charset=utf-8",
        as_attachment=True,
        download_name="contacts.vcf",
    )

@app.get("/healthz")
def healthz():
    return {"ok": True}

if __name__ == "__main__":
    # ローカル実行
    app.run(host="0.0.0.0", port=5000, debug=True)
