"""
TSDR Serial Number Scraper
--------------------------
Excel/CSV file se serial numbers uthata hai, phir har serial ke liye
https://tsdr.uspto.gov/statusview/sn<SERIAL> pe request karta hai
(Webshare proxy ke through), response HTML parse karta hai, aur
result ko do alag CSV files mein save karta hai: valid leads aur missing.

Requirements:
    pip install requests beautifulsoup4 openpyxl --break-system-packages

Usage:
    $env:TSDR_PROXY_USERNAME = "your-webshare-username"
    $env:TSDR_PROXY_PASSWORD = "your-webshare-password"
    python tsdr_scraper.py serials.xlsx
    python tsdr_scraper.py serials.csv

The ten Webshare IP:port proxies shown in the dashboard are used as a
randomized pool. Credentials are read from environment variables.
"""

import csv
import os
import re
import sys
import time
import random
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup

# ===================== CONFIG =====================

# Webshare proxy pool. Har request par random IP select hota hai.
PROXY_LIST = [
    ("31.59.20.176", "6754"),
    ("45.38.107.97", "6014"),
    ("64.137.96.74", "6641"),
    ("198.23.243.226", "6361"),
    ("38.154.185.97", "6370"),
    ("84.247.60.125", "6095"),
    ("142.111.67.146", "5611"),
    ("191.96.254.138", "6185"),
    ("31.58.9.4", "6077"),
    ("198.46.161.42", "5092"),
]

# Credentials environment variables mein rakho; source code ya git history
# mein password save mat karo.
PROXY_USERNAME = os.getenv("TSDR_PROXY_USERNAME", "akzqqhnr")
PROXY_PASSWORD = os.getenv("TSDR_PROXY_PASSWORD", "1k5hw8ogi3os")

USE_PROXY = True           # False karo agar proxy use nahi karna
CONCURRENT_WORKERS = 6     # ek sath kitni requests parallel chalein (5-10 rakho)
REQUEST_DELAY = (0.2, 0.6)  # har thread ke apne request ke beech chhota delay
MAX_RETRIES = 3
TIMEOUT = 20

# Attorney filter: False rakho to sirf "Attorney of Record - None" wali
# leads valid mein aayengi (jyada tar isi type pe kaam hota hai).
# True karoge to sirf wo leads aayengi jinka koi attorney of record hai.
WANT_ATTORNEY = False
#   "deadAbandoned"  -> DEAD / ABANDONED
#   "deadCancelled"  -> DEAD / CANCELLED (registration cancelled/invalidated)
#   "livePending"    -> LIVE / PENDING
#   "liveRegister"   -> LIVE / REGISTERED
#   "all"            -> sab types save honge (matched + tagged)
LEAD_TYPE = "deadAbandoned"

OUTPUT_DIR = "output"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# Session use karte hain taake Windows/system ke HTTP_PROXY, HTTPS_PROXY
# env variables ya kisi aur env-based proxy setting se interference na ho.
# trust_env = False -> sirf humara diya hua PROXY use hoga, kuch aur nahi.
# Har thread ka apna Session hota hai (thread-safe rehne ke liye).
_thread_local = threading.local()


def get_session():
    if not hasattr(_thread_local, "session"):
        s = requests.Session()
        s.trust_env = False
        _thread_local.session = s
    return _thread_local.session


def get_proxy_url():
    """Build the authenticated rotating proxy URL from environment settings."""
    if not PROXY_USERNAME or not PROXY_PASSWORD:
        raise RuntimeError(
            "Proxy credentials missing. Set TSDR_PROXY_USERNAME and "
            "TSDR_PROXY_PASSWORD before starting the scraper."
        )

    host, port = random.choice(PROXY_LIST)
    username = quote(PROXY_USERNAME, safe="")
    password = quote(PROXY_PASSWORD, safe="")
    return f"http://{username}:{password}@{host}:{port}"


CSV_LOCK = threading.Lock()  # CSV files mein ek waqt mein ek hi thread likhega

# ===================== HELPERS =====================


def read_serials(file_path):
    """CSV ya XLSX se pehle column se serial numbers nikaalta hai."""
    serials = []
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".csv":
        with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            for row in reader:
                if not row:
                    continue
                val = str(row[0]).strip()
                if val and val.replace(".", "", 1).isdigit():
                    serials.append(val.split(".")[0])  # xlsx se aaye floats fix

    elif ext in (".xlsx", ".xls"):
        from openpyxl import load_workbook

        wb = load_workbook(file_path, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            if not row:
                continue
            val = row[0]
            if val is None:
                continue
            val = str(val).strip()
            if val and val.replace(".", "", 1).isdigit():
                serials.append(val.split(".")[0])
    else:
        raise ValueError("Sirf .csv, .xlsx, .xls files supported hain")

    # duplicates hatao, order maintain karo
    seen = set()
    unique_serials = []
    for s in serials:
        if s not in seen:
            seen.add(s)
            unique_serials.append(s)

    return unique_serials


def fetch_tsdr_page(serial):
    """Ek serial ke liye TSDR statusview page fetch karta hai (with retries + proxy rotation)."""
    url = f"https://tsdr.uspto.gov/statusview/sn{serial}"
    session = get_session()

    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        proxy_url = get_proxy_url() if USE_PROXY else None
        proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None

        try:
            resp = session.get(
                url,
                headers=HEADERS,
                proxies=proxies,
                timeout=TIMEOUT,
            )
            if resp.status_code == 200 and len(resp.text) > 500:
                return resp.text
            last_error = f"HTTP {resp.status_code}, len={len(resp.text)}"
        except requests.exceptions.RequestException as e:
            last_error = str(e)

        time.sleep(0.5 + attempt * 0.5)  # thoda wait next retry par (dusre proxy ke sath)

    print(f"  [FAIL] serial {serial}: {last_error}")
    return None


def clean(text):
    if text is None:
        return ""
    return re.sub(r"\s+", " ", text).strip()


# ===================== US Phone validation (content.js se port kiya) =====================
# Sirf valid 10-digit US number accept karta hai, standard format
# "XXX-XXX-XXXX" mein return karta hai. Foreign country-code prefix
# wale aur toll-free numbers (800/833/844/855/866/877/888) reject
# ho jate hain (empty string return).
TOLL_FREE_PREFIXES = {"800", "833", "844", "855", "866", "877", "888"}


def extract_us_phone(segment):
    if not segment:
        return ""
    first_line = segment.strip().split("\n")[0].strip()
    if not first_line:
        return ""

    paren_match = re.search(
        r"(1[-.\s]?)?\(\d{3}\)[-.\s]*\d{3}[-.\s]*\d{4}", first_line
    )

    if paren_match:
        before_match = first_line[: paren_match.start()]
        is_extension_prefix = bool(
            re.search(r"\d+\s*(?:x|ext\.?|extension)\.?\s*$", before_match, re.I)
        )
        if re.search(r"\d", before_match) and not is_extension_prefix:
            return ""  # foreign country code prefix - reject
        digits_only = re.sub(r"\D", "", paren_match.group(0))
    else:
        no_ext = re.sub(
            r"^\s*\d{1,8}\s*(?:x|ext\.?|extension)\.?\s*", "", first_line, flags=re.I
        )
        no_ext = re.sub(
            r"\s*(?:x|ext\.?|extension)\.?\s*\d{1,8}\s*$", "", no_ext, flags=re.I
        )
        digits_only = re.sub(r"\D", "", no_ext)

    if len(digits_only) == 10:
        ten_digits = digits_only
    elif len(digits_only) == 11 and digits_only.startswith("1"):
        ten_digits = digits_only[1:]
    else:
        return ""

    if ten_digits[:3] in TOLL_FREE_PREFIXES:
        return ""

    return f"{ten_digits[:3]}-{ten_digits[3:6]}-{ten_digits[6:]}"


# Ye emails filter kar do jo TSDR/USPTO ke generic/system addresses hain,
# actual correspondent ka personal email nahi.
EMAIL_BLACKLIST = re.compile(r"notifications|tmapp|uspto|trademark|lawinfo|legal|tmlaw", re.I)


def pick_valid_email(section_text):
    all_emails = re.findall(
        r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-z]{2,}", section_text, re.I
    )
    for e in all_emails:
        if not EMAIL_BLACKLIST.search(e):
            return e
    return ""


def find_value(soup, key_pattern):
    """soup mein di gayi regex pattern wali <div class='key'> dhoondh kar
    uske saath wali <div class='value'> ka clean text return karta hai."""
    key_div = soup.find("div", class_="key", string=re.compile(key_pattern, re.I))
    if key_div:
        val = key_div.find_next_sibling("div", class_="value")
        if val:
            return clean(val.get_text())
    return ""


def find_value_first_line(soup, key_pattern):
    """find_value jaisa hi, lekin jab value div ke andar kai nested <div>
    hon (jaise 'Correspondent Name/Address:' mein naam + address lines),
    to sirf pehli nested div ka text return karta hai (naam), baaki
    address lines ko ignore kar deta hai.

    Agar andar koi nested div na mile, to poora value text hi return
    hoga (normal single-line fields ke liye).

    Kabhi kabhi naam wali div hoti hi nahi, seedha address se shuru ho
    jata hai (jaise pehli line "2150 S CINCINNATI AVENUE" jaisi street
    address ho). Aisi cheezein naam nahi hoti, isliye pehli line agar
    address-jaisi lagti hai (number se shuru, ya PO Box, Suite, etc.)
    to khaali string return karte hain taake galat cheez 'naam' na ban
    jaye."""
    key_div = soup.find("div", class_="key", string=re.compile(key_pattern, re.I))
    if not key_div:
        return ""
    val = key_div.find_next_sibling("div", class_="value")
    if not val:
        return ""

    inner_divs = val.find_all("div", recursive=False)
    if not inner_divs:
        return clean(val.get_text())

    first_line = clean(inner_divs[0].get_text())

    # Address-jaisi lagne wali pehli line ko naam mat maano
    address_like = re.compile(
        r"^\s*(\d+[\s,]|P\.?O\.?\s*BOX|SUITE\b|STE\b|FLOOR\b|APT\b|UNIT\b)",
        re.I,
    )
    if not first_line or address_like.match(first_line):
        return ""

    return first_line


def parse_tsdr_html(html, serial):
    """
    TSDR ka HTML parse karke ek dict return karta hai jisme wo saari
    cheezein hain jo classify + filter + final output ke liye chahiye.
    """
    soup = BeautifulSoup(html, "html.parser")

    data = {
        "serial": serial,
        "mark": "",
        "status": "",
        "tm5_status": "",
        # saari possible date-types, baad mein lead_type ke hisab se
        # sahi wali "date" field mein copy hogi
        "date_cancelled": "",
        "date_abandoned": "",
        "registration_date": "",
        "filing_date": "",
        "correspondent_name": "",
        "correspondent_phone": "",
        "correspondent_email": "",
        "owner_name": "",
        "attorney_name": "",
        # Default True: agar "Attorney of Record - None" line na mile to
        # lead invalid maani jaye. Sirf exact "None" milne par False hoga.
        "is_attorney": True,
    }

    # ---- Mark ----
    data["mark"] = find_value(soup, "Mark Literal Elements")
    if not data["mark"]:
        mark_div = soup.find("div", class_="markText")
        if mark_div:
            data["mark"] = clean(mark_div.get_text())

    # ---- Status / TM5 ----
    data["status"] = find_value(soup, r"^Status:?$")

    tm5_p = soup.select("div.double.table div.value p")
    if tm5_p:
        first_text = clean(tm5_p[0].get_text())
        if first_text:
            data["tm5_status"] = first_text

    # ---- Dates (sab types ke liye alag alag) ----
    data["date_cancelled"] = find_value(soup, "Date Cancelled")
    data["date_abandoned"] = find_value(soup, "Date Abandoned") or find_value(soup, "Abandonment Date")
    data["registration_date"] = find_value(soup, r"^Registration Date:?$")
    data["filing_date"] = find_value(soup, "Application Filing Date")

    # ---- Owner Name (correspondent na milne par fallback ke liye) ----
    data["owner_name"] = find_value(soup, "Owner Name")

    # ---- Attorney of Record ----
    # ---- Correspondent (name / phone / email) ----
    for caption in soup.find_all("div", class_="caption"):
        text = clean(caption.get_text())
        low = text.lower()

        if low.startswith("attorney of record"):
            value = text.split("-", 1)[-1].strip() if "-" in text else ""
            data["attorney_name"] = value
            # Sirf exact "None" valid hai. Name ho, Not Found ho ya khaali ho -> attorney maana jayega.
            data["is_attorney"] = value.strip().lower() != "none"

        elif low.startswith("correspondent"):
            value = text.split("-", 1)[-1].strip() if "-" in text else ""
            if value.lower() not in ("not found", "none", ""):
                data["correspondent_name"] = value

    # Correspondent block mein agar name/phone/email keys (Correspondent
    # Name/Address, Phone, Email waghera) alag se mile to unhe bhi try karo.
    if not data["correspondent_name"]:
        data["correspondent_name"] = find_value_first_line(soup, "Correspondent Name")

    phone_val = find_value(soup, "Phone")
    if phone_val:
        cleaned = extract_us_phone(phone_val)
        if cleaned:
            data["correspondent_phone"] = cleaned

    email_val = find_value(soup, "Email")
    if email_val and not EMAIL_BLACKLIST.search(email_val):
        data["correspondent_email"] = email_val

    # Correspondent section ke raw text mein se "Phone:" ke baad ka
    # segment (Fax:/Correspondent e-mail: tak) nikaal kar us par
    # extract_us_phone() chalao — bilkul content.js jaisa. Isi section
    # ke text mein se saare emails dhoondh kar blacklist wala filter
    # laga kar pehla valid email chuno.
    corr_section = None
    for heading in soup.find_all("span", attrs={"data-sectiontitle": re.compile("Attorney/Correspondence", re.I)}):
        container = heading.find_parent("h2")
        if container:
            corr_section = container.find_next_sibling("div", class_="toggle_container")
        break

    if corr_section:
        section_text = corr_section.get_text("\n", strip=True)

        if not data["correspondent_phone"]:
            phone_seg_match = re.search(
                r"Phone:\s*([\s\S]*?)(?=Fax:|Correspondent e-mail:|$)",
                section_text,
                re.I,
            )
            if phone_seg_match:
                data["correspondent_phone"] = extract_us_phone(phone_seg_match.group(1))

        if not data["correspondent_email"]:
            data["correspondent_email"] = pick_valid_email(section_text)

    return data


def get_date_for_lead_type(data, lead_type):
    """lead_type ke hisab se sahi date field return karta hai."""
    if lead_type == "deadCancelled":
        return data.get("date_cancelled", "")
    if lead_type == "deadAbandoned":
        return data.get("date_abandoned", "")
    if lead_type == "liveRegister":
        return data.get("registration_date", "")
    if lead_type == "livePending":
        return data.get("filing_date", "")
    return ""


def classify_lead_type(data):
    """
    TSDR ka TM5 Common Status Descriptor (aur fallback status text) dekh kar
    ek in-4 types mein classify karta hai:
      deadAbandoned, deadCancelled, livePending, liveRegister
    Match na ho to "unknown" return karta hai.
    """
    tm5 = data.get("tm5_status", "").upper()
    status = data.get("status", "").upper()
    combined = f"{tm5} {status}"

    is_dead = "DEAD" in tm5 or not tm5 and "ABANDON" in combined
    is_live = "LIVE" in tm5

    if is_dead:
        if "ABANDON" in combined:
            return "deadAbandoned"
        if "CANCEL" in combined or "INVALIDAT" in combined or "EXPIR" in combined:
            return "deadCancelled"
        # DEAD hai lekin exact wajah clear nahi -> abandoned default (sabse common)
        return "deadAbandoned"

    if is_live:
        if "PENDING" in combined or "APPLICATION" in tm5:
            return "livePending"
        if "REGISTR" in combined or "ISSUED" in combined:
            return "liveRegister"
        return "liveRegister"

    return "unknown"


def matches_lead_type(data, lead_type):
    """Configured LEAD_TYPE ke against filter karta hai."""
    if lead_type == "all":
        return True
    return classify_lead_type(data) == lead_type


def is_missing(mark, date_value, correspondent_name, correspondent_phone):
    """
    Required fields: mark, date (type ke hisab se), correspondent name, phone.
    Email optional hai. Inme se koi bhi khaali ho to lead 'missing' consider hoti hai.
    """
    return not mark or not date_value or not correspondent_name or not correspondent_phone


# ===================== MAIN =====================


OUTPUT_FIELDS = ["serial", "lead_type", "mark", "date", "correspondent", "phone", "email"]


def process_serial(serial):
    """
    Ek serial poora process karta hai: fetch + parse + classify + filter.
    Return: (row_dict, bucket) jahan bucket = "valid" / "missing" / "skipped"
    """
    time.sleep(random.uniform(*REQUEST_DELAY))  # thoda stagger taake burst na ho
    html = fetch_tsdr_page(serial)

    if html is None:
        row = {k: "" for k in OUTPUT_FIELDS}
        row["serial"] = serial
        row["lead_type"] = "REQUEST_FAILED"
        return row, "missing"

    data = parse_tsdr_html(html, serial)
    lead_type = classify_lead_type(data)

    # ---- Attorney filter (WANT_ATTORNEY) ----
    if data["is_attorney"] != WANT_ATTORNEY:
        return {"serial": serial, "lead_type": lead_type}, "skipped"

    # ---- Lead type filter (LEAD_TYPE) ----
    if not matches_lead_type({"tm5_status": data["tm5_status"], "status": data["status"]}, LEAD_TYPE):
        return {"serial": serial, "lead_type": lead_type}, "skipped"

    date_value = get_date_for_lead_type(data, lead_type)

    # Sirf Owner Name use karo. Correspondent ko bilkul use nahi karna
    # (address mix hone ka risk hota hai), isliye fallback bhi hata diya.
    correspondent_value = data["owner_name"]

    row = {
        "serial": serial,
        "lead_type": lead_type,
        "mark": data["mark"],
        "date": date_value,
        "correspondent": correspondent_value,
        "phone": data["correspondent_phone"],
        "email": data["correspondent_email"],  # optional
    }

    if is_missing(data["mark"], date_value, correspondent_value, data["correspondent_phone"]):
        return row, "missing"

    return row, "valid"


def main():
    if len(sys.argv) < 2:
        print("Usage: python tsdr_scraper.py <serials_file.csv|.xlsx>")
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"File nahi mili: {input_file}")
        sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("Serials load ho rahe hain...")
    serials = read_serials(input_file)
    total = len(serials)
    print(f"Total {total} unique serials mile. {CONCURRENT_WORKERS} parallel workers ke sath chalega.")

    if not serials:
        print("Koi valid serial nahi mila. File check karo.")
        sys.exit(1)

    if USE_PROXY:
        try:
            get_proxy_url()
        except RuntimeError as e:
            print(f"Proxy configuration error: {e}")
            sys.exit(1)

    valid_count = 0
    missing_count = 0
    skipped_count = 0
    done_count = 0

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    valid_path = os.path.join(OUTPUT_DIR, f"leads_valid_{stamp}.csv")
    missing_path = os.path.join(OUTPUT_DIR, f"leads_missing_{stamp}.csv")

    fieldnames = OUTPUT_FIELDS

    # files ko live write karenge taake beech mein rukna pade to data safe rahe
    valid_f = open(valid_path, "w", newline="", encoding="utf-8-sig")
    missing_f = open(missing_path, "w", newline="", encoding="utf-8-sig")
    valid_writer = csv.DictWriter(valid_f, fieldnames=fieldnames, extrasaction="ignore")
    missing_writer = csv.DictWriter(missing_f, fieldnames=fieldnames, extrasaction="ignore")
    valid_writer.writeheader()
    missing_writer.writeheader()

    try:
        with ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as executor:
            future_to_serial = {executor.submit(process_serial, s): s for s in serials}

            for future in as_completed(future_to_serial):
                serial = future_to_serial[future]
                done_count += 1
                try:
                    data, bucket = future.result()
                except Exception as e:
                    print(f"[{done_count}/{total}] Serial {serial}: unexpected error -> {e}")
                    continue

                with CSV_LOCK:
                    if bucket == "missing":
                        missing_writer.writerow(data)
                        missing_f.flush()
                        missing_count += 1
                        print(f"[{done_count}/{total}] {serial} -> MISSING")
                    elif bucket == "valid":
                        valid_writer.writerow(data)
                        valid_f.flush()
                        valid_count += 1
                        print(f"[{done_count}/{total}] {serial} -> VALID [{data['lead_type']}]: {data['mark']}")
                    else:
                        skipped_count += 1
                        print(f"[{done_count}/{total}] {serial} -> SKIPPED [{data['lead_type']}]")

    except KeyboardInterrupt:
        print("\nUser ne rok diya (Ctrl+C). Ab tak ka data save ho chuka hai.")

    finally:
        valid_f.close()
        missing_f.close()

    print("\n===== DONE =====")
    print(f"Valid leads:   {valid_count}  -> {valid_path}")
    print(f"Missing leads: {missing_count} -> {missing_path}")
    if skipped_count:
        print(f"Skipped (lead type filter se match nahi hue): {skipped_count}")


if __name__ == "__main__":
    main()