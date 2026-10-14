"""Build the trainer evaluations site.

Writes site/ with:
  index.html, forms.json, vendor/      (copied as-is)
  config.json   trainer list (trainers.txt) + where Save & Send goes (EVAL_SUBMIT_URL secret)
  data.json     employee list for the picker, from the SharePoint file sent by the
                Power Automate flow in the EVAL_DATA_URL secret (optional until it's set up)

Local test with a file instead of the flow:
    python build.py "Employee file.xlsx"
"""
import io
import os
import sys
import json
import shutil

import requests
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "site")
SECRETS = json.loads(os.environ.get("ALL_SECRETS") or "{}")


def secret(name):
    return SECRETS.get(name) or os.environ.get(name) or ""


def norm(text):
    return "".join(ch for ch in str(text or "").lower() if ch.isalnum())


ROSTER_COLUMNS = {
    "name": ["name", "employeename", "employee", "fullname", "trainee", "membername"],
    "first": ["firstname", "first"],
    "last": ["lastname", "last"],
    "id": ["employeeid", "empid", "id", "loginid", "eeid"],
    "office": ["office", "homeoffice", "location", "team", "teams", "branch"],
}


def get_data_file():
    if len(sys.argv) > 1:
        with open(sys.argv[1], "rb") as f:
            return f.read()
    url = secret("EVAL_DATA_URL")
    if not url:
        print("No EVAL_DATA_URL secret yet: employee names will be typed in.")
        return None
    res = requests.post(url, json={}, timeout=180)
    if res.status_code != 200 or not res.content.startswith(b"PK"):
        raise SystemExit(f"The employee-file flow didn't send back an Excel file (status {res.status_code}).")
    return res.content


def read_roster(xlsx_bytes):
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for h, row in enumerate(rows[:15]):
            headers = [norm(c) for c in row]
            found = {}
            for key, names in ROSTER_COLUMNS.items():
                for want in names:
                    if want in headers:
                        found[key] = headers.index(want)
                        break
            if "name" in found or ("first" in found and "last" in found):
                people = {}
                for r in rows[h + 1:]:
                    get = lambda k: (str(r[found[k]]).strip() if k in found and found[k] < len(r) and r[found[k]] is not None else "")
                    name = get("name") or (get("first") + " " + get("last")).strip()
                    if not name:
                        continue
                    key = get("id") or name.lower()
                    people.setdefault(key, {"name": name, "id": get("id"), "office": get("office")})
                roster = sorted(people.values(), key=lambda p: p["name"].lower())
                print(f"Employee list: {len(roster)} people from sheet '{ws.title}'.")
                return roster
    raise SystemExit("Couldn't find a Name column (or First/Last Name) in the employee file.")


def main():
    shutil.rmtree(OUT, ignore_errors=True)
    os.makedirs(OUT)
    for name in ("index.html", "forms.json"):
        shutil.copy(os.path.join(HERE, name), OUT)
    shutil.copytree(os.path.join(HERE, "vendor"), os.path.join(OUT, "vendor"))

    trainers = []
    path = os.path.join(HERE, "trainers.txt")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            trainers = [l.strip() for l in f if l.strip() and not l.lstrip().startswith("#")]
    submit = secret("EVAL_SUBMIT_URL")
    if not submit:
        print("::warning title=Save & Send not connected::Add the EVAL_SUBMIT_URL secret; until then trainers can only download the PDF.")
    with open(os.path.join(OUT, "config.json"), "w") as f:
        json.dump({"trainers": trainers, "submitUrl": submit}, f)

    data = get_data_file()
    if data:
        with open(os.path.join(OUT, "data.json"), "w") as f:
            json.dump({"roster": read_roster(data)}, f, separators=(",", ":"))
    with open(os.path.join(OUT, "robots.txt"), "w") as f:
        f.write("User-agent: *\nDisallow: /\n")
    print(f"::notice title=Trainer evaluations::{len(trainers)} trainers, Save & Send {'on' if submit else 'off'}, "
          f"employee list {'loaded' if data else 'not connected'}")


if __name__ == "__main__":
    main()
