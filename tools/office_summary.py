"""Office (department) Training Cycle summary for managers, as a PDF.

One PDF per office, built from the same data as the trainer website:
  history   - training cycle members and their HSI training status (history.json)
  tracker   - Trainer Tracker rows (promotion recommendation, trainer/HSI notes), optional
  needs     - {nameKey: [{date, form, trainer, needs:[[section,[items]]]}]} from the Evaluation Log, optional
  comments  - {nameKey: [{date, trainer, comment}]} from the website's Comments, optional

Layout (manager first: what to act on, then the detail):
  1. At a glance             counts for the office
  2. Action items            delinquent items, cycles ending soon with open items,
                             trainings to schedule (grouped by training), online work to assign
  3. Roster                  everyone currently in the program, one line each
  4. Employee detail         per person: recommendation, dates, completed, still needed
                             (online vs. with a trainer), delinquent, evaluation needs, notes, comments
  5. Recently ended cycles   recommendation outcome
  6. No longer with the company
"""
import io
import re
from datetime import date, datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (CondPageBreak, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

NAVY = colors.HexColor("#22306B")
ORANGE = colors.HexColor("#F05023")
INK = colors.HexColor("#1F2430")
MUTED = colors.HexColor("#5B6475")
LINE = colors.HexColor("#D9DEE8")
GREEN = colors.HexColor("#E2F0D9")
YELLOW = colors.HexColor("#FFF2CC")
RED = colors.HexColor("#F8D7DA")
BLUE = colors.HexColor("#DDEBF7")
GRAY = colors.HexColor("#EEF0F4")

# Trainings an employee does on their own online; everything else is scheduled or done with a trainer.
SELF_ONLINE = ("24 or 40 Hour Hazwoper with Hands On", "Tech II to Tech III Post Test Assessment", "OSHA 30hr")
DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")

S = {
    "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=20, leading=24, textColor=NAVY),
    "sub": ParagraphStyle("sub", fontName="Helvetica", fontSize=10, leading=13, textColor=MUTED),
    "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=NAVY, spaceBefore=12, spaceAfter=6),
    "h3": ParagraphStyle("h3", fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=INK, spaceBefore=4, spaceAfter=2),
    "body": ParagraphStyle("body", fontName="Helvetica", fontSize=9, leading=12, textColor=INK),
    "small": ParagraphStyle("small", fontName="Helvetica", fontSize=8, leading=10, textColor=MUTED),
    "cell": ParagraphStyle("cell", fontName="Helvetica", fontSize=8, leading=10, textColor=INK, alignment=TA_LEFT),
    "cellb": ParagraphStyle("cellb", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=INK),
    "lab": ParagraphStyle("lab", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=MUTED),
    "tile_n": ParagraphStyle("tile_n", fontName="Helvetica-Bold", fontSize=18, leading=21, textColor=NAVY),
    "tile_l": ParagraphStyle("tile_l", fontName="Helvetica", fontSize=7.5, leading=9, textColor=MUTED),
}


def esc(t):
    return (str(t or "")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def name_key(n):
    n = str(n or "")
    if "," in n:
        last, first = n.split(",", 1)
        n = first + " " + last
    t = [x for x in re.sub(r"[^a-z ]", " ", n.lower()).split() if len(x) > 1]
    return f"{t[0]} {t[-1]}" if t else ""


def _d(iso):
    return datetime.strptime(iso, "%Y-%m-%d").date() if iso else None


def mdy(iso):
    d = _d(iso)
    return f"{d.month}/{d.day}/{d.year}" if d else ""


def months_left(end, today):
    d = _d(end)
    if not d:
        return None
    return max(0, (d.year - today.year) * 12 + d.month - today.month)


def is_current(p, today_iso):
    if not p.get("employed", True):
        return False
    return p["end"] >= today_iso if p.get("end") else bool(p.get("inHSIReport"))


def status_lists(p):
    """completed [(t, date)], online [t], scheduled [t], delinquent [t] for one person."""
    done, online, sched, late = [], [], [], []
    for t, v in (p.get("trainings") or {}).items():
        v = str(v or "").strip()
        if not v or v.upper() == "N/A":
            continue
        if DATE_RE.match(v):
            done.append((t, v))
        elif v.lower().startswith("delinq"):
            late.append(t)
        elif t in SELF_ONLINE:
            online.append(t)
        else:
            sched.append(t)
    return done, online, sched, late


def rec_color(rec, not_rec):
    r = (rec or "").lower()
    if not_rec or "not recommended" in r:
        return RED
    if "fully" in r:
        return GREEN
    if "on track" in r:
        return BLUE
    if r:
        return YELLOW
    return None


def _bullets(names, width=4.0 * inch):
    """Names as a bulleted list, in two columns when there are more than four."""
    items = [Paragraph("•&nbsp;&nbsp;" + esc(n), S["cell"]) for n in names]
    if len(items) <= 4:
        rows, widths = [[i] for i in items], [width]
    else:
        half = (len(items) + 1) // 2
        left, right = items[:half], items[half:] + [""] * (2 * half - len(items))
        rows, widths = [[a, b] for a, b in zip(left, right)], [width / 2] * 2
    t = Table(rows, colWidths=widths)
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                           ("TOPPADDING", (0, 0), (-1, -1), 0.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 0.5),
                           ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    return t


def _tbl(data, widths, header=True, extra=None):
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    st = [("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 8),
          ("VALIGN", (0, 0), (-1, -1), "TOP"), ("GRID", (0, 0), (-1, -1), 0.4, LINE),
          ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    if header:
        st += [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
               ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")]
    t.setStyle(TableStyle(st + (extra or [])))
    return t


def _tiles(items):
    cells = [[Paragraph(str(n), S["tile_n"]), Paragraph(esc(l), S["tile_l"])] for n, l, _ in items]
    t = Table([[Table([[c[0]], [c[1]]], colWidths=[1.15 * inch]) for c in cells]], colWidths=[1.22 * inch] * len(items))
    st = [("VALIGN", (0, 0), (-1, -1), "TOP")]
    for i, (_, _, bg) in enumerate(items):
        st += [("BACKGROUND", (i, 0), (i, 0), bg or GRAY), ("BOX", (i, 0), (i, 0), 0.6, LINE)]
    t.setStyle(TableStyle(st + [("LEFTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 6),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return t


def build(history, office, today=None, tracker=None, needs=None, comments=None):
    today = today or date.today()
    tiso = today.isoformat()
    needs, comments = needs or {}, comments or {}
    trk = {}
    for r in (tracker or []):
        k = name_key(r.get("Employee"))
        if k:
            trk[k] = r
    people = [p for p in history["people"] if p.get("office") == office]
    cur = sorted([p for p in people if is_current(p, tiso)], key=lambda p: (p.get("end") or "9999", p["name"]))
    # cycles that ended in the last 120 days, still employed
    ended = sorted([p for p in people if p.get("employed", True) and not is_current(p, tiso) and p.get("end")
                    and (today - _d(p["end"])).days <= 120], key=lambda p: (p["end"], p["name"]), reverse=True)
    gone = sorted([p for p in people if not p.get("employed", True)], key=lambda p: p["name"])
    office_name = re.sub(r"^\d+-", "", office)

    def rec(p):
        r = trk.get(name_key(p["name"]), {})
        return (r.get("Promtion Recommendation Status") or r.get("Promotion Recommendation Status") or "").strip()

    info = {p["name"]: status_lists(p) for p in cur}
    n_late = sum(len(v[3]) for v in info.values())
    n_need = sum(len(v[1]) + len(v[2]) for v in info.values())
    soon = [p for p in cur if p.get("end") and (_d(p["end"]) - today).days <= 60 and (info[p["name"]][1] or info[p["name"]][2] or info[p["name"]][3])]
    new = [p for p in cur if p.get("start") and 0 <= (today - _d(p["start"])).days <= 60]
    full = [p for p in cur + ended if "fully" in rec(p).lower()]

    story = []
    story += [Paragraph(f"{esc(office_name)} — Level Up Training Cycle Summary", S["h1"]),
              Paragraph(f"{esc(office)} · as of {today.strftime('%B %d, %Y').replace(' 0', ' ')} · "
                        "from the HSI Cohort Daily Report, Trainer Tracker and trainer evaluations", S["sub"]),
              Spacer(1, 10)]

    # 1. At a glance
    story.append(_tiles([
        (len(cur), "in a current training cycle", BLUE),
        (len(new), "new to the program (last 60 days)", GRAY),
        (len(soon), "cycles ending in 60 days with items open", YELLOW if soon else GRAY),
        (n_late, "delinquent items", RED if n_late else GRAY),
        (n_need, "items still needed", YELLOW if n_need else GRAY),
        (len(full), "fully recommended", GREEN if full else GRAY),
    ]))

    # 2. Action items
    story.append(Paragraph("Action items for the manager", S["h2"]))
    late_rows = [[Paragraph(esc(p["name"]), S["cellb"]), Paragraph(esc(t), S["cell"]), mdy(p.get("end"))]
                 for p in cur for t in info[p["name"]][3]]
    if late_rows:
        story.append(Paragraph("Delinquent — schedule these first", S["h3"]))
        story.append(_tbl([["Employee", "Training", "Cycle ends"]] + late_rows, [1.9 * inch, 4.1 * inch, 1.0 * inch],
                          extra=[("BACKGROUND", (0, 1), (-1, -1), RED)]))
    if soon:
        story.append(Paragraph("Training cycles ending within 60 days with items still open", S["h3"]))
        rows = [[Paragraph(esc(p["name"]), S["cellb"]), mdy(p.get("end")),
                 str(len(info[p["name"]][1]) + len(info[p["name"]][2]) + len(info[p["name"]][3])),
                 Paragraph(esc(rec(p) or "—"), S["cell"])] for p in soon]
        story.append(_tbl([["Employee", "Cycle ends", "Open items", "Recommendation"]] + rows,
                          [1.9 * inch, 0.9 * inch, 0.8 * inch, 3.4 * inch], extra=[("BACKGROUND", (0, 1), (-1, -1), YELLOW)]))
    by_training = {}
    for p in cur:
        for t in info[p["name"]][2] + [x for x in info[p["name"]][3] if x not in SELF_ONLINE]:
            by_training.setdefault(t, []).append(p["name"])
    if by_training:
        story.append(Paragraph("To schedule with a trainer — grouped by training", S["h3"]))
        rows = [[Paragraph(esc(t), S["cellb"]), str(len(ns)), _bullets(sorted(ns))]
                for t, ns in sorted(by_training.items(), key=lambda x: (-len(x[1]), x[0]))]
        story.append(_tbl([["Training", "People", "Who"]] + rows, [2.4 * inch, 0.6 * inch, 4.0 * inch]))
    online = {}
    for p in cur:
        for t in info[p["name"]][1] + [x for x in info[p["name"]][3] if x in SELF_ONLINE]:
            online.setdefault(t, []).append(p["name"])
    if online:
        story.append(Paragraph("Self work — online (assign or remind)", S["h3"]))
        rows = [[Paragraph(esc(t), S["cellb"]), str(len(ns)), _bullets(sorted(ns))]
                for t, ns in sorted(online.items(), key=lambda x: (-len(x[1]), x[0]))]
        story.append(_tbl([["Training", "People", "Who"]] + rows, [2.4 * inch, 0.6 * inch, 4.0 * inch]))
    if not (late_rows or soon or by_training or online):
        story.append(Paragraph("Nothing open. Everyone in a current training cycle is complete.", S["body"]))

    # 3. Roster
    story.append(CondPageBreak(2.5 * inch))
    story.append(Paragraph(f"Everyone in the program ({len(cur)})", S["h2"]))
    rows, ext = [], []
    for i, p in enumerate(cur, 1):
        done, onl, sch, late = info[p["name"]]
        total = len(done) + len(onl) + len(sch) + len(late)
        r = rec(p) or ("Not Recommended" if p.get("notRecommended") else ("New to Program" if p in new else ""))
        rows.append([Paragraph(("* " if p in new else "") + esc(p["name"]), S["cellb"]),
                     Paragraph(esc(f"{p.get('jobTitle', '')} → {p.get('target', '')}"), S["cell"]),
                     mdy(p.get("end")), f"{len(done)}/{total}", str(len(late)) if late else "",
                     Paragraph(esc(r or "—"), S["cell"])])
        ext += [("BACKGROUND", (0, i), (0, i), BLUE)]
        if late:
            ext += [("BACKGROUND", (4, i), (4, i), RED)]
        c = rec_color(r, p.get("notRecommended"))
        if c:
            ext += [("BACKGROUND", (5, i), (5, i), c)]
    if rows:
        story.append(_tbl([["Employee", "Position → Target", "Cycle ends", "Done", "Late", "Recommendation"]] + rows,
                          [1.6 * inch, 1.9 * inch, 0.75 * inch, 0.5 * inch, 0.45 * inch, 1.8 * inch], extra=ext))
        story.append(Paragraph("* new to the program in the last 60 days. Done = completed / required trainings. "
                               "Late = delinquent items.", S["small"]))

    # 4. Employee detail
    story.append(CondPageBreak(3 * inch))
    story.append(Paragraph("Employee detail", S["h2"]))
    group = None
    for p in cur:
        ml = months_left(p.get("end"), today)
        g = f"Training cycle ending {_d(p['end']).strftime('%B %Y')}" + (f" — {ml} month{'s' if ml != 1 else ''} left" if ml is not None else "") if p.get("end") else "No end date"
        if g != group:
            story.append(Paragraph(esc(g), S["h3"]))
            group = g
        story.append(_person_block(p, info[p["name"]], rec(p), trk.get(name_key(p["name"]), {}),
                                   needs.get(name_key(p["name"]), []), comments.get(name_key(p["name"]), []), p in new))

    # 5. Recently ended
    if ended:
        story.append(CondPageBreak(2 * inch))
        story.append(Paragraph("Recently ended training cycles", S["h2"]))
        rows, ext = [], []
        for i, p in enumerate(ended, 1):
            r = rec(p) or ("Not Recommended" if p.get("notRecommended") else "—")
            t = trk.get(name_key(p["name"]), {})
            rows.append([Paragraph(esc(p["name"]), S["cellb"]), Paragraph(esc(f"{p.get('jobTitle', '')} → {p.get('target', '')}"), S["cell"]),
                         mdy(p.get("end")), Paragraph(esc(r), S["cell"]),
                         Paragraph(esc(t.get("Promoted to") or t.get("If not, why") or ""), S["cell"])])
            c = rec_color(r, p.get("notRecommended"))
            if c:
                ext.append(("BACKGROUND", (3, i), (3, i), c))
        story.append(_tbl([["Employee", "Position → Target", "Ended", "Recommendation", "Promoted to"]] + rows,
                          [1.6 * inch, 1.8 * inch, 0.75 * inch, 1.6 * inch, 1.25 * inch], extra=ext))

    # 6. No longer with the company
    if gone:
        story.append(Paragraph("No longer with the company", S["h2"]))
        story.append(Paragraph("; ".join(f"{esc(p['name'])} ({esc(p.get('jobTitle', ''))})" for p in gone), S["body"]))

    story += [Spacer(1, 10), Paragraph(
        "Colors: light green = complete / fully recommended · light yellow = still needed · light red = delinquent / "
        "not recommended · light blue = in a current training cycle. Self work — online: HAZWOPER, Tech II to Tech III "
        "Post Test, OSHA 30; everything else is scheduled or done with a trainer.", S["small"])]

    buf = io.BytesIO()

    def page(c, d):
        c.saveState()
        c.setFillColor(NAVY)
        c.rect(0, letter[1] - 0.32 * inch, letter[0], 0.32 * inch, stroke=0, fill=1)
        c.setFillColor(ORANGE)
        c.rect(0, letter[1] - 0.36 * inch, letter[0], 0.04 * inch, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(0.6 * inch, letter[1] - 0.21 * inch, "ENVIROVAC UNIVERSITY · LET'S LEVEL UP")
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 7.5)
        c.drawString(0.6 * inch, 0.4 * inch, f"{office_name} Training Cycle Summary · {today.strftime('%m/%d/%Y')}")
        c.drawRightString(letter[0] - 0.6 * inch, 0.4 * inch, f"Page {d.page}")
        c.restoreState()

    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                            topMargin=0.6 * inch, bottomMargin=0.65 * inch,
                            title=f"{office_name} Training Cycle Summary", author="EnviroVac Training")
    doc.build(story, onFirstPage=page, onLaterPages=page)
    return buf.getvalue()


def _person_block(p, lists, rec, trk, needs, comments, is_new):
    done, onl, sch, late = lists
    total = len(done) + len(onl) + len(sch) + len(late)
    r = rec or ("Not Recommended" if p.get("notRecommended") else ("New to Program" if is_new else "—"))
    head = Table([[Paragraph(f"<b>{esc(p['name'])}</b> — {esc(p.get('jobTitle', ''))} → {esc(p.get('target', ''))}", S["body"]),
                   Paragraph(f"<b>{len(done)}/{total}</b> complete", S["body"])]],
                 colWidths=[5.6 * inch, 1.4 * inch])
    head.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), BLUE), ("BACKGROUND", (1, 0), (1, 0), GRAY),
                              ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    rows = [[Paragraph("Recommendation", S["lab"]), Paragraph(esc(r), S["cell"])],
            [Paragraph("Training dates", S["lab"]), Paragraph(esc(f"{mdy(p.get('start'))} to {mdy(p.get('end'))}  ·  {p.get('cohort', '')}"), S["cell"])]]
    ext = []
    c = rec_color(r, p.get("notRecommended"))
    if c:
        ext.append(("BACKGROUND", (1, 0), (1, 0), c))
    if late:
        rows.append([Paragraph("Delinquent", S["lab"]), Paragraph(esc("; ".join(late)), S["cell"])])
        ext.append(("BACKGROUND", (1, len(rows) - 1), (1, len(rows) - 1), RED))
    if onl:
        rows.append([Paragraph("Self work — online", S["lab"]), Paragraph(esc("; ".join(onl)), S["cell"])])
        ext.append(("BACKGROUND", (1, len(rows) - 1), (1, len(rows) - 1), YELLOW))
    if sch:
        rows.append([Paragraph("Schedule with a trainer", S["lab"]), Paragraph(esc("; ".join(sch)), S["cell"])])
        ext.append(("BACKGROUND", (1, len(rows) - 1), (1, len(rows) - 1), YELLOW))
    if done:
        rows.append([Paragraph("Completed", S["lab"]), Paragraph(esc("; ".join(f"{t} ({d})" for t, d in sorted(done, key=lambda x: x[0]))), S["cell"])])
        ext.append(("BACKGROUND", (1, len(rows) - 1), (1, len(rows) - 1), GREEN))
    for e in needs[:3]:
        items = "; ".join((f"{sec}: " if sec else "") + ", ".join(it) for sec, it in e["needs"])
        rows.append([Paragraph("Evaluation needs", S["lab"]),
                     Paragraph(f"<b>{esc(e['form'])}</b> {esc(e['date'])}{(' · ' + esc(e['trainer'])) if e.get('trainer') else ''}: {esc(items)}", S["cell"])])
    notes = [("Skill trainer feedback", trk.get("Evaluation Comments")), ("HSI training note", trk.get("HSI Training Comments")),
             ("Manager note", trk.get("Notes")), ("Not ready because", trk.get("If not, why"))]
    for lab, v in notes:
        if v and str(v).strip():
            rows.append([Paragraph(lab, S["lab"]), Paragraph("“" + esc(str(v).strip()) + "”", S["cell"])])
    for c2 in comments[:3]:
        rows.append([Paragraph("Trainer comment", S["lab"]),
                     Paragraph(f"{esc(c2.get('date', ''))}{(' · ' + esc(c2['trainer'])) if c2.get('trainer') else ''}: {esc(c2.get('comment', ''))}", S["cell"])])
    body = Table(rows, colWidths=[1.35 * inch, 5.65 * inch])
    body.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBELOW", (0, 0), (-1, -1), 0.3, LINE),
                              ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)] + ext))
    return KeepTogether([head, body, Spacer(1, 8)])


if __name__ == "__main__":
    import json, sys
    h = json.load(open(sys.argv[1]))
    trk = json.load(open(sys.argv[2]))["active"] + json.load(open(sys.argv[2]))["termCohort"] if len(sys.argv) > 2 and sys.argv[2] != "-" else None
    office = sys.argv[3] if len(sys.argv) > 3 else "007-Sumter"
    out = sys.argv[4] if len(sys.argv) > 4 else f"{office} Training Cycle Summary.pdf"
    open(out, "wb").write(build(h, office, tracker=trk))
    print("wrote", out)
