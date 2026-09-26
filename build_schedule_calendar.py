#!/usr/bin/env python3
import html
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta
from pathlib import Path


SOURCE = Path("/Users/nikolaibrg/Downloads/rozklad_EZ-25-2_zapovnenyi.xlsx")
OUTPUT = Path("/Users/nikolaibrg/Desktop/пари/index.html")

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def text_of(element):
    return "".join(element.itertext()) if element is not None else ""


def col_index(cell_ref):
    letters = re.match(r"([A-Z]+)", cell_ref).group(1)
    value = 0
    for char in letters:
        value = value * 26 + ord(char) - 64
    return value


def read_workbook(path):
    with zipfile.ZipFile(path) as archive:
        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = [text_of(si) for si in root.findall("main:si", NS)]

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relmap = {
            rel.attrib["Id"]: rel.attrib["Target"].lstrip("/")
            for rel in rels.findall("pkgrel:Relationship", NS)
        }

        sheets = {}
        for sheet in workbook.find("main:sheets", NS):
            rel_id = sheet.attrib[
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            ]
            rows = []
            sheet_xml = ET.fromstring(archive.read(relmap[rel_id]))
            for row in sheet_xml.findall(".//main:sheetData/main:row", NS):
                values = []
                last_col = 0
                for cell in row.findall("main:c", NS):
                    cell_ref = cell.attrib.get("r", "A1")
                    current_col = col_index(cell_ref)
                    while last_col + 1 < current_col:
                        values.append("")
                        last_col += 1

                    cell_type = cell.attrib.get("t")
                    value_node = cell.find("main:v", NS)
                    inline_node = cell.find("main:is", NS)
                    value = ""
                    if cell_type == "s" and value_node is not None:
                        value = shared_strings[int(value_node.text)]
                    elif cell_type == "inlineStr":
                        value = text_of(inline_node)
                    elif value_node is not None:
                        value = value_node.text or ""
                    values.append(value)
                    last_col = current_col

                while values and values[-1] == "":
                    values.pop()
                if any(values):
                    rows.append(values)
            sheets[sheet.attrib["name"]] = rows
        return sheets


def parse_date(value):
    return datetime.strptime(value, "%d.%m.%Y").date()


def split_time(value):
    start, end = re.split("[–-]", value)
    return start.strip(), end.strip()


def event_datetime(day, clock):
    hour, minute = [int(part) for part in clock.split(":")]
    return datetime(day.year, day.month, day.day, hour, minute)


def build_data(sheets):
    rows = sheets["По датах"]
    header = rows[0]
    items = []
    for raw in rows[1:]:
        row = {header[i]: raw[i] if i < len(raw) else "" for i in range(len(header))}
        subject = row.get("Дисципліна", "").strip()
        if not subject or subject == "Немає пари":
            continue
        day = parse_date(row["Дата"])
        start, end = split_time(row["Час"])
        items.append(
            {
                "date": day.isoformat(),
                "displayDate": row["Дата"],
                "day": row["День"],
                "week": row["Тиждень"],
                "pair": int(row["Пара"]),
                "time": row["Час"],
                "start": start,
                "end": end,
                "subject": subject,
                "type": row.get("Вид", "").strip() or "Заняття",
                "note": row.get("Примітка", "").strip(),
                "startAt": event_datetime(day, start).isoformat(),
                "endAt": event_datetime(day, end).isoformat(),
            }
        )

    days = {}
    for item in items:
        days.setdefault(item["date"], []).append(item)
    for events in days.values():
        events.sort(key=lambda item: item["pair"])

    subjects = sorted({item["subject"] for item in items})
    types = sorted({item["type"] for item in items})
    dates = sorted(days)
    meta = {
        "group": "ЕЗ-25-2",
        "semester": "28.09.2026–26.12.2026",
        "source": SOURCE.name,
        "start": dates[0],
        "end": dates[-1],
        "subjects": subjects,
        "types": types,
        "totalEvents": len(items),
        "totalStudyDays": len(days),
    }
    return {"meta": meta, "events": items}


def month_name(month):
    names = [
        "",
        "січень",
        "лютий",
        "березень",
        "квітень",
        "травень",
        "червень",
        "липень",
        "серпень",
        "вересень",
        "жовтень",
        "листопад",
        "грудень",
    ]
    return names[month]


def build_html(data):
    data_json = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return f"""<!doctype html>
<html lang="uk">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Календар пар {html.escape(data['meta']['group'])}</title>
  <style>
    :root {{
      color-scheme: light;
      --bg: #fbf8f2;
      --panel: #fffdf8;
      --ink: #2f3432;
      --muted: #7b827f;
      --line: #ece2d6;
      --accent: #7fa99b;
      --accent-dark: #4f756b;
      --accent-2: #d7a97e;
      --lecture: #e6f2ed;
      --practice: #fff1dc;
      --training: #efeafb;
      --lesson: #e9f0fb;
      --empty: #f4efe7;
      --shadow: 0 10px 26px rgba(79, 94, 88, .08);
    }}

    * {{ box-sizing: border-box; }}

    body {{
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      font-size: 14px;
    }}

    body.modal-open {{
      overflow: hidden;
    }}

    button, input, select {{ font: inherit; }}

    .app {{
      width: min(430px, calc(100% - 18px));
      margin: 0 auto;
      padding: 10px 0 22px;
    }}

    .topbar {{
      display: grid;
      gap: 9px;
      margin-bottom: 10px;
    }}

    .eyebrow {{
      margin: 0 0 3px;
      color: var(--accent-dark);
      font-size: 11px;
      font-weight: 800;
      letter-spacing: .04em;
      text-transform: uppercase;
    }}

    h1 {{
      margin: 0;
      font-size: 24px;
      line-height: 1.06;
      letter-spacing: 0;
    }}

    .subtitle {{
      margin: 5px 0 0;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.35;
    }}

    .toolbar {{
      display: flex;
      justify-content: space-between;
      gap: 8px;
      align-items: center;
      margin-bottom: 8px;
      background: rgba(255,253,248,.92);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 7px;
      box-shadow: var(--shadow);
      position: sticky;
      top: 6px;
      z-index: 20;
      backdrop-filter: blur(14px);
    }}

    .nav-group {{
      display: inline-flex;
      gap: 5px;
      align-items: center;
      justify-content: space-between;
      width: 100%;
    }}

    .icon-btn, .text-btn, .chip {{
      border: 1px solid var(--line);
      background: var(--panel);
      color: var(--ink);
      border-radius: 8px;
      min-height: 32px;
      padding: 0 9px;
      cursor: pointer;
      transition: transform .16s ease, border-color .16s ease, background .16s ease;
    }}

    .icon-btn {{
      width: 32px;
      padding: 0;
      display: inline-grid;
      place-items: center;
      font-size: 18px;
      line-height: 1;
    }}

    .text-btn:hover, .icon-btn:hover, .chip:hover {{
      transform: translateY(-1px);
      border-color: rgba(37,107,93,.45);
    }}

    .month-title {{
      min-width: 118px;
      text-align: center;
      font-size: 15px;
      font-weight: 900;
    }}

    .layout {{
      display: grid;
      gap: 8px;
      align-items: start;
    }}

    .calendar-shell, .side-panel, .agenda-shell {{
      background: rgba(255,253,248,.94);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }}

    .weekday-row, .calendar-grid {{
      display: grid;
      grid-template-columns: repeat(7, minmax(0, 1fr));
    }}

    .weekday {{
      padding: 7px 2px;
      color: var(--muted);
      font-weight: 800;
      font-size: 10px;
      text-align: center;
      border-bottom: 1px solid var(--line);
      background: #fffaf2;
    }}

    .day-cell {{
      min-height: 58px;
      border-right: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
      padding: 4px;
      background: var(--panel);
      cursor: pointer;
      position: relative;
      transition: background .16s ease, box-shadow .16s ease;
      text-align: left;
    }}

    .day-cell:nth-child(7n) {{ border-right: 0; }}

    .day-cell:hover {{ background: #fffaf2; box-shadow: inset 0 0 0 2px rgba(127,169,155,.18); }}
    .day-cell.outside {{ background: #f7f2ea; color: #b6aaa0; }}
    .day-cell.selected {{ box-shadow: inset 0 0 0 2px var(--accent); }}
    .day-cell.today .date-num {{ background: var(--accent); color: #fff; }}

    .day-head {{
      display: flex;
      justify-content: space-between;
      gap: 8px;
      align-items: center;
      margin-bottom: 4px;
    }}

    .date-num {{
      width: 22px;
      height: 22px;
      display: inline-grid;
      place-items: center;
      border-radius: 999px;
      font-weight: 900;
      font-size: 12px;
    }}

    .week-pill {{
      color: var(--muted);
      font-size: 0;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 0;
      width: 8px;
      height: 8px;
      white-space: nowrap;
      background: var(--accent);
      opacity: .55;
    }}

    .event-mini {{
      display: inline-block;
      width: 8px;
      height: 8px;
      margin: 2px 2px 0 0;
      border-radius: 999px;
      padding: 0;
      border-left: 0;
      background: var(--lesson);
      overflow: hidden;
    }}

    .event-mini.Лекція {{ background: var(--lecture); }}
    .event-mini.Практика {{ background: var(--practice); }}
    .event-mini.Заняття {{ background: var(--lesson); }}

    .mini-time {{
      display: none;
    }}

    .mini-title {{
      display: none;
    }}

    .more {{
      color: var(--accent-dark);
      font-weight: 800;
      font-size: 10px;
      margin-top: 2px;
    }}

    .modal-backdrop {{
      position: fixed;
      inset: 0;
      z-index: 80;
      background: rgba(47, 52, 50, .24);
      opacity: 0;
      pointer-events: none;
      transition: opacity .18s ease;
    }}

    body.modal-open .modal-backdrop {{
      opacity: 1;
      pointer-events: auto;
    }}

    .side-panel {{
      position: fixed;
      left: max(9px, env(safe-area-inset-left));
      right: max(9px, env(safe-area-inset-right));
      bottom: max(9px, env(safe-area-inset-bottom));
      z-index: 90;
      max-height: min(78vh, 560px);
      padding: 11px;
      overflow: auto;
      transform: translateY(calc(100% + 18px));
      opacity: 0;
      pointer-events: none;
      transition: transform .2s ease, opacity .2s ease;
    }}

    body.modal-open .side-panel {{
      transform: translateY(0);
      opacity: 1;
      pointer-events: auto;
    }}

    .panel-header {{
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 10px;
      align-items: start;
    }}

    .close-panel {{
      width: 32px;
      height: 32px;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: #fff8eb;
      color: var(--ink);
      cursor: pointer;
      font-size: 20px;
      line-height: 1;
    }}

    .panel-date {{
      margin: 0;
      font-size: 18px;
      line-height: 1.15;
    }}

    .panel-meta {{
      margin: 4px 0 9px;
      color: var(--muted);
      font-size: 12px;
    }}

    .day-range {{
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 6px;
      margin-bottom: 8px;
    }}

    .range-box {{
      background: #fff8eb;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 8px;
    }}

    .range-box span {{
      color: var(--muted);
      display: block;
      font-size: 10px;
    }}

    .range-box strong {{
      display: block;
      margin-top: 2px;
      font-size: 17px;
    }}

    .lesson-list {{
      display: grid;
      gap: 7px;
    }}

    .lesson {{
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 9px;
      background: var(--panel);
    }}

    .lesson-top {{
      display: flex;
      justify-content: space-between;
      gap: 8px;
      align-items: center;
      margin-bottom: 6px;
    }}

    .pair {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 28px;
      height: 24px;
      border-radius: 999px;
      background: var(--accent-dark);
      color: #fff;
      font-weight: 900;
      font-size: 12px;
    }}

    .time {{
      color: var(--accent-dark);
      font-weight: 900;
      white-space: nowrap;
      font-size: 13px;
    }}

    .lesson-title {{
      margin: 0;
      font-size: 14px;
      line-height: 1.25;
    }}

    .lesson-tags {{
      display: flex;
      flex-wrap: wrap;
      gap: 5px;
      margin-top: 7px;
    }}

    .tag {{
      border-radius: 999px;
      padding: 3px 7px;
      background: #f1ece4;
      color: #4f5a55;
      font-size: 11px;
      font-weight: 750;
    }}

    .tag.lecture {{ background: var(--lecture); }}
    .tag.practice {{ background: var(--practice); }}
    .tag.note {{ background: var(--training); }}

    .empty-state {{
      min-height: 92px;
      display: grid;
      place-items: center;
      color: var(--muted);
      text-align: center;
      border: 1px dashed var(--line);
      border-radius: 8px;
      padding: 12px;
      background: #fffaf2;
    }}

    .agenda-shell {{ display: none; }}

    .agenda-list {{
      display: grid;
      gap: 0;
    }}

    .agenda-day {{
      display: grid;
      grid-template-columns: 170px 1fr;
      gap: 16px;
      padding: 16px;
      border-bottom: 1px solid var(--line);
      background: #fff;
    }}

    .agenda-date strong {{
      display: block;
      font-size: 18px;
    }}

    .agenda-date span {{
      display: block;
      color: var(--muted);
      margin-top: 4px;
    }}

    .agenda-events {{
      display: grid;
      gap: 8px;
    }}

    .agenda-card {{
      display: grid;
      grid-template-columns: 96px 1fr auto;
      gap: 12px;
      align-items: start;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px;
      background: #fbfcfa;
    }}

    .hidden {{ display: none !important; }}

    @media (min-width: 760px) {{
      body.modal-open {{ overflow: auto; }}
      .app {{ width: min(980px, calc(100% - 28px)); }}
      .topbar {{
        grid-template-columns: 1fr;
        align-items: end;
      }}
      .toolbar {{
        max-width: 620px;
      }}
      .layout {{
        grid-template-columns: minmax(0, 1fr) 300px;
        gap: 12px;
      }}
      .day-cell {{ min-height: 78px; }}
      .modal-backdrop {{ display: none; }}
      .side-panel {{
        position: sticky;
        top: 84px;
        left: auto;
        right: auto;
        bottom: auto;
        z-index: 1;
        max-height: calc(100vh - 100px);
        transform: none;
        opacity: 1;
        pointer-events: auto;
      }}
      .lesson-list {{
        max-height: calc(100vh - 245px);
        overflow: auto;
        padding-right: 2px;
      }}
      .close-panel {{ display: none; }}
    }}
  </style>
</head>
<body>
  <main class="app">
    <header class="topbar">
      <div>
        <p class="eyebrow">Розклад групи {html.escape(data['meta']['group'])}</p>
        <h1>Календар пар</h1>
        <p class="subtitle">Семестр {html.escape(data['meta']['semester'])}</p>
      </div>
    </header>

    <section class="toolbar" aria-label="Керування календарем">
      <div class="nav-group">
        <button class="icon-btn" id="prevMonth" type="button" title="Попередній місяць" aria-label="Попередній місяць">‹</button>
        <div class="month-title" id="monthTitle"></div>
        <button class="icon-btn" id="nextMonth" type="button" title="Наступний місяць" aria-label="Наступний місяць">›</button>
      </div>
    </section>

    <section class="layout">
      <section class="calendar-shell" aria-label="Місячний календар">
        <div class="weekday-row">
          <div class="weekday">Пн</div>
          <div class="weekday">Вт</div>
          <div class="weekday">Ср</div>
          <div class="weekday">Чт</div>
          <div class="weekday">Пт</div>
          <div class="weekday">Сб</div>
          <div class="weekday">Нд</div>
        </div>
        <div class="calendar-grid" id="calendarGrid"></div>
      </section>

      <div class="modal-backdrop" id="detailsBackdrop" aria-hidden="true"></div>

      <aside class="side-panel" aria-label="Деталі дня">
        <div class="panel-header">
          <h2 class="panel-date" id="panelDate"></h2>
          <button class="close-panel" id="closePanel" type="button" aria-label="Закрити деталі">×</button>
        </div>
        <p class="panel-meta" id="panelMeta"></p>
        <div class="day-range" id="dayRange"></div>
        <div class="lesson-list" id="lessonList"></div>
      </aside>

      <section class="agenda-shell" aria-label="Список занять">
        <div class="agenda-list" id="agendaList"></div>
      </section>
    </section>
  </main>

  <script>
    const schedule = {data_json};
    const events = schedule.events.map((event) => ({{
      ...event,
      dateObj: new Date(event.date + "T00:00:00"),
      startObj: new Date(event.startAt),
      endObj: new Date(event.endAt)
    }}));

    const byDate = new Map();
    for (const event of events) {{
      if (!byDate.has(event.date)) byDate.set(event.date, []);
      byDate.get(event.date).push(event);
    }}

    const monthNames = ["січень","лютий","березень","квітень","травень","червень","липень","серпень","вересень","жовтень","листопад","грудень"];
    const dayNames = ["Неділя","Понеділок","Вівторок","Середа","Четвер","П’ятниця","Субота"];
    const startDate = new Date(schedule.meta.start + "T00:00:00");
    const endDate = new Date(schedule.meta.end + "T00:00:00");
    const now = new Date();
    const firstUpcoming = events.find((event) => event.endObj >= now) || events[0];

    let currentMonth = new Date((firstUpcoming || events[0]).date + "T00:00:00");
    let selectedDate = (firstUpcoming || events[0]).date;

    const calendarGrid = document.getElementById("calendarGrid");
    const monthTitle = document.getElementById("monthTitle");
    const panelDate = document.getElementById("panelDate");
    const panelMeta = document.getElementById("panelMeta");
    const dayRange = document.getElementById("dayRange");
    const lessonList = document.getElementById("lessonList");
    const agendaList = document.getElementById("agendaList");
    const detailsBackdrop = document.getElementById("detailsBackdrop");
    const closePanel = document.getElementById("closePanel");
    const mobileDetailsQuery = window.matchMedia("(max-width: 759px)");

    function isoDate(date) {{
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      return `${{year}}-${{month}}-${{day}}`;
    }}

    function displayDate(iso) {{
      const date = new Date(iso + "T00:00:00");
      return `${{date.getDate()}} ${{monthNames[date.getMonth()]}} ${{date.getFullYear()}}`;
    }}

    function eventsForDate(iso) {{
      return byDate.get(iso) || [];
    }}

    function dayInfo(iso) {{
      const all = eventsForDate(iso);
      return all[0] || null;
    }}

    function typeClass(type) {{
      if (type === "Лекція") return "lecture";
      if (type === "Практика") return "practice";
      return "";
    }}

    function escapeHtml(value) {{
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
    }}

    function openDetailsModal() {{
      if (mobileDetailsQuery.matches) {{
        document.body.classList.add("modal-open");
      }}
    }}

    function closeDetailsModal() {{
      document.body.classList.remove("modal-open");
    }}

    function renderCalendar() {{
      const year = currentMonth.getFullYear();
      const month = currentMonth.getMonth();
      monthTitle.textContent = `${{monthNames[month]}} ${{year}}`;
      calendarGrid.innerHTML = "";

      const firstOfMonth = new Date(year, month, 1);
      const mondayOffset = (firstOfMonth.getDay() + 6) % 7;
      const gridStart = new Date(year, month, 1 - mondayOffset);

      for (let index = 0; index < 42; index++) {{
        const date = new Date(gridStart);
        date.setDate(gridStart.getDate() + index);
        const iso = isoDate(date);
        const filteredEvents = eventsForDate(iso);
        const info = dayInfo(iso);

        const cell = document.createElement("button");
        cell.type = "button";
        cell.className = "day-cell";
        if (date.getMonth() !== month) cell.classList.add("outside");
        if (iso === selectedDate) cell.classList.add("selected");
        if (iso === isoDate(new Date())) cell.classList.add("today");
        cell.addEventListener("click", () => {{
          selectedDate = iso;
          render();
          openDetailsModal();
        }});

        const head = document.createElement("div");
        head.className = "day-head";
        head.innerHTML = `<span class="date-num">${{date.getDate()}}</span>${{info ? `<span class="week-pill">${{info.week}}</span>` : ""}}`;
        cell.appendChild(head);

        for (const event of filteredEvents.slice(0, 3)) {{
          const mini = document.createElement("div");
          mini.className = `event-mini ${{event.type}}`;
          mini.innerHTML = `<span class="mini-time">${{event.start}}-${{event.end}}</span><span class="mini-title">${{escapeHtml(event.subject)}}</span>`;
          cell.appendChild(mini);
        }}

        if (filteredEvents.length > 3) {{
          const more = document.createElement("div");
          more.className = "more";
          more.textContent = `+${{filteredEvents.length - 3}} ще`;
          cell.appendChild(more);
        }}

        calendarGrid.appendChild(cell);
      }}
    }}

    function renderPanel() {{
      const events = eventsForDate(selectedDate);
      const info = dayInfo(selectedDate);
      const date = new Date(selectedDate + "T00:00:00");
      panelDate.textContent = `${{dayNames[date.getDay()]}}, ${{displayDate(selectedDate)}}`;
      panelMeta.textContent = info ? `${{info.week}} • ${{events.length}} пар` : "Немає занять у таблиці";

      if (events.length) {{
        dayRange.innerHTML = `
          <div class="range-box"><span>Початок</span><strong>${{events[0].start}}</strong></div>
          <div class="range-box"><span>Кінець</span><strong>${{events[events.length - 1].end}}</strong></div>
        `;
        lessonList.innerHTML = events.map((event) => `
          <article class="lesson">
            <div class="lesson-top">
              <span class="pair">${{event.pair}}</span>
              <span class="time">${{event.start}}-${{event.end}}</span>
            </div>
            <h3 class="lesson-title">${{escapeHtml(event.subject)}}</h3>
            <div class="lesson-tags">
              <span class="tag ${{typeClass(event.type)}}">${{escapeHtml(event.type)}}</span>
              <span class="tag">${{escapeHtml(event.week)}}</span>
              ${{event.note ? `<span class="tag note">${{escapeHtml(event.note)}}</span>` : ""}}
            </div>
          </article>
        `).join("");
      }} else {{
        dayRange.innerHTML = "";
        lessonList.innerHTML = `<div class="empty-state">На цей день за вибраними фільтрами пар немає.</div>`;
      }}
    }}

    function renderAgenda() {{
      const dates = [...byDate.keys()].sort();
      const visibleDays = dates.map((iso) => [iso, eventsForDate(iso)]).filter(([, events]) => events.length);
      agendaList.innerHTML = visibleDays.map(([iso, events]) => {{
        const info = dayInfo(iso);
        return `
          <article class="agenda-day">
            <div class="agenda-date">
              <strong>${{displayDate(iso)}}</strong>
              <span>${{escapeHtml(info.day)}} • ${{escapeHtml(info.week)}}</span>
            </div>
            <div class="agenda-events">
              ${{events.map((event) => `
                <button class="agenda-card" type="button" data-date="${{iso}}">
                  <span class="time">${{event.start}}-${{event.end}}</span>
                  <strong>${{escapeHtml(event.subject)}}</strong>
                  <span class="tag ${{typeClass(event.type)}}">${{escapeHtml(event.type)}}</span>
                </button>
              `).join("")}}
            </div>
          </article>
        `;
      }}).join("") || `<div class="empty-state">За цими фільтрами нічого не знайдено.</div>`;

      for (const card of agendaList.querySelectorAll(".agenda-card")) {{
        card.addEventListener("click", () => {{
          selectedDate = card.dataset.date;
          currentMonth = new Date(selectedDate + "T00:00:00");
          render();
          window.scrollTo({{ top: 0, behavior: "smooth" }});
        }});
      }}
    }}

    function render() {{
      renderCalendar();
      renderPanel();
      renderAgenda();
    }}

    document.getElementById("prevMonth").addEventListener("click", () => {{
      currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1);
      render();
    }});

    document.getElementById("nextMonth").addEventListener("click", () => {{
      currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
      render();
    }});

    closePanel.addEventListener("click", closeDetailsModal);
    detailsBackdrop.addEventListener("click", closeDetailsModal);

    document.addEventListener("keydown", (event) => {{
      if (event.key === "Escape") closeDetailsModal();
    }});

    mobileDetailsQuery.addEventListener("change", (event) => {{
      if (!event.matches) closeDetailsModal();
    }});

    render();
  </script>
</body>
</html>
"""


def main():
    sheets = read_workbook(SOURCE)
    data = build_data(sheets)
    OUTPUT.write_text(build_html(data), encoding="utf-8")
    print(OUTPUT)
    print(f"{data['meta']['totalEvents']} events, {data['meta']['totalStudyDays']} study days")


if __name__ == "__main__":
    main()
