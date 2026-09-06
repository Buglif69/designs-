"""Builds Javiar_Vehicle_Conduct_Review_Evidence_Pack.xlsx.

Source of the timeline rows: ServiceM8 job diaries (check-in / check-out /
on-the-way SMS stamps) pulled 6 Sep 2026. Camera event rows come from
Fleeta Cloud (BlackVue) and are entered by hand as clips are reviewed.
Re-run:  python build_evidence_pack.py
"""
from datetime import date, time, timedelta
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

OUT = "Javiar_Vehicle_Conduct_Review_Evidence_Pack.xlsx"

FONT = "Arial"
NAVY = "1F3A5F"
GREY = "F2F2F2"
YELLOW = "FFF2CC"
RED = "F8CBAD"
GREEN = "E2EFDA"

hdr_font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
hdr_fill = PatternFill("solid", fgColor=NAVY)
body = Font(name=FONT, size=10)
bold = Font(name=FONT, size=10, bold=True)
title_font = Font(name=FONT, size=14, bold=True, color=NAVY)
note_font = Font(name=FONT, size=9, italic=True, color="595959")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="top", wrap_text=True)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = center
        cell.border = border


def style_body(ws, r1, r2, ncols, fill=None):
    for r in range(r1, r2 + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            if cell.font == Font():
                cell.font = body
            cell.alignment = wrap
            cell.border = border
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def td(h=0, m=0, s=0):
    return timedelta(hours=h, minutes=m, seconds=s)


wb = Workbook()

# ---------------------------------------------------------------- Summary
ws = wb.active
ws.title = "Summary"
ws["A1"] = "Vehicle Conduct Review - Evidence Pack"
ws["A1"].font = title_font
ws["A2"] = ("Vehicle: 'king Javiar' (Raptor, BlackVue via Fleeta Cloud).  "
            "Driver: Javiar Cartledge.  Prepared 06/09/2026.  CONFIDENTIAL - HR file.")
ws["A2"].font = note_font
ws.merge_cells("A2:D2")

rows = [
    ("Item", "Finding / status", "Source", "Next action"),
    ("Fri 28 Aug & Sat 29 Aug - grey days in Fleeta",
     "0 jobs allocated to Javiar either day (other techs had 5 jobs on 28/08). "
     "Zero camera events is consistent with the vehicle not being driven. Not evidence of a fault or of concealment.",
     "ServiceM8 schedule, Fleeta calendar", "Close out. Optional: Fleeta Trip History 28-29/08 to confirm zero ignition."),
    ("Fri 4 Sep - 40+ triggered events",
     "7 stops, ~180 km (estimate), Sunnybank > Fortitude Valley > Thornlands > Murarrie > Morayfield > Narangba > Murrumba Downs. "
     "Long motorway legs. Arrived 63 min early to last job. One 33-second non-customer stop (Narangba, 'drop cash', staff address).",
     "ServiceM8 diaries + Fleeta", "Match every event to a leg on the Timeline sheet. Download all Driving Impact clips."),
    ("Sat 5 Sep - 17 events 07:03-08:34",
     "Chapel Hill job (booked 08:00): travel clock started 05:19, checked out 07:46 with 4 seconds on site, job not done. "
     "Then bed bug job South Brisbane 08:11-08:20 (9 min, 1 mandatory task skipped). Parked 08:34 (Parking Impact). "
     "Reported sick by ~11:07; Jason took his 11:15 job.",
     "ServiceM8 diaries + Fleeta", "PRIORITY: Fleeta Trip History 05:00-08:40 on 05/09 to account for the 2h27m. Download all 17 clips."),
    ("Cabin audio",
     "Unknown until a clip is opened. Qld Invasion of Privacy Act 1971 s43: recording a private conversation you are not party to is an offence. "
     "If audio is ON, do NOT listen, share or rely on it until a lawyer has advised. Treat it as a liability, not an asset.",
     "BlackVue config", "Check camera audio setting. Get legal advice before any use. Consider turning audio OFF fleet-wide with written notice."),
    ("SD card",
     "August already overwritten. September continuous footage still on card but cycling. Cloud only holds trigger clips, "
     "the card holds the footage BETWEEN triggers (the real driving picture).",
     "BlackVue SD picker", "Pull the card Mon 07/09 before the vehicle moves. Image it, hash it, bag it. Fit a fresh formatted card."),
    ("Passenger",
     "Identity, role (staff / trainee / customer / other) and whether they have made a complaint are not yet on file. "
     "This decides which process runs: driving/WHS conduct, or harassment complaint (positive duty), or both.",
     "Jason", "Confirm before the allegations letter is drafted. If a complaint exists, the complainant's account comes first."),
    ("Process risk",
     "Driver shares the owner's surname. Perceived conflict of interest in any disciplinary outcome. Small business (<15 staff): "
     "Small Business Fair Dismissal Code applies but procedural fairness is still required.",
     "Staff list", "Have an independent person (external HR adviser) run the meeting and record it."),
]
for i, r in enumerate(rows, start=4):
    for j, v in enumerate(r, start=1):
        ws.cell(row=i, column=j, value=v)
style_header(ws, 4, 4)
style_body(ws, 5, 4 + len(rows) - 1, 4)
ws.cell(row=8, column=2).fill = PatternFill("solid", fgColor=RED)
ws.cell(row=9, column=2).fill = PatternFill("solid", fgColor=YELLOW)

r = 4 + len(rows) + 1
ws.cell(row=r, column=1, value="Event counts (live from Event_Log sheet)").font = bold
r += 1
hdr = ["Date", "Total events", "Driving Impact", "Hard Braking", "Hard Cornering", "Hard Acceleration", "Parking Impact"]
for j, h in enumerate(hdr, start=1):
    ws.cell(row=r, column=j, value=h)
style_header(ws, r, 7)
for k, d in enumerate([date(2026, 8, 28), date(2026, 8, 29), date(2026, 9, 4), date(2026, 9, 5)], start=1):
    rr = r + k
    ws.cell(row=rr, column=1, value=d).number_format = "ddd dd/mm/yyyy"
    ws.cell(row=rr, column=2, value=f"=COUNTIF(Event_Log!$A:$A,A{rr})")
    for j, typ in enumerate(["Driving Impact", "Hard Braking", "Hard Cornering", "Hard Acceleration", "Parking Impact"], start=3):
        ws.cell(row=rr, column=j, value=f'=COUNTIFS(Event_Log!$A:$A,A{rr},Event_Log!$C:$C,"{typ}")')
style_body(ws, r + 1, r + 4, 7)
ws.cell(row=r + 5, column=1,
        value="Counts fill in as you enter events on Event_Log. Fleeta showed ~40+ on 04/09 and 17 (07:03-08:34) on 05/09 at first look.").font = note_font
widths(ws, [34, 80, 26, 52, 16, 16, 16])
ws.freeze_panes = "A5"

# ---------------------------------------------------------------- Timeline
ws = wb.create_sheet("Timeline")
ws["A1"] = "Driver timeline from ServiceM8 (all times local AEST, hh:mm:ss). Camera-event column counts Event_Log rows falling in each leg's travel window."
ws["A1"].font = note_font
ws.merge_cells("A1:N1")
hdr = ["Date", "Leg", "To (job / suburb)", "Job #", "Booked", "On-the-way SMS", "Travel start (derived)",
       "Check-in", "Check-out", "Travel time", "On-site time", "Camera events in travel window",
       "Flag", "Note (source: ServiceM8 diary)"]
for j, h in enumerate(hdr, start=1):
    ws.cell(row=2, column=j, value=h)
style_header(ws, 2, len(hdr))

D4 = date(2026, 9, 4)
D5 = date(2026, 9, 5)
# (date, leg, to, job, booked, otw, travel_start, checkin, checkout, travel, onsite, flag, note)
legs = [
    (D4, 1, "Sunnybank - commercial food premises", "24795", time(6, 30), time(6, 7), time(6, 6, 47), time(6, 27), time(7, 8),
     td(0, 20, 13), td(0, 41, 2), "",
     "On-the-way SMS sent twice (06:07, 06:23). Booking note: someone on site from 06:30."),
    (D4, 2, "Fortitude Valley - German cockroach", "24316", time(8, 30), time(7, 49), time(7, 49, 30), time(8, 1), time(8, 55),
     td(0, 11, 42), td(0, 54, 3), "Gap",
     "07:08-07:49 (41 min) between check-out and travel start is unaccounted. 06:50 SMS to client 'technician delayed on a previous job' was sent while still at Sunnybank."),
    (D4, 3, "Thornlands - termite inspection", "24654", time(10, 0), time(8, 56), time(8, 55, 3), time(9, 42), time(10, 11),
     td(0, 46, 57), td(0, 29, 10), "Long leg",
     "Inner city to bayside, ~32 km, motorway. 48-min ETA quoted."),
    (D4, 4, "Murarrie - general pest", "24740", time(11, 30), time(10, 17), time(10, 12, 53), time(10, 51), time(11, 29),
     td(0, 38, 7), td(0, 37, 57), "", "Arrived 39 min before booking."),
    (D4, 5, "Morayfield - general pest", "24769", time(12, 30), time(11, 29), time(11, 29, 40), time(12, 29), time(13, 23),
     td(0, 59, 20), td(0, 54, 23), "Long leg",
     "Murarrie to Morayfield ~55 km via Gateway / Bruce Hwy. Longest motorway run of the day."),
    (D4, 6, "Narangba - 'drop cash' (staff home address)", "24244", time(13, 45), time(13, 25), time(13, 24, 42), time(13, 39), time(13, 40),
     td(0, 15, 18), td(0, 0, 33), "Non-customer stop",
     "33 seconds on site. Job card title 'ash p', description 'drop cash'. Confirm this errand was authorised and what was dropped."),
    (D4, 7, "Murrumba Downs - Termidor HE quote", "24756", time(15, 0), time(13, 40), time(13, 39, 27), time(13, 57), time(14, 25),
     td(0, 17, 33), td(0, 28, 16), "63 min early",
     "Booked 15:00, arrived 13:57. Quote $3,969 sent 14:20. Day compressed to finish early."),
    (D4, 8, "End of day - return (destination unknown)", "", None, None, time(14, 25), None, None, None, None, "Unmatched",
     "Match any Fleeta events after 14:25 to this leg. Trip History will show where the vehicle was parked overnight."),
    (D5, 1, "Chapel Hill - St James Estate (follow-up)", "24717", time(8, 0), None, time(5, 19, 30), time(7, 46, 34), time(7, 46, 38),
     td(2, 27, 4), td(0, 0, 4), "PRIORITY",
     "Travel start derived: check-out 07:46:38 minus total 02:27:08. Office to Chapel Hill is ~12 km / 25 min. 2h27m of travel is unexplained. "
     "07:28 call to client lasted 7 s (no answer inferred). Job not performed, still Work Order. Fleeta events 07:03-07:46 fall in this window."),
    (D5, 2, "South Brisbane - bed bug inspection/treatment", "24815", time(11, 0), time(7, 46), time(7, 46, 43), time(8, 11), time(8, 20, 33),
     td(0, 24, 39), td(0, 9, 16), "Quality",
     "07:46 called client to come early 'while nearby'. Booked as 1-hour service, on site 9 min, 1 mandatory checklist task skipped at check-out, invoiced $110. "
     "Travel from Chapel Hill 25 min is normal."),
    (D5, 3, "Parked (Parking Impact 08:34) then off sick", "", None, None, time(8, 20, 33), None, None, None, None, "Sick",
     "South Brisbane job is 1.8 km from office. Vehicle parked by 08:34. Jason called the 11:15 client at 11:07: 'technician went home sick'. "
     "Job 24831 reassigned to Jason and completed Sun 06/09 08:30."),
]
r = 3
for L in legs:
    d, leg, to, job, booked, otw, tstart, cin, cout, trav, onsite, flag, note = L
    ws.cell(row=r, column=1, value=d).number_format = "ddd dd/mm"
    ws.cell(row=r, column=2, value=leg)
    ws.cell(row=r, column=3, value=to)
    ws.cell(row=r, column=4, value=job)
    for col, val in [(5, booked), (6, otw), (7, tstart), (8, cin), (9, cout)]:
        c = ws.cell(row=r, column=col, value=val)
        c.number_format = "hh:mm:ss"
    for col, val in [(10, trav), (11, onsite)]:
        c = ws.cell(row=r, column=col, value=val)
        c.number_format = "[h]:mm:ss"
    # camera events between travel start and check-in (or end of day if no check-in)
    end_ref = f"H{r}" if cin else "TIME(23,59,59)"
    ws.cell(row=r, column=12,
            value=f'=COUNTIFS(Event_Log!$A:$A,A{r},Event_Log!$B:$B,">="&G{r},Event_Log!$B:$B,"<="&{end_ref})')
    ws.cell(row=r, column=13, value=flag)
    ws.cell(row=r, column=14, value=note)
    r += 1
last = r - 1
style_body(ws, 3, last, len(hdr))
for rr in range(3, last + 1):
    f = ws.cell(row=rr, column=13).value
    if f == "PRIORITY":
        for c in range(1, len(hdr) + 1):
            ws.cell(row=rr, column=c).fill = PatternFill("solid", fgColor=RED)
    elif f in ("Non-customer stop", "Quality", "Gap"):
        for c in range(1, len(hdr) + 1):
            ws.cell(row=rr, column=c).fill = PatternFill("solid", fgColor=YELLOW)
r += 1
ws.cell(row=r, column=1, value="Totals 04/09").font = bold
ws.cell(row=r, column=10, value=f"=SUMIFS(J3:J{last},A3:A{last},DATE(2026,9,4))").number_format = "[h]:mm:ss"
ws.cell(row=r, column=11, value=f"=SUMIFS(K3:K{last},A3:A{last},DATE(2026,9,4))").number_format = "[h]:mm:ss"
ws.cell(row=r, column=12, value=f"=SUMIFS(L3:L{last},A3:A{last},DATE(2026,9,4))")
r += 1
ws.cell(row=r, column=1, value="Totals 05/09").font = bold
ws.cell(row=r, column=10, value=f"=SUMIFS(J3:J{last},A3:A{last},DATE(2026,9,5))").number_format = "[h]:mm:ss"
ws.cell(row=r, column=11, value=f"=SUMIFS(K3:K{last},A3:A{last},DATE(2026,9,5))").number_format = "[h]:mm:ss"
ws.cell(row=r, column=12, value=f"=SUMIFS(L3:L{last},A3:A{last},DATE(2026,9,5))")
style_body(ws, r - 1, r, len(hdr), fill=GREY)
r += 2
ws.cell(row=r, column=1, value="Context: Fri 28/08 and Sat 29/08 - no jobs allocated to Javiar (ServiceM8). Fleeta shows zero events both days. Consistent with vehicle not driven.").font = note_font
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
r += 1
ws.cell(row=r, column=1, value="Travel start is derived (check-out minus total time on the ServiceM8 check-out stamp). ServiceM8 'travel' begins when the tech taps On The Way, so it is a self-reported window, not GPS.").font = note_font
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
widths(ws, [11, 5, 38, 8, 9, 11, 12, 10, 10, 10, 10, 12, 16, 90])
ws.freeze_panes = "C3"

# ---------------------------------------------------------------- Event_Log
ws = wb.create_sheet("Event_Log")
ws["A1"] = ("Enter one row per Fleeta Cloud event. Yellow cells are inputs. Date as dd/mm/yyyy, Time as hh:mm:ss local (confirm Fleeta is showing AEST). "
            "Row 4 is the one event already confirmed from the Fleeta calendar view.")
ws["A1"].font = note_font
ws.merge_cells("A1:M1")
hdr = ["Date", "Time", "Event type", "Speed (km/h)", "Location / suburb", "GPS (lat, long)", "Clip file name",
       "Downloaded (Y/N)", "SHA-256", "Matched Timeline leg", "Audio track (Y/N/Unk)", "Reviewed by", "Notes"]
for j, h in enumerate(hdr, start=1):
    ws.cell(row=3, column=j, value=h)
style_header(ws, 3, len(hdr))
ws.cell(row=4, column=1, value=date(2026, 9, 5)).number_format = "dd/mm/yyyy"
ws.cell(row=4, column=2, value=time(8, 34)).number_format = "hh:mm:ss"
ws.cell(row=4, column=3, value="Parking Impact")
ws.cell(row=4, column=8, value="N")
ws.cell(row=4, column=10, value="05/09 leg 3")
ws.cell(row=4, column=11, value="Unk")
ws.cell(row=4, column=13, value="Seen in Fleeta calendar 06/09. Vehicle parked ~14 min after South Brisbane check-out. Clip not yet downloaded.")
style_body(ws, 4, 120, len(hdr), fill=YELLOW)
for rr in range(4, 121):
    ws.cell(row=rr, column=1).number_format = "dd/mm/yyyy"
    ws.cell(row=rr, column=2).number_format = "hh:mm:ss"
dv = DataValidation(type="list",
                    formula1='"Driving Impact,Hard Braking,Hard Cornering,Hard Acceleration,Parking Impact,Overspeed,Manual,Other"',
                    allow_blank=True)
ws.add_data_validation(dv)
dv.add("C4:C120")
dv2 = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
ws.add_data_validation(dv2)
dv2.add("H4:H120")
dv3 = DataValidation(type="list", formula1='"Y,N,Unk"', allow_blank=True)
ws.add_data_validation(dv3)
dv3.add("K4:K120")
widths(ws, [12, 10, 18, 10, 24, 22, 34, 11, 66, 16, 12, 14, 50])
ws.freeze_panes = "C4"

# ---------------------------------------------------------------- Clip_Register
ws = wb.create_sheet("Clip_Register")
ws["A1"] = ("Chain of custody. One row per file (video clip, SD image, trip export, ServiceM8 diary PDF, call recording). "
            "Hash with: certutil -hashfile <file> SHA256 (Windows) or shasum -a 256 <file> (Mac). Never edit originals; work on copies.")
ws["A1"].font = note_font
ws.merge_cells("A1:K1")
hdr = ["Ref", "Source", "Original file name", "Event / content date-time", "Collected by", "Collected at (date-time)",
       "Stored at (path / drive)", "SHA-256", "Copies given to (who, when)", "Retention hold until", "Notes"]
for j, h in enumerate(hdr, start=3 - 2):
    ws.cell(row=3, column=j, value=h)
style_header(ws, 3, len(hdr))
example = ["EV-001", "Fleeta Cloud", "20260905_083400_PF.mp4", "05/09/2026 08:34", "J. Cartledge", "07/09/2026 09:00",
           "SAB-HR\\Vehicle-Review-2026-09\\clips\\", "(paste hash)", "-", "05/09/2027", "EXAMPLE ROW - overwrite with the first real clip"]
for j, v in enumerate(example, start=1):
    ws.cell(row=4, column=j, value=v)
style_body(ws, 4, 60, len(hdr), fill=YELLOW)
ws.cell(row=4, column=11).font = Font(name=FONT, size=10, italic=True, color="C00000")
dv4 = DataValidation(type="list",
                     formula1='"Fleeta Cloud,SD card image,SD card clip,Fleeta Trip History export,ServiceM8 diary PDF,Phone call recording,Witness statement,Other"',
                     allow_blank=True)
ws.add_data_validation(dv4)
dv4.add("B4:B60")
widths(ws, [9, 20, 30, 20, 14, 20, 40, 66, 26, 16, 40])
ws.freeze_panes = "B4"

# ---------------------------------------------------------------- Gaps
ws = wb.create_sheet("Gaps")
hdr = ["#", "Gap / question", "Why it matters", "Owner", "Due", "Status"]
for j, h in enumerate(hdr, start=1):
    ws.cell(row=1, column=j, value=h)
style_header(ws, 1, len(hdr))
gaps = [
    ("Where was the vehicle 05:19-07:46 on Sat 05/09?", "2h27m of self-reported travel for a 25-min drive, then 4 s on site. Fleeta Trip History answers this.", "Browser session", date(2026, 9, 7)),
    ("Who is the passenger, what is their role, have they complained?", "Decides the process: WHS driving conduct vs harassment complaint (positive duty). Complainant welfare comes first.", "Jason", date(2026, 9, 7)),
    ("Is cabin audio ON in the camera config?", "If ON, recording may breach Invasion of Privacy Act 1971 (Qld) s43. Do not listen or share before legal advice.", "Jason", date(2026, 9, 7)),
    ("Has Javiar signed a vehicle use / camera monitoring policy?", "Notice of monitoring underpins fair use of the footage and any disciplinary reliance on it.", "Ashleigh (HR file)", date(2026, 9, 7)),
    ("Fleeta Cloud event retention period on the current plan", "Sets the hard deadline for downloads. Assume short; download this week regardless.", "Browser session", date(2026, 9, 7)),
    ("Pull and image the SD card before September overwrites", "Card holds continuous footage between triggers for 04-05/09. Cloud does not.", "Jason", date(2026, 9, 7)),
    ("Was the Narangba 'drop cash' stop (04/09 13:39) authorised, and what was dropped?", "Non-customer stop at a staff home on a job card. Needs a one-line explanation on file.", "Jason / Ashleigh", date(2026, 9, 8)),
    ("Was Javiar rostered on Sat 05/09 and was he told the 08:00 Chapel Hill scope?", "Booked 04/09 12:29 by Ashleigh. Confirms expectation vs outcome (job not done).", "Ashleigh", date(2026, 9, 8)),
    ("Who spoke to Javiar about going home sick on 05/09, at what time?", "Fixes the end of the driving window and whether the vehicle moved again that day.", "Jason", date(2026, 9, 8)),
    ("Is the 'conduct issue' driving behaviour, treatment of the passenger, unauthorised use, or all three?", "Each allegation must be put separately and specifically in writing before any meeting.", "Jason", date(2026, 9, 8)),
    ("Who runs the meeting given the family relationship?", "Perceived bias. An external HR adviser or independent manager should conduct and minute it.", "Jason", date(2026, 9, 9)),
    ("Confirm Fleeta timestamps are AEST and match ServiceM8", "A timezone offset would misalign every event to the wrong leg.", "Browser session", date(2026, 9, 7)),
    ("Bed bug job 24815: was a 9-minute inspection with a skipped mandatory task acceptable?", "Separate service-quality/performance matter. Keep it out of the conduct allegations unless relevant to the day.", "Jason", date(2026, 9, 9)),
]
for i, g in enumerate(gaps, start=2):
    ws.cell(row=i, column=1, value=i - 1)
    ws.cell(row=i, column=2, value=g[0])
    ws.cell(row=i, column=3, value=g[1])
    ws.cell(row=i, column=4, value=g[2])
    ws.cell(row=i, column=5, value=g[3]).number_format = "ddd dd/mm"
    ws.cell(row=i, column=6, value="Open")
style_body(ws, 2, 1 + len(gaps), len(hdr))
dv5 = DataValidation(type="list", formula1='"Open,In progress,Closed,Blocked"', allow_blank=True)
ws.add_data_validation(dv5)
dv5.add(f"F2:F{1 + len(gaps)}")
widths(ws, [4, 62, 78, 18, 11, 12])
ws.freeze_panes = "A2"

# ---------------------------------------------------------------- Process
ws = wb.create_sheet("Process")
hdr = ["Step", "Action", "Owner", "Due", "Done (Y/N)", "Evidence / note"]
for j, h in enumerate(hdr, start=1):
    ws.cell(row=1, column=j, value=h)
style_header(ws, 1, len(hdr))
steps = [
    ("Preserve", "Pull SD card, image it (full disk image, not file copy), hash, bag and label. Fit new card.", "Jason", date(2026, 9, 7)),
    ("Preserve", "Fleeta: export Trip History 04/09 and 05/09 (and 28-29/08 to close those out). Screenshot the calendar view.", "Browser session", date(2026, 9, 7)),
    ("Preserve", "Fleeta: enumerate every event 04/09 and 05/09 into Event_Log. Download all clips, hash each, log in Clip_Register.", "Browser session", date(2026, 9, 8)),
    ("Preserve", "ServiceM8: export job diaries 24717, 24815, 24831 (05/09) and 24795, 24316, 24654, 24740, 24769, 24244, 24756 (04/09) to PDF.", "Ashleigh", date(2026, 9, 8)),
    ("Preserve", "Phone system: secure recordings of calls 05/09 07:28 (Gerry) and 07:46 (Charlize) if the system stores audio.", "Ashleigh", date(2026, 9, 8)),
    ("Legal", "Confirm camera audio setting. If ON: seek legal advice before any use; consider disabling fleet-wide with written notice to staff.", "Jason", date(2026, 9, 8)),
    ("Legal", "Check HR file for signed vehicle use and monitoring policy. If none, note it and issue one going forward.", "Ashleigh", date(2026, 9, 8)),
    ("Scope", "Speak with the passenger (welfare first). Record their account. Decide if a formal complaint exists.", "Jason (or independent)", date(2026, 9, 8)),
    ("Scope", "Write each allegation as a specific, dated statement (what, when, where, which clip). No conclusions.", "Independent adviser", date(2026, 9, 9)),
    ("Fairness", "Issue written allegations + meeting invite with 24-48 h notice, right to a support person, copies of the clips relied on.", "Independent adviser", date(2026, 9, 9)),
    ("Fairness", "Hold the meeting. Put each allegation, hear the response, take minutes, adjourn before deciding.", "Independent adviser", date(2026, 9, 11)),
    ("Decide", "Decide on the evidence and the response. Options: coaching + written warning, final warning, or termination if serious misconduct is made out.", "Jason on advice", date(2026, 9, 14)),
    ("Systemise", "Fleet safety: telematics scoring review weekly, harsh-event threshold alerts to Jason, driver coaching SOP, passenger authorisation rule.", "Jason", date(2026, 9, 18)),
]
for i, s in enumerate(steps, start=2):
    ws.cell(row=i, column=1, value=s[0])
    ws.cell(row=i, column=2, value=s[1])
    ws.cell(row=i, column=3, value=s[2])
    ws.cell(row=i, column=4, value=s[3]).number_format = "ddd dd/mm"
    ws.cell(row=i, column=5, value="N")
style_body(ws, 2, 1 + len(steps), len(hdr))
dv6 = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
ws.add_data_validation(dv6)
dv6.add(f"E2:E{1 + len(steps)}")
widths(ws, [11, 96, 22, 11, 10, 40])
ws.freeze_panes = "A2"

for sheet in wb.worksheets:
    for row in sheet.iter_rows():
        for cell in row:
            if cell.font.name != FONT:
                cell.font = Font(name=FONT, size=cell.font.size or 10, bold=cell.font.bold,
                                 italic=cell.font.italic, color=cell.font.color)

wb.save(OUT)
print("wrote", OUT)
