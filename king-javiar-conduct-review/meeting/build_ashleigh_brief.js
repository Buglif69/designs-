const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType, LevelFormat, convertInchesToTwip,
} = require("docx");

const FONT = "Arial";
const NAVY = "1F3A5F";
const GREY = "F2F2F2";
const AMBER = "FFF2CC";

const p = (text, o = {}) => new Paragraph({
  spacing: { after: o.after === undefined ? 150 : o.after, line: 276 },
  alignment: o.align,
  children: [new TextRun({ text, font: FONT, size: o.size || 20, bold: o.bold, italics: o.italics, color: o.color })],
});

const runs = (parts, o = {}) => new Paragraph({
  spacing: { after: o.after === undefined ? 150 : o.after, line: 276 },
  children: parts.map(x => new TextRun({ text: x.t, font: FONT, size: o.size || 20, bold: x.b, italics: x.i, color: x.c })),
});

const h1 = (text) => new Paragraph({
  spacing: { before: 300, after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
  children: [new TextRun({ text, font: FONT, size: 24, bold: true, color: NAVY })],
});

const h2 = (text) => new Paragraph({
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, font: FONT, size: 21, bold: true })],
});

const bullet = (text) => new Paragraph({
  numbering: { reference: "dots", level: 0 },
  spacing: { after: 100, line: 276 },
  children: [new TextRun({ text, font: FONT, size: 20 })],
});

const cell = (text, o = {}) => new TableCell({
  width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill } : undefined,
  margins: { top: 70, bottom: 70, left: 100, right: 100 },
  children: (Array.isArray(text) ? text : [text]).map(t => new Paragraph({
    spacing: { after: 0, line: 264 },
    children: [new TextRun({ text: t, font: FONT, size: o.size || 19, bold: o.bold, color: o.color })],
  })),
});

const table = (widths, header, rows, opts = {}) => new Table({
  columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9" },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9" },
  },
  rows: [
    new TableRow({
      tableHeader: true,
      children: header.map((t, i) => cell(t, { w: widths[i], fill: NAVY, bold: true, color: "FFFFFF" })),
    }),
    ...rows.map(r => new TableRow({
      children: r.map((t, i) => cell(t, { w: widths[i], fill: opts.fill })),
    })),
  ],
});

const callout = (title, body, fill) => new Table({
  columnWidths: [9300],
  width: { size: 9300, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
    left: { style: BorderStyle.SINGLE, size: 18, color: fill === AMBER ? "BF8F00" : NAVY },
    right: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
  },
  rows: [new TableRow({
    children: [new TableCell({
      width: { size: 9300, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill },
      margins: { top: 130, bottom: 130, left: 150, right: 150 },
      children: [
        new Paragraph({ spacing: { after: 70 }, children: [new TextRun({ text: title, font: FONT, size: 21, bold: true })] }),
        ...body.map(t => new Paragraph({ spacing: { after: 60, line: 276 }, children: [new TextRun({ text: t, font: FONT, size: 20 })] })),
      ],
    })],
  })],
});

const spacer = (h = 200) => new Paragraph({ spacing: { after: h }, children: [new TextRun({ text: "", font: FONT, size: 2 })] });

const FRI = "24795, 24316, 24654, 24740, 24769, 24244, 24756";
const SAT = "24717, 24815, 24831";

const doc = new Document({
  creator: "10X Pty Ltd t/a Slug-A-Bug Pest Control",
  title: "Evidence Collection Brief",
  numbering: {
    config: [{ reference: "dots", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: convertInchesToTwip(0.28), hanging: convertInchesToTwip(0.17) } } } }] }],
  },
  sections: [{
    properties: { page: { margin: { top: 900, bottom: 900, left: 1000, right: 1000 } } },
    children: [
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: "PRIVATE AND CONFIDENTIAL", font: FONT, size: 18, bold: true, color: "C00000" })],
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: "EVIDENCE COLLECTION BRIEF", font: FONT, size: 30, bold: true, color: NAVY })],
      }),
      p("Workplace review — company vehicle BUGL1FE and related work records", { size: 21, color: "595959", after: 200 }),

      table([1700, 7600], ["Field", "Detail"], [
        ["To", "Ashleigh Pring, Executive Assistant"],
        ["From", "Jason Cartledge, Director"],
        ["Date issued", "____ / ____ / 2026"],
        ["Everything due by", "____ / ____ / 2026, close of business"],
        ["Save everything to", "________________________________________________"],
        ["Questions to", "Jason directly — not to any other staff member"],
      ]),

      spacer(240),

      callout("Read this first", [
        "This is a formal workplace review. Everything in this brief is confidential and must not be discussed with any other staff member, including the person it concerns.",
        "Do not contact Javiar about any of it. If he asks you anything, tell him to speak to Jason and let Jason know.",
        "Collect what is there. Do not summarise, interpret, or leave anything out because it looks unhelpful — material that cuts both ways still gets collected.",
        "If something on this list does not exist, write “not found” against it. That is a valid and useful answer.",
      ], GREY),

      h1("1.  ServiceM8 exports"),

      p("These come from the ServiceM8 web reports, not the app. Export to PDF or Excel and save with the naming convention in section 5."),

      table([3400, 4300, 1600], ["What", "Where / how", "Done"], [
        ["Job Activity report — Javiar", "Reports → Staff → Job Activity. Date range 8 Aug – 5 Sep 2026.", ""],
        ["Job Activity report — Noah", "Same report, same dates. Needed as a comparison.", ""],
        ["Timesheets — Javiar", "Reports → Staff → Timesheets. 8 Aug – 5 Sep 2026.", ""],
        ["Timesheets — Noah", "Same report, same dates.", ""],
        ["Staff location / GPS history", "Staff location map. Export or screenshot 4 and 5 Sep in full.", ""],
        ["Job cards — Friday 4 Sep", "Print to PDF, full diary including timestamps: " + FRI, ""],
        ["Job cards — Saturday 5 Sep", "Print to PDF, full diary including timestamps: " + SAT, ""],
        ["Open / incomplete jobs", "Any job allocated to Javiar still sitting as Work Order or In Progress, 1 Jul – 5 Sep.", ""],
      ]),

      spacer(160),
      runs([{ t: "Note on job 24717 (St James Estate, Chapel Hill): ", b: true },
             { t: "this one matters most. Export the complete diary showing every timestamp, including the check-out on Sat 5 Sep and the earlier attendance on Wed 3 Sep." }]),

      h1("2.  Fleeta Cloud — camera and trip data"),

      p("Vehicle: BUGL1FE (2023 Ford Ranger, camera name “king Javiar”). Cloud storage overwrites on a cycle, so this is time-critical."),

      table([3400, 4300, 1600], ["What", "Where / how", "Done"], [
        ["Confirm the retention period", "Account or plan settings. Write the number of days on this sheet — it sets the deadline.", ""],
        ["Confirm the timezone", "Open one event and check it matches a known ServiceM8 time. Report the offset if it is not AEST.", ""],
        ["Event list — Fri 4 Sep", "Export or screenshot every event with time, type, speed and location.", ""],
        ["Event list — Sat 5 Sep", "Same. The cluster between 07:03 and 08:34 is the priority.", ""],
        ["Trip History — Sat 5 Sep", "05:00 to 08:40. Export the route and note every stop with address and dwell time.", ""],
        ["Trip History — Fri 4 Sep", "Full day, same detail.", ""],
        ["Trip History — last 4 weeks", "8 Aug – 5 Sep. Note start and finish locations against job addresses.", ""],
        ["Download all Impact clips", "Both days. Driving Impact first, then Parking Impact.", ""],
        ["Download the Sat 07:03–08:34 clips", "All of them, in order.", ""],
      ]),

      spacer(200),

      callout("STOP — check this before downloading any clip", [
        "Open one clip and check whether it has an audio track.",
        "If there is cabin audio: stop immediately. Do not play it, do not download further clips, do not send it to anyone. Note “audio present” and tell Jason the same day.",
        "Recording a private conversation you are not part of can breach the Invasion of Privacy Act 1971 (Qld). Until we have legal advice, that audio is a risk to the company, not evidence.",
        "If there is no audio track, carry on with the list above.",
      ], AMBER),

      h1("3.  SD card"),

      p("The card holds continuous footage between camera triggers, which the cloud does not. August has already been overwritten and September is cycling."),

      table([3400, 4300, 1600], ["What", "Where / how", "Done"], [
        ["Remove the card", "Before the vehicle is next driven. Jason to arrange access.", ""],
        ["Image the card", "Full disk image, not a file copy. IT or Jason to do this.", ""],
        ["Label and store", "Bag it, label with date, time and who removed it. Store locked.", ""],
        ["Fit a replacement", "New formatted card so the vehicle keeps recording.", ""],
      ]),

      h1("4.  HR file and communications"),

      table([3400, 4300, 1600], ["What", "Where / how", "Done"], [
        ["Signed employment contract", "Drive: “2025_Contract_Full-time_Salary_Javier Cartledge-1 - signed.pdf”. Save a copy to the review folder.", ""],
        ["Position description", "Whatever is on file for Termite Inspector & Pest Technician.", ""],
        ["Vehicle or camera policy", "Check whether a signed vehicle use / monitoring policy exists. If none, write “not found”.", ""],
        ["Any prior warnings or file notes", "Whole employment period. If none exist, write “none on file”.", ""],
        ["Sick leave email", "From Javiar_7@hotmail.com to info@, Sun 6 Sep 5:58pm, “Sick leave Javiar CARTLEDGE”. Print to PDF with full headers.", ""],
        ["Medical certificate", "Whatever was provided, if any.", ""],
        ["Leave records", "Personal and carer’s leave taken and balance remaining, last 6 months.", ""],
        ["Office emails to and from Javiar", "info@ mailbox, 1 Aug – 6 Sep. Export the thread list and anything about attendance, jobs or the vehicle.", ""],
        ["The BDM email", "Kristy’s performance email about Javiar is not in info@. Ask her to forward it to info@, then save it.", ""],
      ]),

      h1("5.  How to save and name everything"),

      p("One folder for the whole review. Inside it, one subfolder per section above."),

      table([2600, 6700], ["Rule", "Detail"], [
        ["File names", "YYYY-MM-DD_Source_What.pdf — e.g. 2026-09-05_ServiceM8_JobCard-24717.pdf"],
        ["Never edit originals", "Save the export as it comes out. Work on copies if anything needs marking up."],
        ["Video files", "Keep the original file name from Fleeta as well, in the notes column of the register."],
        ["Register", "Log every file in the Clip_Register sheet of the evidence workbook: what it is, where it came from, who collected it, when."],
        ["Hash the video files", "Windows: certutil -hashfile \"<file>\" SHA256. Paste the result into the register. This proves the file has not been altered."],
        ["Access", "Restrict the folder to Jason and yourself only."],
      ]),

      h1("6.  Housekeeping item — do this first"),

      callout("The “Daily Communication” document", [
        "That document currently holds every business password in plain text — ATO, QBCC, WorkCover, BP fleet, Amex, Facebook, Twilio and the staff email logins including javiar@slugabug.com.au.",
        "Move all credentials into a password manager and delete that section from the document today.",
        "Then change the staff email passwords listed in it.",
        "This is not about anyone in particular. While a formal review is running, shared plain-text credentials mean we cannot prove who accessed what — and that works against the company.",
      ], AMBER),

      h1("7.  What not to do"),

      bullet("Do not discuss this with any staff member other than Jason."),
      bullet("Do not contact Javiar about any of it, and do not reply if he raises it — refer him to Jason."),
      bullet("Do not delete, edit or tidy up any record, message or file, including anything that looks unhelpful to the company."),
      bullet("Do not listen to, copy or forward any cabin audio. See the stop notice in section 2."),
      bullet("Do not send any of this material outside the company without Jason’s approval."),

      h1("8.  Progress"),

      p("Fill this in as you go and send it back with the material."),

      table([2200, 1700, 1700, 3700], ["Section", "Started", "Completed", "Notes / anything not found"], [
        ["1. ServiceM8", "", "", ""],
        ["2. Fleeta", "", "", ""],
        ["3. SD card", "", "", ""],
        ["4. HR file", "", "", ""],
        ["5. Filing", "", "", ""],
        ["6. Passwords", "", "", ""],
      ]),

      spacer(300),
      p("Thanks Ashleigh. If anything on this list is unclear or you hit a wall, call me rather than guessing.", { italics: true }),
      spacer(200),
      p("________________________________________", { after: 60 }),
      p("Jason Cartledge", { bold: true, after: 20 }),
      p("Director, 10X Pty Ltd t/a Slug-A-Bug Pest Control", { size: 19 }),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync("/home/user/designs-/king-javiar-conduct-review/meeting/Evidence-Collection-Brief-Ashleigh.docx", b);
  console.log("written");
});
