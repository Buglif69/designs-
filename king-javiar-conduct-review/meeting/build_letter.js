const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType,
  LevelFormat, convertInchesToTwip,
} = require("docx");

const FONT = "Arial";
const NAVY = "1F3A5F";

const p = (text, o = {}) => new Paragraph({
  spacing: { after: o.after === undefined ? 160 : o.after, line: 276 },
  alignment: o.align,
  indent: o.indent,
  children: [new TextRun({ text, font: FONT, size: o.size || 21, bold: o.bold, italics: o.italics, color: o.color })],
});

const runs = (parts, o = {}) => new Paragraph({
  spacing: { after: o.after === undefined ? 160 : o.after, line: 276 },
  indent: o.indent,
  children: parts.map(x => new TextRun({
    text: x.t, font: FONT, size: o.size || 21, bold: x.b, italics: x.i, color: x.c,
  })),
});

const h = (text) => new Paragraph({
  spacing: { before: 260, after: 130 },
  children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: NAVY })],
});

const rule = () => new Paragraph({
  spacing: { before: 80, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "BFBFBF", space: 6 } },
  children: [new TextRun({ text: "", font: FONT, size: 2 })],
});

const bullet = (text) => new Paragraph({
  numbering: { reference: "dots", level: 0 },
  spacing: { after: 110, line: 276 },
  children: [new TextRun({ text, font: FONT, size: 21 })],
});

const num = (text) => new Paragraph({
  numbering: { reference: "nums", level: 0 },
  spacing: { after: 130, line: 276 },
  children: [new TextRun({ text, font: FONT, size: 21 })],
});

// two-column detail block
const detail = (rows) => new Table({
  columnWidths: [2200, 7100],
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
  },
  rows: rows.map(([k, v]) => new TableRow({
    children: [
      new TableCell({
        width: { size: 2200, type: WidthType.DXA },
        margins: { top: 40, bottom: 40, right: 120 },
        children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: k, font: FONT, size: 21, bold: true })] })],
      }),
      new TableCell({
        width: { size: 7100, type: WidthType.DXA },
        margins: { top: 40, bottom: 40 },
        children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: v, font: FONT, size: 21 })] })],
      }),
    ],
  })),
});

const doc = new Document({
  creator: "10X Pty Ltd t/a Slug-A-Bug Pest Control",
  title: "Notice of Formal Meeting",
  numbering: {
    config: [
      { reference: "dots", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: convertInchesToTwip(0.3), hanging: convertInchesToTwip(0.18) } } } }] },
      { reference: "nums", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: convertInchesToTwip(0.32), hanging: convertInchesToTwip(0.32) } } } }] },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
    children: [
      p("[ PLACE SLUG-A-BUG LETTERHEAD HERE ]", { align: AlignmentType.CENTER, size: 17, color: "9E9E9E", italics: true, after: 260 }),

      new Paragraph({
        spacing: { after: 60 },
        children: [new TextRun({ text: "PRIVATE AND CONFIDENTIAL", font: FONT, size: 19, bold: true, color: NAVY })],
      }),
      p("Addressee only", { size: 18, color: "595959", after: 240 }),

      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun({ text: "NOTICE OF FORMAL MEETING", font: FONT, size: 28, bold: true, color: NAVY })],
      }),

      detail([
        ["To", "Javiar Cartledge"],
        ["From", "Jason Cartledge, Director, 10X Pty Ltd t/a Slug-A-Bug Pest Control"],
        ["Date", "____ / ____ / 2026"],
        ["Meeting date", "____ / ____ / 2026  at  ______ am / pm"],
        ["Location", "________________________________________"],
        ["Chairing", "________________________________________"],
        ["Also present", "________________________________________  (note-taker)"],
      ]),

      rule(),

      p("Dear Javiar,"),

      p("I am writing to ask you to attend a formal meeting to discuss a number of concerns about your work. This letter sets out what will be discussed so that you can prepare."),

      runs([
        { t: "No decision has been made about any of these matters. ", b: true },
        { t: "The purpose of the meeting is to put these concerns to you and to hear your response before anything is decided." },
      ]),

      h("Matters to be discussed"),

      p("1.  Job attendance on Saturday 5 September 2026", { bold: true, after: 90 }),
      p("Company records show that on Saturday 5 September you commenced travel at approximately 5:19 am for a job scheduled at 8:00 am at St James Estate, Chapel Hill (job 24717). The records show you were on site for approximately four seconds, checked out at 7:46 am, and that the job was not performed and remains open. I want to understand what happened during that period and why the job was not completed.", { after: 200 }),

      p("2.  Following a direction to contact the office at the end of each day", { bold: true, after: 90 }),
      p("Last week I directed you to telephone the office at the completion of your final job each day before leaving the area. I want to discuss whether that direction has been followed and, if not, why.", { after: 200 }),

      p("3.  Job records and scheduling in ServiceM8", { bold: true, after: 90 }),
      p("I want to discuss whether jobs are being updated in ServiceM8 to reflect actual attendance times, and the effect this has on the company’s ability to schedule work, invoice accurately and verify hours worked.", { after: 200 }),

      p("4.  Use of the company vehicle and passengers", { bold: true, after: 90 }),
      p("The company vehicle (BUGL1FE) is insured for use by employees engaged in company business. I want to discuss who has travelled in the vehicle during working hours, and to confirm your understanding of the company’s requirements about passengers and vehicle use.", { after: 200 }),

      p("5.  Driving conduct", { bold: true, after: 90 }),
      p("The vehicle’s camera system has recorded a number of harsh braking, harsh cornering, harsh acceleration and impact events. I want to discuss these with you from a work health and safety perspective.", { after: 200 }),

      p("6.  Communication and availability during working hours", { bold: true, after: 90 }),
      p("I want to discuss the arrangements for contacting you during working hours, and what has occurred on occasions where the office has been unable to reach you.", { after: 200 }),

      p("7.  Working relationships", { bold: true, after: 90 }),
      p("I want to discuss your working relationships with other members of the team.", { after: 220 }),

      h("Your rights"),

      bullet("You may bring a support person of your choice to the meeting. Please let us know beforehand who that will be."),
      bullet("You will be given a full opportunity to respond to each matter, and to give any explanation or context you wish."),
      bullet("You may ask to see any record or document relied on in relation to any matter above. Tell us before the meeting and it will be provided to you."),
      bullet("Notes will be taken. You will be provided with a copy."),
      bullet("If you need more time to prepare, or would prefer a different date, contact the office and we will accommodate a reasonable request."),

      h("Possible outcomes"),

      p("Depending on what is established at the meeting and your response, possible outcomes include that no further action is taken, that additional training or supervision is put in place, or that disciplinary action is taken. Disciplinary action may include a written warning or, if a matter is found to be serious, termination of your employment. No decision will be made until after you have had the opportunity to respond."),

      h("Two things I want to be clear about"),

      runs([
        { t: "This meeting is not about you being unwell, and it is not about the leave you took to care for a family member. ", b: true },
        { t: "You are entitled to personal and carer’s leave and no action is being taken against you for using it. Where absence appears in this letter, the concern is about notification and about work that was scheduled, not about your entitlement to leave." },
      ]),

      p("If there is anything going on — health, family, or anything else — that is affecting your work, I would rather know about it. Raising it will not be held against you and it may change how these matters are dealt with."),

      p("If you have any questions about this letter, please contact the office."),

      p("Yours sincerely,", { after: 420 }),

      p("________________________________________", { after: 60 }),
      p("Jason Cartledge", { bold: true, after: 20 }),
      p("Director, 10X Pty Ltd t/a Slug-A-Bug Pest Control", { size: 19, after: 340 }),

      rule(),

      p("Acknowledgement of receipt", { bold: true, after: 90 }),
      p("Signing below confirms only that you have received this letter. It does not indicate agreement with anything in it.", { size: 19, italics: true, after: 200 }),

      detail([
        ["Signed", "________________________________________"],
        ["Name", "________________________________________"],
        ["Date", "______ / ______ / 2026"],
      ]),

      p("If the employee declines to sign, record: “Copy provided and receipt declined”, with the date, and have the witness initial it.", { size: 17, italics: true, color: "808080", after: 0 }),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync("/home/user/designs-/king-javiar-conduct-review/meeting/Notice-of-Formal-Meeting-Javiar-Cartledge.docx", b);
  console.log("written");
});
