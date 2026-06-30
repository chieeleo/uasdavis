// Builds the Eye-Health exploratory report (.docx) with 3 static charts + 300-word Data Story.
// Run: node build_report.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, LevelFormat
} = require("docx");

const HERE = __dirname;
const A = path.join(HERE, "report_assets");
const img = (f) => fs.readFileSync(path.join(A, f));

const INK = "0F2A30", TEAL = "0D9488", TEALD = "115E59", CORAL = "EF6F53", MUTE = "52696F";
const LINE = "DDE8E6", HEADFILL = "E2F1EE";

/* ---------- Data Story (target ~300 words) ---------- */
const story = {
  intro:
    "Across 10,000 complete respondents (11 attributes, zero missing or duplicate rows), the average " +
    "eye-health score sits at a moderate 65.3 out of 100, with most people clustered between the mid-50s " +
    "and mid-70s. That headline average, however, hides who is at risk and why.",
  audienceLead: "Who this is for. ",
  audience:
    "The dashboard is built for digital-wellbeing and public-health communicators, and for the everyday " +
    "screen user who wants to know which habits actually matter. It is exploratory, not clinical.",
  insightLead: "The single most important insight. ",
  insight:
    "Eye-health is shaped by two kinds of forces. The strongest correlates — prescription level " +
    "(r = -0.56) and age (r = -0.49) — are largely outside a person’s control: mean scores fall " +
    "from 73.6 in the under-20s to 56.5 in the over-70s. The key takeaway is the largest modifiable factor: " +
    "daily screen time (r = -0.49). Heavy-screen users score roughly 20 points lower than light users, while " +
    "outdoor light (+0.23) and exercise (+0.24) pull scores back up. Age sets the baseline, but daily " +
    "behaviour decides where you land on it — and screen time is the lever a user can actually pull.",
  biasLead: "How bias was addressed. ",
  bias:
    "Because this is observational, cross-sectional data, every relationship is described as an association, " +
    "never a cause. The dual-axis chart was deliberately re-based onto age, where both series genuinely move, " +
    "rather than exaggerating a near-flat screen-time line through a stretched axis. Age stands in for a " +
    "timeline only because the data carries no date field, and this is labelled openly. Distributions are " +
    "shown in full, not just averages; group counts are always displayed; the ordinal glasses scale uses a " +
    "single-hue sequence (not a rainbow) so order is not mistaken for category; and the 300 capped " +
    "screen-time records are retained and disclosed."
};
const storyWords = Object.values(story).join(" ").trim().split(/\s+/).length;
console.log("Data Story word count:", storyWords);

/* ---------- helpers ---------- */
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const P  = (runs, opts={}) => new Paragraph({ spacing: { after: 140, line: 300 }, children: runs, ...opts });
const caption = (t) => new Paragraph({
  spacing: { before: 60, after: 240 }, alignment: AlignmentType.LEFT,
  children: [new TextRun({ text: t, italics: true, size: 19, color: MUTE })]
});
const figure = (file) => new Paragraph({
  spacing: { before: 80, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new ImageRun({ type: "png", data: img(file),
    transformation: { width: 552, height: 358 },
    altText: { title: file, description: file, name: file } })]
});

/* ---------- correlation table ---------- */
const border = { style: BorderStyle.SINGLE, size: 1, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };
const cell = (text, { head=false, w, bold=false, color, align="left" }={}) => new TableCell({
  borders, width: { size: w, type: WidthType.DXA },
  shading: { fill: head ? HEADFILL : "FFFFFF", type: ShadingType.CLEAR },
  margins: { top: 70, bottom: 70, left: 130, right: 130 },
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({ alignment: align==="right"?AlignmentType.RIGHT:AlignmentType.LEFT,
    children: [new TextRun({ text, bold: head||bold, color: color||(head?TEALD:INK), size: 20 })] })]
});
const COLW = [3360, 1800, 4200];
const trow = (cells) => new TableRow({ children: cells });
const corrTable = new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: COLW,
  rows: [
    trow([cell("Factor",{head:true,w:COLW[0]}), cell("Pearson r",{head:true,w:COLW[1],align:"right"}), cell("Reading",{head:true,w:COLW[2]})]),
    trow([cell("Glasses prescription level",{w:COLW[0]}), cell("−0.56",{w:COLW[1],align:"right",bold:true,color:CORAL}), cell("Strongest driver; not modifiable",{w:COLW[2]})]),
    trow([cell("Age",{w:COLW[0]}), cell("−0.49",{w:COLW[1],align:"right",color:CORAL}), cell("Steady life-course decline; not modifiable",{w:COLW[2]})]),
    trow([cell("Daily screen time (hrs)",{w:COLW[0]}), cell("−0.49",{w:COLW[1],align:"right",bold:true,color:CORAL}), cell("Largest MODIFIABLE risk factor",{w:COLW[2]})]),
    trow([cell("Screen brightness",{w:COLW[0]}), cell("−0.19",{w:COLW[1],align:"right",color:CORAL}), cell("Minor negative association",{w:COLW[2]})]),
    trow([cell("Outdoor light exposure (hrs)",{w:COLW[0]}), cell("+0.23",{w:COLW[1],align:"right",color:TEALD}), cell("Protective; modifiable",{w:COLW[2]})]),
    trow([cell("Exercise (hrs)",{w:COLW[0]}), cell("+0.24",{w:COLW[1],align:"right",color:TEALD}), cell("Protective; modifiable",{w:COLW[2]})]),
  ]
});

/* ---------- document ---------- */
const doc = new Document({
  creator: "Eye-Health Analytics",
  title: "Eye-Health Analytics — Exploratory Report",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22, color: INK } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, color: TEALD, font: "Calibri" },
        paragraph: { spacing: { before: 300, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: INK, font: "Calibri" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
    ]
  },
  numbering: { config: [
    { reference: "facts", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 280 } } } }] }
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 },
      margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    footers: { default: new Footer({ children: [ new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 6 } },
      children: [ new TextRun({ text: "Eye-Health Analytics  ·  exploratory report  ·  page ", size: 16, color: MUTE }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: MUTE }) ] }) ] }) },
    children: [
      // ---- title block ----
      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "EYE-HEALTH ANALYTICS", bold: true, size: 18, color: TEAL, characterSpacing: 30 }) ] }),
      new Paragraph({ spacing: { after: 60 }, children: [
        new TextRun({ text: "Exploratory Report & Data Story", bold: true, size: 44, color: INK }) ] }),
      new Paragraph({ spacing: { after: 40 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 8 } },
        children: [ new TextRun({ text: "Lifestyle, screen habits and vision across 10,000 respondents", size: 22, color: MUTE }) ] }),
      new Paragraph({ spacing: { before: 120, after: 200 }, children: [
        new TextRun({ text: "Source: eye_score_cleaned.xlsx", size: 18, color: MUTE }),
        new TextRun({ text: "    |    Companion file: index.html (interactive dashboard)", size: 18, color: MUTE }) ] }),

      // ---- dataset at a glance ----
      H1("1.  Dataset at a glance"),
      P([ new TextRun("The analysis uses a pre-cleaned survey of "),
          new TextRun({ text: "10,000 respondents", bold: true }),
          new TextRun(" described by 11 numeric attributes plus an ID. The cleaning log confirms the set is analysis-ready:") ]),
      new Paragraph({ numbering: { reference: "facts", level: 0 }, children: [ new TextRun("0 missing values and 0 duplicate rows — no imputation required.") ] }),
      new Paragraph({ numbering: { reference: "facts", level: 0 }, children: [ new TextRun("Float columns standardised to 4 decimals; domain ranges verified (age 5–80, screen time 0–16 h, scores 0–100).") ] }),
      new Paragraph({ numbering: { reference: "facts", level: 0 }, children: [ new TextRun("Ceiling values (e.g. 300 records at the 16-hour screen-time cap) were retained as valid and are disclosed.") ] }),
      new Paragraph({ numbering: { reference: "facts", level: 0 }, children: [ new TextRun({ text: "The set carries no date and no geographic field", bold: true }), new TextRun(", so the dashboard substitutes age / screen-time ranges for a date slider and a correlation heatmap for a map.") ] }),
      P([ new TextRun("The strongest linear relationships with the eye-health score are summarised below.") ], { spacing: { before: 80, after: 120 } }),
      corrTable,
      caption("Table 1. Pearson correlation of each attribute with the eye-health score (n = 10,000). Negative = score falls as the factor rises."),

      // ---- exploratory analysis ----
      H1("2.  Three exploratory charts"),
      H2("2.1  Correlation"),
      figure("chart_correlation.png"),
      caption("Figure 1. Each point is one respondent; the coral line is an ordinary-least-squares fit. Eye-health scores decline as daily screen time rises (r = −0.49, p < 0.001) — the clearest modifiable relationship in the data."),
      H2("2.2  Distribution"),
      figure("chart_distribution.png"),
      caption("Figure 2. The eye-health score is roughly bell-shaped and centred on a mean of 65.3 (median 65.5), spanning 21.7 to the 100-point ceiling. Mean and median almost coincide, indicating little skew."),
      H2("2.3  Time-series / life-course trend"),
      figure("chart_trend.png"),
      caption("Figure 3. Mean eye-health score across age, used as a life-course timeline because the dataset has no date field. Scores fall steadily from ~75 in childhood to ~56 by age 80 (slope ≈ −0.27 points per year, r = −0.49)."),

      // ---- data story ----
      new Paragraph({ pageBreakBefore: true, heading: HeadingLevel.HEADING_1, children: [ new TextRun("3.  Data Story") ] }),
      new Paragraph({ spacing: { after: 120 }, children: [ new TextRun({ text: `What protects our eyes?  (${storyWords} words)`, bold: true, size: 24, color: INK }) ] }),
      P([ new TextRun(story.intro) ]),
      P([ new TextRun({ text: story.audienceLead, bold: true, color: TEALD }), new TextRun(story.audience) ]),
      P([ new TextRun({ text: story.insightLead, bold: true, color: TEALD }), new TextRun(story.insight) ]),
      P([ new TextRun({ text: story.biasLead, bold: true, color: TEALD }), new TextRun(story.bias) ]),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  const out = path.join(HERE, "Eye-Health_Report_and_Data_Story.docx");
  fs.writeFileSync(out, buf);
  console.log("Wrote", out, (buf.length/1024).toFixed(0)+" KB");
});
