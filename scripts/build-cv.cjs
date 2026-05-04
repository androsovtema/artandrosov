// Build CV — Art Androsov
// US Letter, single column, ATS-friendly, tight one-page layout

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, TabStopType, ExternalHyperlink
} = require('docx');

// 0.7" left/right margins, 0.6" top/bottom for max content area
// Content width = 12240 - 2*1008 = 10224
const MARGIN_LR = 1008;  // 0.7"
const MARGIN_TB = 864;   // 0.6"
const TAB_RIGHT = 12240 - 2 * MARGIN_LR;

const r = (text, opts = {}) => new TextRun({ text, font: "Arial", ...opts });
const dot = () => r(" · ", { color: "999999" });

// Section header — bold uppercase with bottom border, tight spacing
const sectionHeader = (text) => new Paragraph({
  spacing: { before: 180, after: 60 },
  border: {
    bottom: { color: "BBBBBB", style: BorderStyle.SINGLE, size: 4, space: 1 }
  },
  children: [r(text, { bold: true, size: 22, color: "222222" })]
});

// Project / experience entry: title row + description, tight
const entry = (titleParts, date, description) => {
  const titleChildren = [];
  titleParts.forEach((part, i) => {
    if (i === 0) {
      titleChildren.push(r(part, { bold: true, size: 22 }));
    } else {
      titleChildren.push(r("  ·  ", { size: 22, color: "999999" }));
      titleChildren.push(r(part, { size: 22 }));
    }
  });
  titleChildren.push(r("\t" + date, { size: 22, color: "555555" }));

  return [
    new Paragraph({
      spacing: { before: 80, after: 20 },
      tabStops: [{ type: TabStopType.RIGHT, position: TAB_RIGHT }],
      children: titleChildren
    }),
    new Paragraph({
      spacing: { after: 0 },
      children: [r(description, { size: 20, color: "333333" })]
    })
  ];
};

const doc = new Document({
  creator: "Artem Androsov",
  title: "Art Androsov — Product Designer CV",
  description: "Curriculum Vitae",
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22 } }
    },
    characterStyles: [
      {
        id: "Hyperlink",
        name: "Hyperlink",
        basedOn: "DefaultParagraphFont",
        run: { color: "0563C1", underline: { type: "single" } }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: MARGIN_TB, right: MARGIN_LR, bottom: MARGIN_TB, left: MARGIN_LR }
      }
    },
    children: [
      // ===== HEADER =====
      new Paragraph({
        spacing: { after: 60 },
        children: [r("ART ANDROSOV", { bold: true, size: 44 })]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [r("Product Designer  ·  Bali (GMT+8)", { size: 24, color: "333333" })]
      }),
      new Paragraph({
        spacing: { after: 20 },
        children: [r("Open to remote Product Designer roles  ·  contract or full-time", { size: 21, color: "555555" })]
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [r("Working hours overlap from Bali (GMT+8): 8h EU  ·  5h US East  ·  2h US West", { size: 21, color: "555555" })]
      }),

      // Contacts
      new Paragraph({
        spacing: { after: 80 },
        border: { bottom: { color: "BBBBBB", style: BorderStyle.SINGLE, size: 4, space: 6 } },
        children: [
          new ExternalHyperlink({
            children: [r("art.androsov@gmail.com", { style: "Hyperlink", size: 22, color: "0563C1", underline: { type: "single" } })],
            link: "mailto:art.androsov@gmail.com"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("linkedin.com/in/artandrosov", { style: "Hyperlink", size: 22, color: "0563C1", underline: { type: "single" } })],
            link: "https://www.linkedin.com/in/artandrosov/"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("cal.com/artemandrosov/intro", { style: "Hyperlink", size: 22, color: "0563C1", underline: { type: "single" } })],
            link: "https://cal.com/artemandrosov/intro"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("androsov.art", { style: "Hyperlink", size: 22, color: "0563C1", underline: { type: "single" } })],
            link: "https://androsov.art"
          })
        ]
      }),

      // ===== SUMMARY =====
      new Paragraph({
        spacing: { before: 120, after: 60 },
        children: [r(
          "Freelance product designer with 6+ years inside fast-shipping creator-economy teams. Specialised in product strategy, UX/UI, and art direction for ambiguous, high-stakes briefs. Currently working with clients across the UK, Switzerland, Latvia, Russia, and Indonesia.",
          { size: 21 }
        )]
      }),

      // ===== SELECTED WORK =====
      sectionHeader("SELECTED WORK"),

      ...entry(
        ["VIRIL", "Product, Brand Strategy, UX/UI"], "Oct 2024",
        "European men's health brand built from zero in 5 weeks. 64 screens across EN/ES × desktop/mobile, including a 7-question recommendation quiz. Live at viril.live."
      ),
      ...entry(
        ["EAZY ENERGY", "Product, Art Direction"], "2023 — 2025",
        "National energy drink. Led creative across packaging, 7-city tour, and racing sponsorships. 13,359+ units sold · 4.5★ across 2,571 verified Wildberries ratings · plus offline + Ozon retail."
      ),
      ...entry(
        ["GABARSHOP", "Product, E-commerce"], "2020 — present",
        "Creator merch store. 4 years in-house, 2+ years on-call. End-to-end ownership: from material selection at the factory and packing orders by hand to the full storefront UX rebuild. 100K+ units across 40+ SKUs. Live at gabarshop.com."
      ),
      ...entry(
        ["EMTECH INVEST × DAVOS", "Product, Art Direction"], "Jan 2026",
        "Closed-door investment event during WEF week. Brand built from near-zero in 6 weeks, 66 participant profiles, full event lifecycle materials. Press partner: Cryptopolitan."
      ),
      ...entry(
        ["OUR WALL", "Product, UX (self-initiated)"], "Sep 2025 — present",
        "Anonymous social network built on AI-assisted code in one week. 1,132 organic users, 1,311 posts, single YouTube video as the only acquisition source. Live at ourwall.ru."
      ),

      // ===== EXPERIENCE =====
      sectionHeader("EXPERIENCE"),

      ...entry(
        ["Freelance Product Designer"], "2024 — present",
        "Independent practice. Clients: Viril (UK), EmTech Invest (CH), LOYO & BONDAR (ID/LV), Eazy Energy (RU), Scoville (RU), Gabarshop (RU, on-call)."
      ),
      ...entry(
        ["YouTube Media", "Product Designer / Head of Digital"], "2020 — 2024",
        "Owned the full digital surface for a top YouTube creator's team: merchandise (100K+ units sold), website, social channels (grew to 500K+ subscribers on one platform), content production. Product thinking applied to every brief."
      ),
      ...entry(
        ["We Designerz", "YouTube Channel founder"], "2024 — present",
        "Independent channel about AI in work and design. End-to-end production — shooting, editing, motion graphics, sound design. 2,700 subscribers · 383K+ views."
      ),

      // ===== SKILLS =====
      sectionHeader("SKILLS"),
      new Paragraph({
        spacing: { before: 40, after: 0 },
        children: [r("Product strategy  ·  UX architecture  ·  Figma (advanced)  ·  Brand systems  ·  Stakeholder management  ·  Art direction  ·  Adobe CC (Pr, Ae, Ps, Ai)  ·  DaVinci Resolve  ·  AI tools", { size: 21 })]
      }),

      // ===== EDUCATION =====
      sectionHeader("EDUCATION"),

      new Paragraph({
        spacing: { before: 40, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: TAB_RIGHT }],
        children: [
          r("Moscow University for the Humanities", { bold: true, size: 22 }),
          r("  ·  ", { size: 22, color: "999999" }),
          r("Banking", { size: 22 }),
          r("  ·  ", { size: 22, color: "999999" }),
          r("With honours", { size: 21, italics: true, color: "555555" }),
          r("\t2015 — 2018", { size: 22, color: "555555" })
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: TAB_RIGHT }],
        children: [
          r("RANEPA", { bold: true, size: 22 }),
          r("  ·  ", { size: 22, color: "999999" }),
          r("Bachelor of Economics", { size: 22 }),
          r("\t2019 — 2023", { size: 22, color: "555555" })
        ]
      }),

      // ===== LANGUAGES =====
      sectionHeader("LANGUAGES"),
      new Paragraph({
        spacing: { before: 40, after: 0 },
        children: [r("Russian (native)  ·  English (B2)", { size: 21 })]
      }),

      // ===== FOOTER =====
      new Paragraph({
        spacing: { before: 140 },
        alignment: AlignmentType.CENTER,
        children: [
          r("Full portfolio  ·  ", { size: 19, color: "777777" }),
          new ExternalHyperlink({
            children: [r("androsov.art", { style: "Hyperlink", size: 19, color: "0563C1", underline: { type: "single" } })],
            link: "https://androsov.art"
          })
        ]
      })
    ]
  }]
});

const outPath = process.argv[2] || "/sessions/jolly-beautiful-thompson/mnt/artandrosov/cv/Art_Androsov_CV.docx";
const outDir = path.dirname(outPath);
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Wrote ${outPath} (${buffer.length} bytes)`);
});
