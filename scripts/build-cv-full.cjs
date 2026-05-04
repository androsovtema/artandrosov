// Build CV (full) — Art Androsov
// US Letter, 8 cases, airy 2-page layout with real page footer

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, TabStopType, ExternalHyperlink, Footer, PageNumber
} = require('docx');

// 0.8" all margins for cleaner US-Letter look
const MARGIN = 1152;
const TAB_RIGHT = 12240 - 2 * MARGIN;

const r = (text, opts = {}) => new TextRun({ text, font: "Arial", ...opts });
const dot = () => r(" · ", { color: "999999" });

// Section header — 12pt bold, more breathing room
const sectionHeader = (text) => new Paragraph({
  spacing: { before: 360, after: 160 },
  border: {
    bottom: { color: "BBBBBB", style: BorderStyle.SINGLE, size: 4, space: 2 }
  },
  children: [r(text, { bold: true, size: 24, color: "222222" })]
});

// Project / experience entry: title row + description, airy spacing
const entry = (titleParts, date, description) => {
  const titleChildren = [];
  titleParts.forEach((part, i) => {
    if (i === 0) {
      titleChildren.push(r(part, { bold: true, size: 24 }));
    } else {
      titleChildren.push(r("  ·  ", { size: 24, color: "999999" }));
      titleChildren.push(r(part, { size: 24 }));
    }
  });
  titleChildren.push(r("\t" + date, { size: 24, color: "555555" }));

  return [
    new Paragraph({
      keepNext: true,
      spacing: { before: 200, after: 60 },
      tabStops: [{ type: TabStopType.RIGHT, position: TAB_RIGHT }],
      children: titleChildren
    }),
    new Paragraph({
      keepLines: true,
      spacing: { after: 80 },
      children: [r(description, { size: 22, color: "333333" })]
    })
  ];
};

const doc = new Document({
  creator: "Artem Androsov",
  title: "Art Androsov — Product Designer CV",
  description: "Curriculum Vitae",
  styles: {
    default: {
      document: { run: { font: "Arial", size: 24 } }  // 12pt body default
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
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            r("Full portfolio  ·  ", { size: 18, color: "888888" }),
            new ExternalHyperlink({
              children: [r("androsov.art", { size: 18, color: "888888", underline: { type: "single" } })],
              link: "https://androsov.art"
            }),
            r("    ·    Page ", { size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888", font: "Arial" }),
            r(" of ", { size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "888888", font: "Arial" })
          ]
        })]
      })
    },
    children: [
      // ===== HEADER =====
      new Paragraph({
        spacing: { after: 100 },
        children: [r("ART ANDROSOV", { bold: true, size: 48 })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [r("Product Designer  ·  Bali (GMT+8)", { size: 26, color: "333333" })]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [r("Open to remote Product Designer roles  ·  contract or full-time", { size: 22, color: "555555" })]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [r("Working hours overlap from Bali (GMT+8): 8h EU  ·  5h US East  ·  2h US West", { size: 22, color: "555555" })]
      }),

      // Contacts
      new Paragraph({
        spacing: { after: 120 },
        border: { bottom: { color: "BBBBBB", style: BorderStyle.SINGLE, size: 4, space: 8 } },
        children: [
          new ExternalHyperlink({
            children: [r("art.androsov@gmail.com", { size: 24, color: "0563C1", underline: { type: "single" } })],
            link: "mailto:art.androsov@gmail.com"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("linkedin.com/in/artandrosov", { size: 24, color: "0563C1", underline: { type: "single" } })],
            link: "https://www.linkedin.com/in/artandrosov/"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("cal.com/artemandrosov/intro", { size: 24, color: "0563C1", underline: { type: "single" } })],
            link: "https://cal.com/artemandrosov/intro"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("androsov.art", { size: 24, color: "0563C1", underline: { type: "single" } })],
            link: "https://androsov.art"
          })
        ]
      }),

      // ===== SUMMARY =====
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [r(
          "Freelance product designer with 6+ years inside fast-shipping creator-economy teams. Specialised in product strategy, UX/UI, and art direction for ambiguous, high-stakes briefs. Currently working with clients across the UK, Switzerland, Latvia, Russia, and Indonesia.",
          { size: 24 }
        )]
      }),

      // ===== SELECTED WORK (8 cases) =====
      sectionHeader("SELECTED WORK"),

      ...entry(
        ["VIRIL", "Product, Brand Strategy, UX/UI"], "Oct 2024",
        "European men's health brand built from zero in 5 weeks. 64 screens across EN/ES × desktop/mobile, including a 7-question recommendation quiz that replaces self-diagnosis with guided choice. Live at viril.live."
      ),
      ...entry(
        ["EAZY ENERGY", "Product, Art Direction"], "2023 — 2025",
        "National energy drink. Led creative across packaging, 7-city tour, and racing sponsorships over 2 years. 13,359+ units sold · 4.5★ across 2,571 verified Wildberries ratings · plus offline + Ozon retail."
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
      ...entry(
        ["TALENT", "Product, UX/UI (self-initiated)"], "Dec 2024",
        "Multi-format licensing platform for independent creators — footage, music, SFX, images, licensed text under one roof. 45 screens, 17 modal states, fully prototyped in Figma with realistic content."
      ),
      ...entry(
        ["LOYO & BONDAR", "Product, Art Direction"], "Dec 2025 — Mar 2026",
        "Premium Bali villa developer. External art director sitting upstream of execution: decoded fragmented briefs from four marketing leads, defined per-audience communication logic, ran a pipeline pushing past 100 slides per peak week. Live at loyobondar.com/en/."
      ),
      ...entry(
        ["SCOVILLE", "Product, Packaging Design"], "Feb 2026",
        "Spicy ice cream packaging for Samokat — designed twice. First version killed by retailer mid-process; second built from near-zero in days, kept the cold/spicy product tension intact. Live in Samokat catalogue."
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
        spacing: { before: 100, after: 0 },
        children: [r("Product strategy  ·  UX architecture  ·  Figma (advanced)  ·  Brand systems  ·  Stakeholder management  ·  Art direction  ·  Adobe CC (Pr, Ae, Ps, Ai)  ·  DaVinci Resolve  ·  AI tools", { size: 24 })]
      }),

      // ===== EDUCATION =====
      sectionHeader("EDUCATION"),

      new Paragraph({
        spacing: { before: 100, after: 30 },
        tabStops: [{ type: TabStopType.RIGHT, position: TAB_RIGHT }],
        children: [
          r("Moscow University for the Humanities", { bold: true, size: 24 }),
          r("  ·  ", { size: 24, color: "999999" }),
          r("Banking", { size: 24 }),
          r("\t2015 — 2018", { size: 24, color: "555555" })
        ]
      }),
      new Paragraph({
        spacing: { after: 140 },
        children: [r("With honours", { size: 22, italics: true, color: "555555" })]
      }),

      new Paragraph({
        spacing: { before: 80, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: TAB_RIGHT }],
        children: [
          r("RANEPA", { bold: true, size: 24 }),
          r("  ·  ", { size: 24, color: "999999" }),
          r("Bachelor of Economics", { size: 24 }),
          r("\t2019 — 2023", { size: 24, color: "555555" })
        ]
      }),

      // ===== LANGUAGES =====
      sectionHeader("LANGUAGES"),
      new Paragraph({
        spacing: { before: 100, after: 0 },
        children: [r("Russian (native)  ·  English (B2)", { size: 24 })]
      }),

      // ===== APPROACH =====
      sectionHeader("APPROACH"),
      new Paragraph({
        spacing: { before: 80, after: 60 },
        children: [r("— Lead briefs with Jobs-To-Be-Done — extract what the user is actually trying to accomplish before designing the artifact.", { size: 22, color: "333333" })]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [r("— Build the system before the artifact — define what the brief requires, then design within that frame.", { size: 22, color: "333333" })]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [r("— Iterate on user behaviour, not stakeholder preference. Ship, watch, adjust.", { size: 22, color: "333333" })]
      }),
      new Paragraph({
        spacing: { after: 0 },
        children: [r("— AI-assisted prototyping when speed-to-real-users matters.", { size: 22, color: "333333" })]
      })
    ]
  }]
});

const outPath = process.argv[2] || "/sessions/jolly-beautiful-thompson/mnt/artandrosov/cv/Art_Androsov_CV_full.docx";
const outDir = path.dirname(outPath);
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Wrote ${outPath} (${buffer.length} bytes)`);
});
