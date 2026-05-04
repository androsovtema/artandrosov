// Build CV (RU) — Art Androsov
// US Letter, single column, ATS-friendly, tight one-page layout

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, TabStopType, ExternalHyperlink
} = require('docx');

const MARGIN_LR = 1008;
const MARGIN_TB = 864;
const TAB_RIGHT = 12240 - 2 * MARGIN_LR;

const r = (text, opts = {}) => new TextRun({ text, font: "Arial", ...opts });
const dot = () => r(" · ", { color: "999999" });

const sectionHeader = (text) => new Paragraph({
  spacing: { before: 180, after: 60 },
  border: {
    bottom: { color: "BBBBBB", style: BorderStyle.SINGLE, size: 4, space: 1 }
  },
  children: [r(text, { bold: true, size: 22, color: "222222" })]
});

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
  title: "Артём Андросов — резюме продуктового дизайнера",
  description: "Резюме",
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
      new Paragraph({
        spacing: { after: 60 },
        children: [r("АРТЁМ АНДРОСОВ", { bold: true, size: 44 })]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [r("Продуктовый дизайнер  ·  Бали (GMT+8)", { size: 24, color: "333333" })]
      }),
      new Paragraph({
        spacing: { after: 20 },
        children: [r("Открыт к удалённым ролям Product Designer  ·  контракт или full-time", { size: 21, color: "555555" })]
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [r("Пересечение рабочих часов из Бали (GMT+8): 8ч EU  ·  5ч US East  ·  2ч US West", { size: 21, color: "555555" })]
      }),

      new Paragraph({
        spacing: { after: 80 },
        border: { bottom: { color: "BBBBBB", style: BorderStyle.SINGLE, size: 4, space: 6 } },
        children: [
          new ExternalHyperlink({
            children: [r("art.androsov@gmail.com", { style: "Hyperlink", size: 22 })],
            link: "mailto:art.androsov@gmail.com"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("linkedin.com/in/artandrosov", { style: "Hyperlink", size: 22 })],
            link: "https://www.linkedin.com/in/artandrosov/"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("cal.com/artemandrosov/intro", { style: "Hyperlink", size: 22 })],
            link: "https://cal.com/artemandrosov/intro"
          }),
          dot(),
          new ExternalHyperlink({
            children: [r("androsov.art", { style: "Hyperlink", size: 22 })],
            link: "https://androsov.art"
          })
        ]
      }),

      new Paragraph({
        spacing: { before: 120, after: 60 },
        children: [r(
          "Продуктовый дизайнер с 6+ годами опыта в быстро движущихся creator-economy командах. Специализируюсь на продуктовой стратегии, UX/UI и арт-дирекшене для неоднозначных задач с высокой ставкой. Сейчас работаю с клиентами из Великобритании, Швейцарии, Латвии, России и Индонезии.",
          { size: 21 }
        )]
      }),

      sectionHeader("ИЗБРАННЫЕ ПРОЕКТЫ"),

      ...entry(
        ["VIRIL", "Продукт, бренд-стратегия, UX/UI"], "Окт 2024",
        "Европейский бренд мужского здоровья, собранный с нуля за 5 недель. 64 экрана для EN/ES × desktop/mobile, включая квиз из 7 вопросов для подбора продукта. Live: viril.live."
      ),
      ...entry(
        ["EAZY ENERGY", "Продукт, арт-дирекшен"], "2023 — 2025",
        "Национальный энергетический напиток. Вёл креатив по упаковке, туру на 7 городов и гоночным спонсорствам. 13 359+ проданных единиц · 4.5★ по 2 571 проверенному отзыву Wildberries · плюс офлайн и Ozon."
      ),
      ...entry(
        ["GABARSHOP", "Продукт, e-commerce"], "2020 — сейчас",
        "Мерч-магазин creator-команды. 4 года внутри команды, 2+ года on-call. End-to-end ownership: от выбора материалов на фабрике и упаковки заказов до полного редизайна storefront UX. 100K+ единиц в 40+ SKU. Live: gabarshop.com."
      ),
      ...entry(
        ["EMTECH INVEST × DAVOS", "Продукт, арт-дирекшен"], "Янв 2026",
        "Закрытое инвестиционное событие во время недели WEF в Давосе. Бренд собран почти с нуля за 6 недель: 66 профилей участников, материалы для полного lifecycle события. Press partner: Cryptopolitan."
      ),
      ...entry(
        ["OUR WALL", "Продукт, UX (self-initiated)"], "Сен 2025 — сейчас",
        "Анонимная социальная сеть, собранная на AI-assisted code за одну неделю. 1 132 органических пользователя, 1 311 постов, единственный источник привлечения — одно YouTube-видео. Live: ourwall.ru."
      ),

      sectionHeader("ОПЫТ"),

      ...entry(
        ["Freelance Product Designer"], "2024 — сейчас",
        "Независимая практика. Клиенты: Viril (UK), EmTech Invest (CH), LOYO & BONDAR (ID/LV), Eazy Energy (RU), Scoville (RU), Gabarshop (RU, on-call)."
      ),
      ...entry(
        ["YouTube Media", "Product Designer / Head of Digital"], "2020 — 2024",
        "Отвечал за digital-поверхность команды топового YouTube-автора: мерч (100K+ продаж), сайт, соцканалы (рост до 500K+ подписчиков на одной платформе), контент-продакшен. Продуктовое мышление применял к каждому брифу."
      ),
      ...entry(
        ["We Designerz", "Основатель YouTube-канала"], "2024 — сейчас",
        "Независимый канал про AI в работе и дизайне. End-to-end production: съёмка, монтаж, motion graphics, звук. 2 700 подписчиков · 383K+ просмотров."
      ),

      sectionHeader("НАВЫКИ"),
      new Paragraph({
        spacing: { before: 40, after: 0 },
        children: [r("Продуктовая стратегия  ·  UX-архитектура  ·  Figma (advanced)  ·  Бренд-системы  ·  Работа со стейкхолдерами  ·  Арт-дирекшен  ·  Adobe CC (Pr, Ae, Ps, Ai)  ·  DaVinci Resolve  ·  AI-assisted prototyping", { size: 21 })]
      }),

      sectionHeader("ОБРАЗОВАНИЕ"),

      new Paragraph({
        spacing: { before: 40, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: TAB_RIGHT }],
        children: [
          r("Московский гуманитарный университет", { bold: true, size: 22 }),
          r("  ·  ", { size: 22, color: "999999" }),
          r("Банковское дело", { size: 22 }),
          r("  ·  ", { size: 22, color: "999999" }),
          r("С отличием", { size: 21, italics: true, color: "555555" }),
          r("\t2015 — 2018", { size: 22, color: "555555" })
        ]
      }),
      new Paragraph({
        spacing: { before: 40, after: 0 },
        tabStops: [{ type: TabStopType.RIGHT, position: TAB_RIGHT }],
        children: [
          r("РАНХиГС", { bold: true, size: 22 }),
          r("  ·  ", { size: 22, color: "999999" }),
          r("Бакалавр экономики", { size: 22 }),
          r("\t2019 — 2023", { size: 22, color: "555555" })
        ]
      }),

      sectionHeader("ЯЗЫКИ"),
      new Paragraph({
        spacing: { before: 40, after: 0 },
        children: [r("Русский (родной)  ·  Английский (B2)", { size: 21 })]
      }),

      new Paragraph({
        spacing: { before: 140 },
        alignment: AlignmentType.CENTER,
        children: [
          r("Полное портфолио  ·  ", { size: 19, color: "777777" }),
          new ExternalHyperlink({
            children: [r("androsov.art", { style: "Hyperlink", size: 19 })],
            link: "https://androsov.art"
          })
        ]
      })
    ]
  }]
});

const outPath = process.argv[2] || path.resolve(__dirname, "../cv/ru/Art_Androsov_Resume_RU.docx");
const outDir = path.dirname(outPath);
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Wrote ${outPath} (${buffer.length} bytes)`);
});
