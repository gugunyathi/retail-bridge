import fs from "node:fs";
import PptxGenJS from "pptxgenjs";

const RED = "D0103A";
const RED_DEEP = "A50C2D";
const BLUE = "1B3A6B";
const BLUE_DEEP = "10243F";
const PAPER = "F7F7F9";
const INK = "16233A";
const MUTED = "5C6779";
const GOLD = "F2B233";
const WHITE = "FFFFFF";

const img = (p) => `image/${p.endsWith(".png") ? "png" : "jpeg"};base64,${fs.readFileSync(p).toString("base64")}`;
const LOGO = img("src/assets/tmpnp-logo.png");
const DOOR = img("src/assets/delivery-door.jpg");
const SHOPPER = img("src/assets/diaspora-shopper.jpg");
const PICKING = img("src/assets/store-picking.jpg");
const BIKE = img("src/assets/bike-courier.jpg");

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9"; // 10 x 5.625 in, landscape
const W = 10;
const H = 5.625;
const HEAD = "Georgia";
const BODY = "Calibri";

function chrome(slide, kicker, n, dark, narrow = 0) {
  const right = narrow ? narrow : W - 0.45; // right edge available for chrome text
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 0.3, w: 1.5, h: 0.4, fill: { color: WHITE }, line: { color: WHITE }, rectRadius: 0.06 });
  slide.addImage({ data: LOGO, x: 0.53, y: 0.37, w: 1.34, h: 0.26 });
  slide.addText(kicker.toUpperCase(), { x: 2.1, y: 0.33, w: right - 2.7, h: 0.35, fontFace: BODY, fontSize: 11, bold: true, charSpacing: 2, color: dark ? "FFFFFF" : RED });
  slide.addText(String(n).padStart(2, "0"), { x: right - 0.55, y: 0.33, w: 0.55, h: 0.35, align: "right", fontFace: BODY, fontSize: 11, bold: true, color: dark ? "AAB6C8" : MUTED });
  slide.addText("TM Pick n Pay Express — Diaspora-to-Door", { x: 0.45, y: H - 0.5, w: 5, h: 0.3, fontFace: BODY, fontSize: 9, color: dark ? "8C9BB0" : MUTED });
  if (!narrow) slide.addText("Confidential · Executive Board Proposal", { x: W - 5.45, y: H - 0.5, w: 5, h: 0.3, align: "right", fontFace: BODY, fontSize: 9, color: dark ? "8C9BB0" : MUTED });
}


function title(slide, text, dark, y = 0.95) {
  slide.addText(text, { x: 0.45, y, w: W - 0.9, h: 0.8, fontFace: HEAD, fontSize: 30, bold: true, color: dark ? WHITE : BLUE });
}

function card(slide, { x, y, w, h, heading, body, accent = RED, dark = false, hs = 14, bs = 12, hh = 0.58 }) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: dark ? "1E426F" : WHITE }, line: { color: dark ? "1E426F" : "E2E5EB" }, rectRadius: 0.08 });
  slide.addShape(pptx.ShapeType.roundRect, { x: x + 0.25, y: y + 0.2, w: 0.45, h: 0.06, fill: { color: accent }, line: { color: accent }, rectRadius: 0.03 });
  slide.addText(heading, { x: x + 0.25, y: y + 0.36, w: w - 0.5, h: hh, fontFace: HEAD, fontSize: hs, bold: true, color: dark ? WHITE : BLUE, margin: 0, valign: "top" });
  slide.addText(body, { x: x + 0.25, y: y + 0.36 + hh + 0.04, w: w - 0.5, h: h - (0.36 + hh + 0.04) - 0.18, fontFace: BODY, fontSize: bs, color: dark ? "D5DEEA" : MUTED, margin: 0, valign: "top" });
}


/* 1 — Title */
{
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  s.addImage({ data: DOOR, x: 4.6, y: 0, w: 5.4, h: H, sizing: { type: "cover", w: 5.4, h: H } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 5.6, h: H, fill: { color: BLUE_DEEP } });
  s.addShape(pptx.ShapeType.rect, { x: 5.6, y: 0, w: 1.1, h: H, fill: { color: BLUE_DEEP, transparency: 45 }, line: { color: BLUE_DEEP, transparency: 100 } });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 0.45, w: 1.8, h: 0.48, fill: { color: WHITE }, line: { color: WHITE }, rectRadius: 0.07 });
  s.addImage({ data: LOGO, x: 0.55, y: 0.54, w: 1.6, h: 0.31 });
  s.addText("EXECUTIVE BOARD PROPOSAL", { x: 0.45, y: 1.6, w: 5, h: 0.3, fontFace: BODY, fontSize: 12, bold: true, charSpacing: 3, color: GOLD });
  s.addText("TM Pick n Pay Express", { x: 0.45, y: 2.0, w: 5.3, h: 0.75, fontFace: HEAD, fontSize: 33, bold: true, color: WHITE });
  s.addText("Evolving Click & Collect into Diaspora-to-Door delivery", { x: 0.45, y: 2.85, w: 5.0, h: 0.7, fontFace: BODY, fontSize: 17, color: "D5DEEA" });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 3.75, w: 1.2, h: 0.08, fill: { color: RED }, line: { color: RED }, rectRadius: 0.04 });
  s.addText("Prepared for the Executive Board · TM Pick n Pay Zimbabwe & Meikles Limited", { x: 0.45, y: 4.1, w: 5, h: 0.5, fontFace: BODY, fontSize: 12, color: "9FB0C6" });
}

/* 2 — Opportunity */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 01 · The Opportunity", 1, false);
  title(s, "Monetising the digital infrastructure you already own", false);
  const cw = 2.9, ch = 2.0;
  card(s, { x: 0.45, y: 1.9, w: cw, h: ch, heading: "The foundation", body: "tmpnponline.co.zw and the dedicated app are already live, running localized Click & Collect across the estate." });
  card(s, { x: 0.45 + cw + 0.15, y: 1.9, w: cw, h: ch, heading: "The optimization gap", accent: BLUE, body: "Collection demands transport, fuel and time. Diaspora buyers still pay Malayitsha vans purely for doorstep convenience." });
  card(s, { x: 0.45 + 2 * (cw + 0.15), y: 1.9, w: cw, h: ch, heading: "Our value proposition", body: "A Diaspora UI mode plus a decentralised last-mile network turns 57+ branches into on-demand fulfilment nodes." });
  s.addImage({ data: PICKING, x: 0.45, y: 4.05, w: 9.1, h: 1.05, sizing: { type: "cover", w: 9.1, h: 1.05 } });
}

/* 3 — Status quo vs evolution */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 02 · Status Quo vs Evolution", 2, false);
  title(s, "Bypassing physical logistics friction", false);
  const flow = (y, label, dot, steps) => {
    s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y, w: 9.1, h: 1.15, fill: { color: WHITE }, line: { color: "E2E5EB" }, rectRadius: 0.08 });
    s.addShape(pptx.ShapeType.ellipse, { x: 0.7, y: y + 0.22, w: 0.16, h: 0.16, fill: { color: dot }, line: { color: dot } });
    s.addText(label, { x: 0.95, y: y + 0.13, w: 8.4, h: 0.35, fontFace: HEAD, fontSize: 14, bold: true, color: BLUE, margin: 0 });
    const bw = (8.6 - (steps.length - 1) * 0.18) / steps.length;
    steps.forEach((st, i) => {
      const x = 0.7 + i * (bw + 0.18);
      s.addShape(pptx.ShapeType.roundRect, { x, y: y + 0.6, w: bw, h: 0.42, fill: { color: PAPER }, line: { color: "E2E5EB" }, rectRadius: 0.06 });
      s.addText(st, { x, y: y + 0.6, w: bw, h: 0.42, align: "center", valign: "middle", fontFace: BODY, fontSize: 10, bold: true, color: INK, margin: 0 });
    });
  };
  flow(1.95, "Current system — Click & Collect", RED, ["Diaspora shopper", "Web / app order", "Recipient must travel", "Urban branches only"]);
  flow(3.25, "The evolution — On-demand diaspora engine", BLUE, ["Diaspora shopper", "Targeted ads", "Pure US$ gateway", "Bike courier", "Recipient's door"]);
  s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 4.55, w: 4.45, h: 0.6, fill: { color: BLUE }, line: { color: BLUE }, rectRadius: 0.08 });
  s.addText("Rural and elderly recipients cannot easily reach a flagship urban branch to collect heavy hampers.", { x: 0.65, y: 4.55, w: 4.05, h: 0.6, valign: "middle", fontFace: BODY, fontSize: 10.5, color: WHITE, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 5.1, y: 4.55, w: 4.45, h: 0.6, fill: { color: RED }, line: { color: RED }, rectRadius: 0.08 });
  s.addText("Door-to-door fulfilment adds sponsors who want confirmation that food arrived safely.", { x: 5.3, y: 4.55, w: 4.05, h: 0.6, valign: "middle", fontFace: BODY, fontSize: 10.5, color: WHITE, margin: 0 });
}

/* 4 — GMV assumptions */
{
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  chrome(s, "Slide 03 · P&L Assumptions", 3, true);
  title(s, "Monetising every stage of the cross-border basket", true);
  const stats = [["40,000", "Active diaspora customers"], ["1.5", "Orders per month"], ["US$85", "Average basket size"], ["12", "Months"]];
  stats.forEach(([v, l], i) => {
    const x = 0.45 + i * 2.31;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.95, w: 2.11, h: 1.2, fill: { color: "1E426F" }, line: { color: "1E426F" }, rectRadius: 0.08 });
    s.addText(v, { x: x + 0.2, y: 2.1, w: 1.75, h: 0.6, fontFace: HEAD, fontSize: 26, bold: true, color: WHITE, margin: 0 });
    s.addText(l, { x: x + 0.2, y: 2.68, w: 1.75, h: 0.4, fontFace: BODY, fontSize: 10, color: "AEBED2", margin: 0 });
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 3.35, w: 9.1, h: 1.5, fill: { color: RED }, line: { color: RED }, rectRadius: 0.1 });
  s.addText("ANNUAL GROSS MERCHANDISE VALUE", { x: 0.85, y: 3.55, w: 5, h: 0.3, fontFace: BODY, fontSize: 11, bold: true, charSpacing: 2, color: "F7CBD4", margin: 0 });
  s.addText("US$61,200,000", { x: 0.85, y: 3.85, w: 5, h: 0.8, fontFace: HEAD, fontSize: 36, bold: true, color: WHITE, margin: 0 });
  s.addText("40,000 customers × 1.5 orders × US$85 × 12 months — migrated from cash remittances and Malayitsha vans into a tracked US$ retail pipeline.", { x: 6.0, y: 3.6, w: 3.2, h: 1.05, fontFace: BODY, fontSize: 11, color: "FBE3E8", margin: 0, valign: "middle" });
}

/* 5 — Revenue streams */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 03 · Five Revenue Streams", 4, false);
  title(s, "Five stacked revenue streams", false);
  const items = [
    ["1. Retail product margins", "Spread captured on goods otherwise bought in South African cash-and-carries. 21% gross margin.", "US$12,852,000"],
    ["2. Last-mile delivery share", "US$4.50 fee within 10km; the platform retains US$1.50 net per drop.", "US$1,080,000"],
    ["3. Cross-border surcharge", "3% checkout fee on international cards originating outside Zimbabwe.", "US$1,836,000"],
    ["4. Diaspora Priority tiers", "US$8.99/month for free delivery and recurring staple baskets. 15% adoption.", "US$647,280"],
    ["5. Retail media network", "FMCG brands bid for sponsored placement in front of high-spend diaspora buyers.", "US$734,400"],
  ];
  const cw = 2.95, ch = 1.35;
  items.forEach((it, i) => {
    const x = 0.45 + (i % 3) * (cw + 0.15);
    const y = 1.95 + Math.floor(i / 3) * (ch + 0.15);
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: cw, h: ch, fill: { color: WHITE }, line: { color: "E2E5EB" }, rectRadius: 0.08 });
    s.addText(it[0], { x: x + 0.2, y: y + 0.15, w: cw - 0.4, h: 0.3, fontFace: HEAD, fontSize: 13, bold: true, color: BLUE, margin: 0 });
    s.addText(it[1], { x: x + 0.2, y: y + 0.48, w: cw - 0.4, h: 0.5, fontFace: BODY, fontSize: 10, color: MUTED, margin: 0, valign: "top" });
    s.addText(it[2], { x: x + 0.2, y: y + 1.0, w: cw - 0.4, h: 0.3, fontFace: HEAD, fontSize: 15, bold: true, color: RED, margin: 0 });
  });
  const x = 0.45 + 2 * (cw + 0.15), y = 1.95 + (ch + 0.15);
  s.addShape(pptx.ShapeType.roundRect, { x, y, w: cw, h: ch, fill: { color: BLUE }, line: { color: BLUE }, rectRadius: 0.08 });
  s.addText("TOTAL ECOSYSTEM", { x: x + 0.2, y: y + 0.2, w: cw - 0.4, h: 0.25, fontFace: BODY, fontSize: 10, bold: true, charSpacing: 2, color: GOLD, margin: 0 });
  s.addText("US$17,149,680", { x: x + 0.2, y: y + 0.5, w: cw - 0.4, h: 0.45, fontFace: HEAD, fontSize: 22, bold: true, color: WHITE, margin: 0 });
  s.addText("Projected annual revenue", { x: x + 0.2, y: y + 0.98, w: cw - 0.4, h: 0.3, fontFace: BODY, fontSize: 10, color: "C6D3E4", margin: 0 });
}

/* 6 — Breakdown chart */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 04 · Revenue Breakdown", 5, false);
  title(s, "Projected total ecosystem annual revenue", false);
  s.addText("US$17,149,680", { x: 0.45, y: 1.62, w: 4, h: 0.4, fontFace: HEAD, fontSize: 20, bold: true, color: RED });
  const rows = [
    ["Incremental retail margins", 12852000, "21% gross retail margin"],
    ["Last-mile delivery share", 1080000, "US$1.50 net per drop"],
    ["Cross-border tech surcharge", 1836000, "3% on international cards"],
    ["Diaspora Priority subscriptions", 647280, "6,000 subscribers @ US$8.99"],
    ["Supplier-funded retail media", 734400, "1.2% of platform GMV"],
  ];
  const max = 12852000;
  rows.forEach((r, i) => {
    const y = 2.2 + i * 0.55;
    s.addText(r[0], { x: 0.45, y, w: 3.0, h: 0.26, fontFace: BODY, fontSize: 12, bold: true, color: INK, margin: 0 });
    s.addText(r[2], { x: 0.45, y: y + 0.24, w: 3.0, h: 0.24, fontFace: BODY, fontSize: 9.5, color: MUTED, margin: 0 });
    s.addShape(pptx.ShapeType.roundRect, { x: 3.6, y: y + 0.1, w: 4.35, h: 0.3, fill: { color: "E7EAF0" }, line: { color: "E7EAF0" }, rectRadius: 0.15 });
    s.addShape(pptx.ShapeType.roundRect, { x: 3.6, y: y + 0.1, w: Math.max(0.35, (r[1] / max) * 4.35), h: 0.3, fill: { color: RED }, line: { color: RED }, rectRadius: 0.15 });
    s.addText("US$" + r[1].toLocaleString("en-US"), { x: 8.1, y: y + 0.08, w: 1.45, h: 0.32, align: "right", fontFace: HEAD, fontSize: 13, bold: true, color: BLUE, margin: 0 });
  });
}

/* 7 & 8 — Business model options */
function optionSlide(n, heading, kicker, opts) {
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, kicker, n, false);
  title(s, heading, false);
  opts.forEach((o, i) => {
    const x = 0.45 + i * 4.7;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.9, w: 4.4, h: 2.9, fill: { color: WHITE }, line: { color: "E2E5EB" }, rectRadius: 0.08 });
    s.addShape(pptx.ShapeType.roundRect, { x: x + 0.22, y: 2.06, w: 0.95, h: 0.28, fill: { color: RED }, line: { color: RED }, rectRadius: 0.14 });
    s.addText(o.tag, { x: x + 0.22, y: 2.06, w: 0.95, h: 0.28, align: "center", valign: "middle", fontFace: BODY, fontSize: 9.5, bold: true, color: WHITE, margin: 0 });
    s.addText(o.title, { x: x + 1.25, y: 2.02, w: 3.0, h: 0.34, fontFace: HEAD, fontSize: 13.5, bold: true, color: BLUE, margin: 0, valign: "middle" });
    s.addText(o.body, { x: x + 0.22, y: 2.45, w: 3.98, h: 1.35, fontFace: BODY, fontSize: 10.5, color: MUTED, margin: 0, valign: "top" });
    [["Setup", o.setup], ["Monthly", o.opex], ["Revenue", o.rev]].forEach(([k, v], j) => {
      const cx = x + 0.22 + j * 1.33;
      s.addText(k.toUpperCase(), { x: cx, y: 3.9, w: 1.28, h: 0.2, fontFace: BODY, fontSize: 8.5, bold: true, charSpacing: 1, color: MUTED, margin: 0 });
      s.addText(v, { x: cx, y: 4.1, w: 1.28, h: 0.45, fontFace: BODY, fontSize: 10, bold: true, color: INK, margin: 0, valign: "top" });
    });
  });
}

optionSlide(6, "Aligning risk, capital and structure", "Slide 05 · Business Model Options 1–2", [
  { tag: "OPTION 1", title: "Independent Concierge", body: "We operate as a standalone entity mirroring your catalog via API. Senders pay us in US$; we buy stock at a negotiated wholesale discount and fulfil with our own driver network.", setup: "US$15,000", opex: "US$8,500", rev: "5–8% markup + 3–5% rebate" },
  { tag: "OPTION 2", title: "White-Label Licensing", body: "We build the cross-border storefront extension and license it to TM PnP. It integrates into tmpnponline.co.zw, natively branded, fulfilled by third-party Zimbabwean couriers.", setup: "US$25,000", opex: "US$2,000", rev: "US$40k + US$3.5k/mo or 1.5–2% GMV" },
]);

optionSlide(7, "Partnership and hybrid structures", "Slide 05 · Business Model Options 3–4", [
  { tag: "OPTION 3", title: "Tripartite Joint Venture", body: "PnP + Quatrohaus / Code Virtus + our firm. We provide the cross-border framework, payment tunnels and diaspora marketing; branch staff pick, third-party bike networks deliver.", setup: "US$5,000", opex: "Absorbed in IT budget", rev: "60 / 20 / 20 net profit split" },
  { tag: "OPTION 4", title: "Remittance Tokens", body: "Diaspora shoppers buy a US$ Eco-Voucher: auto-dispatch a staple hamper by bike courier, or push a secure barcode via SMS/WhatsApp for in-store redemption, bypassing cash-out queues.", setup: "US$10,000", opex: "US$3,000", rev: "4–6% commission + float yield" },
]);

/* 9 — Matrix */
{
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  chrome(s, "Slide 06 · Configuration Matrix", 8, true);
  title(s, "Commercial model comparison", true);
  const head = ["Metric", "1 · Reseller", "2 · White-Label", "3 · Joint Venture", "4 · Voucher Hybrid"];
  const rows = [
    ["Upfront capital", "US$15,000", "US$25,000", "US$5,000", "US$10,000"],
    ["Operational effort", "Extremely high", "Very low", "Low (oversight)", "Medium"],
    ["Time to market", "3–4 months", "4–5 months", "1–2 months", "2 months"],
    ["Primary revenue", "Product markups", "Licensing fees", "Net profit pool", "Commission + float"],
    ["TM PnP risk tier", "Zero risk", "Technology buyer", "Co-owner", "Channel partner"],
  ];
  const body = [
    head.map((h) => ({ text: h, options: { fill: { color: RED }, color: WHITE, bold: true, fontSize: 12, fontFace: HEAD } })),
    ...rows.map((r, i) => r.map((c, j) => ({
      text: c,
      options: { fill: { color: i % 2 ? "16304F" : "1B3A6B" }, color: j === 0 ? GOLD : "E3EAF4", bold: j === 0, fontSize: 11, fontFace: BODY },
    }))),
  ];
  s.addTable(body, { x: 0.45, y: 1.95, w: 9.1, colW: [1.9, 1.8, 1.8, 1.8, 1.8], rowH: 0.42, valign: "middle", margin: 0.1, border: { type: "solid", color: BLUE_DEEP, pt: 1 } });
}

/* 10 — Integrations */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 07 · Critical Integrations", 9, false);
  title(s, "Building beyond the current platform", false);
  s.addText("Three new components layer over the existing web framework.", { x: 0.45, y: 1.72, w: 8, h: 0.3, fontFace: BODY, fontSize: 13, color: MUTED });
  const cw = 2.9, ch = 2.6;
  card(s, { x: 0.45, y: 2.15, w: cw, h: ch, heading: "Geo-fenced marketing core", body: "Hyper-targeted social and digital advertising aimed at high-density Zimbabwean pockets — Hillbrow and Randburg in South Africa, London and Leeds in the UK." });
  card(s, { x: 0.45 + cw + 0.2, y: 2.15, w: cw, h: ch, accent: BLUE, heading: "API logistics middleware", body: "A dispatch interface linking the TM PnP back-end to localized courier networks and on-demand e-bike fleets for route optimisation and proof of delivery." });
  card(s, { x: 0.45 + 2 * (cw + 0.2), y: 2.15, w: cw, h: ch, heading: "Real-time substitution logic", body: "An automated WhatsApp bot lets the sender or recipient approve alternatives instantly when an item goes out of stock during picking." });
}

/* 11 — First mover */
{
  const s = pptx.addSlide();
  s.background = { color: RED_DEEP };
  s.addImage({ data: DOOR, x: 5.4, y: 0, w: 4.6, h: H, sizing: { type: "cover", w: 4.6, h: H } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 6.0, h: H, fill: { color: RED_DEEP } });
  chrome(s, "Slide 08 · First-Mover Advantage", 10, true, 5.9);
  s.addText("A defensive moat against OK Zimbabwe & Choppies", { x: 0.45, y: 0.95, w: 5.3, h: 0.85, fontFace: HEAD, fontSize: 25, bold: true, color: WHITE });
  const block = (y, tag, text) => {
    s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y, w: 5.3, h: 1.35, fill: { color: "C21740" }, line: { color: "C21740" }, rectRadius: 0.08 });
    s.addShape(pptx.ShapeType.roundRect, { x: 0.68, y: y + 0.18, w: 1.75, h: 0.3, fill: { color: WHITE }, line: { color: WHITE }, rectRadius: 0.15 });
    s.addText(tag, { x: 0.68, y: y + 0.18, w: 1.75, h: 0.3, align: "center", valign: "middle", fontFace: BODY, fontSize: 9.5, bold: true, color: RED_DEEP, margin: 0 });
    s.addText(text, { x: 0.68, y: y + 0.58, w: 4.85, h: 0.68, fontFace: BODY, fontSize: 11.5, color: "FDE7EC", margin: 0, valign: "top" });
  };
  block(1.95, "Market leadership", "While competitors stay focused on brick-and-mortar or basic localized delivery, TM PnP becomes the definitive cross-border retail pipeline for the diaspora ecosystem.");
  block(3.45, "Maximising group assets", "Collection desks convert into high-volume dispatch stations, lifting stock turnover speed across all primary product lines.");
}

/* 12 — Value proposition */
{
  const VALUE_PROPS = [
    ["White-label solution", "Customizable platform retailers brand as their own, with flexible commercial models — subscription or commission-based."],
    ["Time-to-market advantage", "Leverage existing integrations (API, SIM switch, payments) to launch quickly and capture customers before competitors."],
    ["Robust and secure platform", "A reliable, scalable and secure system compared to informal, unregulated channels."],
    ["Bank-agnostic integration", "Participation across multiple banks, expanding reach and customer access."],
    ["Critical mass creation", "Support in building user adoption through partnerships with banks and diaspora communities."],
    ["Hybrid commercial model", "Transaction fees, subscription or discounts — flexibility for different business strategies."],
    ["Adaptability to market dynamics", "Analytics to adjust pricing, packaging and delivery models as consumer behaviour shifts."],
    ["Direct-to-Consumer", "Moving beyond brick-and-mortar retail into a direct relationship with the shopper."],
  ];
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 09 · Value Proposition", 11, false);
  title(s, "What the platform uniquely delivers", false);
  const cw = 2.2, ch = 1.48, gx = 0.24, gy = 0.18, x0 = 0.45, y0 = 1.8;
  VALUE_PROPS.forEach(([h, b], i) => {
    const x = x0 + (i % 4) * (cw + gx);
    const y = y0 + Math.floor(i / 4) * (ch + gy);
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: cw, h: ch, fill: { color: WHITE }, line: { color: "E2E5EB" }, rectRadius: 0.08 });
    const accent = i % 2 ? BLUE : RED;
    s.addShape(pptx.ShapeType.roundRect, { x: x + 0.2, y: y + 0.18, w: 0.4, h: 0.055, fill: { color: accent }, line: { color: accent }, rectRadius: 0.03 });
    s.addText(h, { x: x + 0.2, y: y + 0.32, w: cw - 0.4, h: 0.42, fontFace: HEAD, fontSize: 11.5, bold: true, color: BLUE, margin: 0, valign: "top" });
    s.addText(b, { x: x + 0.2, y: y + 0.74, w: cw - 0.4, h: ch - 0.9, fontFace: BODY, fontSize: 8.5, color: MUTED, margin: 0, valign: "top" });
  });
}

/* 13 — Benefits */
{
  const BENEFITS = [
    ["Faster market entry", "Existing developer integrations accelerate rollout."],
    ["Customer convenience", "Digital shift lets diaspora and local customers buy from anywhere."],
    ["Expanded customer base", "Diaspora markets plus multiple banks' captured audiences."],
    ["Replacement of informal trading", "A legitimate, tax-compliant alternative to unreliable channels."],
    ["Cross-border resilience", "Bypasses constraints on physical goods movement between countries."],
    ["Economies of scale", "Better prices and smaller packages matched to consumer cash flow."],
    ["Increased loyalty & retention", "Secure, reliable delivery builds trust vs informal traders."],
    ["Revenue diversification", "Subscriptions, fees or discounts plus wholesale partnerships."],
    ["Scalability", "Regional expansion with packaging adapted to local needs."],
    ["Convenience-driven adoption", "Small daily-use packages fit township and village habits."],
    ["Government alignment", "Formalized trade supports taxation and regulation."],
  ];
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  chrome(s, "Slide 10 · Benefits", 12, true);
  title(s, "Benefits to TM Pick n Pay and the market", true);
  const cw = 2.98, ch = 0.72, gx = 0.18, gy = 0.1, x0 = 0.45, y0 = 1.78;
  BENEFITS.forEach(([h, b], i) => {
    const x = x0 + (i % 3) * (cw + gx);
    const y = y0 + Math.floor(i / 3) * (ch + gy);
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: cw, h: ch, fill: { color: "1E426F" }, line: { color: "1E426F" }, rectRadius: 0.07 });
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.18, y: y + 0.16, w: 0.12, h: 0.12, fill: { color: GOLD }, line: { color: GOLD } });
    s.addText(h, { x: x + 0.4, y: y + 0.1, w: cw - 0.6, h: 0.24, fontFace: HEAD, fontSize: 10.5, bold: true, color: WHITE, margin: 0, valign: "top" });
    s.addText(b, { x: x + 0.4, y: y + 0.34, w: cw - 0.6, h: 0.36, fontFace: BODY, fontSize: 8.5, color: "C6D3E4", margin: 0, valign: "top" });
  });
}

/* ---------------- Downstream innovation slides ---------------- */

function panel(s, { x, y, w, h, heading, body, fill = "1E426F", head = WHITE, text = "D5DEEA" }) {
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08 });
  s.addText(heading, { x: x + 0.22, y: y + 0.18, w: w - 0.44, h: 0.42, fontFace: HEAD, fontSize: 13, bold: true, color: head, margin: 0, valign: "top" });
  s.addText(body, { x: x + 0.22, y: y + 0.62, w: w - 0.44, h: h - 0.8, fontFace: BODY, fontSize: 10, color: text, margin: 0, valign: "top" });
}

/* 15 — Why now */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 14 · Why Now", 15, false);
  title(s, "Sell market access and intelligence — not “a diaspora solution”", false);
  const items = [
    ["Brick-and-mortar is no longer defensible", "Footfall retail cannot hold share on its own. The defensible asset is the customer relationship, not the store estate."],
    ["Buyers are price-led, not brand-led", "Households shop the cheapest basket available on the day. Informal channels are winning that comparison by default."],
    ["The Malayitsha channel is being constricted", "One-stop border posts, electronic travel authority and tighter vehicle declarations are squeezing the informal cross-border van trade."],
    ["The opening left behind", "Diaspora funding + bank rails + owned delivery + wholesale into the informal channel — a brand that lives in people's homes."],
  ];
  items.forEach(([h, b], i) => {
    card(s, { x: 0.45 + (i % 2) * 4.68, y: 1.85 + Math.floor(i / 2) * 1.35, w: 4.42, h: 1.22, heading: h, body: b, accent: i % 2 ? BLUE : RED });
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 4.6, w: 9.1, h: 0.62, fill: { color: BLUE }, line: { color: BLUE }, rectRadius: 0.08 });
  s.addText("Positioning discipline: the pitch is market access and demand intelligence. Framed as a diaspora product, it reads as a niche remittance play and stalls in committee.", { x: 0.7, y: 4.6, w: 8.6, h: 0.62, valign: "middle", fontFace: BODY, fontSize: 10.5, color: WHITE, margin: 0 });
}

/* 16 — Retail-agnostic end state */
{
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  chrome(s, "Slide 15 · Downstream Innovation", 16, true);
  title(s, "The end state is a trading platform, not a single-retailer app", true);
  const phases = [
    ["Phase 1 · Entry", "TM Pick n Pay value-add", "Single-tenant. TM branded, TM catalogue, TM fulfilment. The platform earns its place inside one retailer before anything else is discussed."],
    ["Phase 2 · Depth", "Wholesale and informal channel", "TM supplies tuck shops and repackagers through the same rails. Volume grows without adding a competing banner to the storefront."],
    ["Phase 3 · End state", "Retail-agnostic trading platform", "Shopper asks for cooking oil; the request goes to every participating supplier and the best price wins. Retailers become suppliers on a marketplace."],
  ];
  phases.forEach(([k, t, d], i) => {
    const x = 0.45 + i * 3.1;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.85, w: 2.9, h: 2.0, fill: { color: "1E426F" }, line: { color: "1E426F" }, rectRadius: 0.08 });
    s.addShape(pptx.ShapeType.roundRect, { x: x + 0.22, y: 2.02, w: 1.5, h: 0.3, fill: { color: GOLD }, line: { color: GOLD }, rectRadius: 0.15 });
    s.addText(k, { x: x + 0.22, y: 2.02, w: 1.5, h: 0.3, align: "center", valign: "middle", fontFace: BODY, fontSize: 9, bold: true, color: INK, margin: 0 });
    s.addText(t, { x: x + 0.22, y: 2.42, w: 2.46, h: 0.4, fontFace: HEAD, fontSize: 12.5, bold: true, color: WHITE, margin: 0 });
    s.addText(d, { x: x + 0.22, y: 2.84, w: 2.46, h: 0.9, fontFace: BODY, fontSize: 9.5, color: "C6D3E4", margin: 0, valign: "top" });
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 4.05, w: 4.45, h: 1.05, fill: { color: RED }, line: { color: RED }, rectRadius: 0.08 });
  s.addText("Multi-tenant has already been prototyped — three banners running in one app — then deliberately pulled back to a single tenant for the entry conversation.", { x: 0.68, y: 4.05, w: 4.0, h: 1.05, valign: "middle", fontFace: BODY, fontSize: 10, color: "FDE7EC", margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 5.1, y: 4.05, w: 4.45, h: 1.05, fill: { color: "1E426F" }, line: { color: "1E426F" }, rectRadius: 0.08 });
  s.addText("Sequencing is the safeguard: introduce the platform as a TM advantage, without signalling that competitors sit on the same rails on day one.", { x: 5.33, y: 4.05, w: 4.0, h: 1.05, valign: "middle", fontFace: BODY, fontSize: 10, color: "C6D3E4", margin: 0 });
}

/* 17 — Price comparison engine */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 16 · Price Comparison Engine", 17, false);
  title(s, "Item-level price comparison is the reason to participate", false);
  s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 1.85, w: 5.0, h: 3.25, fill: { color: WHITE }, line: { color: "E2E5EB" }, rectRadius: 0.08 });
  s.addText("SHOPPER REQUEST · “COOKING OIL”", { x: 0.7, y: 2.02, w: 4.5, h: 0.28, fontFace: BODY, fontSize: 9, bold: true, charSpacing: 1.5, color: MUTED, margin: 0 });
  const rows = [
    ["Supplier A · 2L cooking oil", "US$3.10", "Best price · 2.1km"],
    ["Supplier B · 2L cooking oil", "US$3.45", "Faster slot · 0.8km"],
    ["Supplier C · 2L cooking oil", "US$3.60", "Bundle offer"],
    ["Tuck shop · 500ml repack", "US$0.95", "Local fulfilment"],
  ];
  rows.forEach(([item, price, note], i) => {
    const y = 2.4 + i * 0.66;
    const best = i === 0;
    s.addShape(pptx.ShapeType.roundRect, { x: 0.7, y, w: 4.5, h: 0.56, fill: { color: best ? BLUE : PAPER }, line: { color: best ? BLUE : "E2E5EB" }, rectRadius: 0.07 });
    s.addText(item, { x: 0.9, y: y + 0.07, w: 2.9, h: 0.24, fontFace: BODY, fontSize: 10, bold: true, color: best ? WHITE : INK, margin: 0 });
    s.addText(note, { x: 0.9, y: y + 0.3, w: 2.9, h: 0.22, fontFace: BODY, fontSize: 8.5, color: best ? "C6D3E4" : MUTED, margin: 0 });
    s.addText(price, { x: 3.85, y, w: 1.15, h: 0.56, align: "right", valign: "middle", fontFace: HEAD, fontSize: 14, bold: true, color: best ? WHITE : INK, margin: 0 });
  });
  card(s, { x: 5.65, y: 1.85, w: 3.9, h: 1.55, heading: "Comparison across every participating supplier", body: "One request fans out to all listed suppliers. The shopper sees price, distance and slot side by side and picks. Price competition, not shelf position, decides the sale." });
  card(s, { x: 5.65, y: 3.55, w: 3.9, h: 1.55, accent: BLUE, heading: "AI ranking of most-tradable items", body: "The engine ranks demand by item, area and week — a what-to-stock signal feeding order-to-cash: buy the right depth, in the right pack size, before the demand lands." });
}

/* 18 — Subscription is the model */
{
  const s = pptx.addSlide();
  s.background = { color: RED_DEEP };
  chrome(s, "Slide 17 · The Revenue Model", 18, true);
  title(s, "Anything that allows a subscription is the model", true);
  const subs = [
    ["Rider fees", "Scooter rent-to-buy instalments, then a standing platform fee per dollar earned."],
    ["Advertising", "Scooter livery, in-app placement and supplier media sold as monthly inventory."],
    ["Garage services", "Servicing, parts and accessories on a maintenance plan, not a per-repair invoice."],
    ["Tenant fees", "Each retailer, wholesaler and tuck shop pays a tiered monthly platform fee."],
    ["Shopper plans", "Priority delivery and family basket sharing on a recurring plan."],
    ["Data products", "Demand and price intelligence licensed by subscription to suppliers."],
  ];
  subs.forEach(([h, b], i) => {
    panel(s, { x: 0.45 + (i % 3) * 3.1, y: 1.85 + Math.floor(i / 3) * 1.3, w: 2.9, h: 1.18, heading: h, body: b, fill: "C21740", text: "FDE7EC" });
  });
  s.addText("Every line above is bent toward recurring revenue. One-off transaction fees fund a project; recurring fees fund a company and give the valuation a base to stand on.", { x: 0.45, y: 4.55, w: 9.1, h: 0.6, fontFace: BODY, fontSize: 11.5, color: "FDE7EC", margin: 0 });
}

/* 19 — Scooter economics */
{
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  chrome(s, "Slide 18 · Owned Delivery Network", 19, true);
  title(s, "The last mile is owned, not outsourced", true);
  const stats = [["500–2,000", "Platform-owned electric scooters"], ["12 months", "Rent-to-buy, then the rider owns it"], ["~5 months", "Asset pays itself back"], ["~10%", "Platform fee per dollar earned after ownership"]];
  stats.forEach(([v, l], i) => {
    const x = 0.45 + i * 2.31;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.8, w: 2.11, h: 1.05, fill: { color: "1E426F" }, line: { color: "1E426F" }, rectRadius: 0.08 });
    s.addText(v, { x: x + 0.2, y: 1.9, w: 1.75, h: 0.45, fontFace: HEAD, fontSize: 20, bold: true, color: GOLD, margin: 0 });
    s.addText(l, { x: x + 0.2, y: 2.35, w: 1.75, h: 0.42, fontFace: BODY, fontSize: 9, color: "C6D3E4", margin: 0, valign: "top" });
  });
  const blocks = [
    ["Rent-to-buy", "The rider operates for twelve months and then owns the scooter. The asset repays itself in roughly five months; months six to twelve are margin."],
    ["Advertising rights", "Livery is sold on subscription: 80% reserved for the anchor retailer, 20% open inventory. A moving media network across every delivery route."],
    ["Own the garage", "Repairs, parts and accessories stay inside the platform on a maintenance plan. Electric tricycles extend the fleet into rural routes, with women riders as the first cohort."],
  ];
  blocks.forEach(([h, b], i) => {
    panel(s, { x: 0.45 + i * 3.1, y: 3.0, w: 2.9, h: 1.55, heading: h, body: b });
  });
  s.addImage({ data: BIKE, x: 0.45, y: 4.68, w: 9.1, h: 0.42, sizing: { type: "cover", w: 9.1, h: 0.42 } });
}

/* 20 — Loyalty */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 19 · Loyalty & Device Migration", 20, false);
  title(s, "Trade US$100 a month for five months — earn the handset", false);
  ["Month 1", "Month 2", "Month 3", "Month 4", "Month 5"].forEach((m, i) => {
    const x = 0.45 + i * 1.85;
    const last = i === 4;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.85, w: 1.7, h: 1.0, fill: { color: last ? RED : WHITE }, line: { color: last ? RED : "E2E5EB" }, rectRadius: 0.08 });
    s.addText(m, { x, y: 2.02, w: 1.7, h: 0.32, align: "center", fontFace: HEAD, fontSize: 13, bold: true, color: last ? WHITE : INK, margin: 0 });
    s.addText(last ? "5G-lite smartphone awarded" : "US$100 traded", { x: x + 0.15, y: 2.36, w: 1.4, h: 0.42, align: "center", fontFace: BODY, fontSize: 9, color: last ? "FDE7EC" : MUTED, margin: 0 });
  });
  const cards = [
    ["Discovery-Vitality mechanics", "Consecutive qualifying months unlock the reward. Behaviour is earned, not bought, and the tier resets if trading lapses.", BLUE],
    ["Family basket sharing", "Several household members contribute to one qualifying basket, so families reach the threshold faster and consolidate spend on the platform.", RED],
    ["The real purpose", "A ~US$50 landed handset with the markup absorbed migrates 2G and 3G users onto the platform. Device cost is customer acquisition, booked as such.", BLUE],
  ];
  cards.forEach(([h, b, a], i) => {
    card(s, { x: 0.45 + i * 3.1, y: 3.05, w: 2.9, h: 2.0, heading: h, body: b, accent: a });
  });
}

/* 21 — Bank-agnostic rails */
{
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  chrome(s, "Slide 20 · Payment Rails", 21, true);
  title(s, "Bank-agnostic by design, multi-bank in practice", true);
  panel(s, { x: 0.45, y: 1.85, w: 4.45, h: 2.35, heading: "What we say to banks", body: "Plug your remittance APIs into the platform and we bring transaction share-of-wallet: recurring diaspora flows landing as retail settlement rather than cash-out.\n\nNo exclusivity is requested and none is given. Every rail is one integration among several, so pricing stays competitive and no single institution gates the platform." });
  panel(s, { x: 5.1, y: 1.85, w: 4.45, h: 2.35, heading: "What we say to the retailer", body: "Banks bring critical mass. Their diaspora bases are already captured audiences; the platform converts those balances into baskets inside your estate.\n\nInstitutions under discussion are prospects at this stage. Nothing is presented in-app as a confirmed partnership until an agreement is signed." });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.45, y: 4.35, w: 9.1, h: 0.72, fill: { color: RED }, line: { color: RED }, rectRadius: 0.08 });
  s.addText("Design rule: no bank-specific logic in the core. Rails are adapters, so adding or dropping an institution is a configuration change, never a rebuild.", { x: 0.7, y: 4.35, w: 8.6, h: 0.72, valign: "middle", fontFace: BODY, fontSize: 11, color: WHITE, margin: 0 });
}

/* 22 — Wholesale and tuck shops */
{
  const s = pptx.addSlide();
  s.background = { color: PAPER };
  chrome(s, "Slide 21 · Wholesale & Informal Channel", 22, false);
  title(s, "Tuck shops become distribution extensions, not competitors", false);
  const steps = [
    ["Aggregate", "Local orders in a suburb are pooled in the app rather than fragmented across trips."],
    ["Fulfil locally", "The nearest tuck shop picks and hands over, cutting distance, fuel and delivery cost."],
    ["Repackage", "Bulk stock is broken into the 200g-type pack sizes the market actually buys."],
    ["Supply", "The retailer sells bulk into the channel as a wholesaler and keeps the volume."],
  ];
  steps.forEach(([h, b], i) => {
    const x = 0.45 + i * 2.31;
    const dark = i % 2 === 0;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.85, w: 2.11, h: 1.4, fill: { color: dark ? BLUE : WHITE }, line: { color: dark ? BLUE : "E2E5EB" }, rectRadius: 0.08 });
    s.addText(`STEP ${i + 1}`, { x: x + 0.2, y: 1.97, w: 1.75, h: 0.22, fontFace: BODY, fontSize: 8, bold: true, charSpacing: 1.5, color: dark ? "AEBED2" : MUTED, margin: 0 });
    s.addText(h, { x: x + 0.2, y: 2.19, w: 1.75, h: 0.3, fontFace: HEAD, fontSize: 12.5, bold: true, color: dark ? WHITE : BLUE, margin: 0 });
    s.addText(b, { x: x + 0.2, y: 2.5, w: 1.75, h: 0.65, fontFace: BODY, fontSize: 8.5, color: dark ? "C6D3E4" : MUTED, margin: 0, valign: "top" });
  });
  card(s, { x: 0.45, y: 3.4, w: 4.45, h: 1.6, heading: "Pack size is the unlock", body: "Households buy 200g, not 2kg. The retailer sells bulk to the tuck shop, the tuck shop repackages to the price point the street can afford — the small-multipack playbook proven by packaged brands expanding regionally." });
  card(s, { x: 5.1, y: 3.4, w: 4.45, h: 1.6, accent: BLUE, heading: "The government argument", body: "Routing informal trade through the platform makes it visible: tax is captured, stock is traceable and counterfeit product is squeezed out. Legitimisation, not enforcement." });
}

/* 23 — Commercial and entity structure */
{
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  chrome(s, "Slide 22 · Commercial & Entity Structure", 23, true);
  title(s, "Hybrid commercial model, two-country structure", true);
  const model = [
    ["White-label", "The storefront and app ship under the retailer's brand, licensed rather than sold."],
    ["Subscription", "Tiered monthly platform fees per tenant, opening at a 100,000-participant tier."],
    ["Commission", "A thin transaction share on GMV, sitting on top of the recurring base."],
  ];
  model.forEach(([h, b], i) => {
    panel(s, { x: 0.45 + i * 3.1, y: 1.85, w: 2.9, h: 1.25, heading: h, body: b });
  });
  panel(s, { x: 0.45, y: 3.25, w: 4.45, h: 1.3, heading: "South African company", body: "Holds the platform IP and raises capital where liquidity and investor access are deepest. Contracts with tenants and banks sit here." });
  panel(s, { x: 5.1, y: 3.25, w: 4.45, h: 1.3, heading: "Zimbabwean subsidiary", body: "Runs logistics, the scooter fleet and local agreements on the ground, with local employment and local regulatory standing." });
  s.addText("Directional as at 13 August — structure to be confirmed with tax and legal counsel before any binding term sheet.", { x: 0.45, y: 4.68, w: 9.1, h: 0.32, fontFace: BODY, fontSize: 9.5, color: "8C9BB0", margin: 0 });
}

/* 24 — Close */


{
  const s = pptx.addSlide();
  s.background = { color: BLUE_DEEP };
  s.addShape(pptx.ShapeType.roundRect, { x: 4.1, y: 1.05, w: 1.8, h: 0.5, fill: { color: WHITE }, line: { color: WHITE }, rectRadius: 0.07 });
  s.addImage({ data: LOGO, x: 4.2, y: 1.15, w: 1.6, h: 0.31 });
  s.addText("Turn 57+ branches into a diaspora fulfilment network", { x: 1.0, y: 1.95, w: 8, h: 1.0, align: "center", valign: "top", fontFace: HEAD, fontSize: 30, bold: true, color: WHITE });
  s.addShape(pptx.ShapeType.roundRect, { x: 4.4, y: 3.25, w: 1.2, h: 0.08, fill: { color: RED }, line: { color: RED }, rectRadius: 0.04 });
  s.addText("Recommended next step: select a commercial structure and mandate a 60-day pilot on two flagship Harare branches.", { x: 1.6, y: 3.6, w: 6.8, h: 0.9, align: "center", fontFace: BODY, fontSize: 15, color: "C6D3E4" });
  s.addText("TM Pick n Pay Express", { x: 0.45, y: H - 0.5, w: 4, h: 0.3, fontFace: BODY, fontSize: 9, color: "8C9BB0" });
  s.addText("Thank you", { x: W - 4.45, y: H - 0.5, w: 4, h: 0.3, align: "right", fontFace: BODY, fontSize: 9, color: "8C9BB0" });
}

await pptx.writeFile({ fileName: "/tmp/deck/TM-Pick-n-Pay-Express.pptx" });
console.log("written");
