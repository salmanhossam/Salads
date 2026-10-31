// Builds noya-or-rooms/Noya_OR_Patient_Journey_and_Branding.pptx
// Run: NODE_PATH=<deps>/node_modules node build_deck.js <skillDir> <outFile>
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(require("path").join(process.argv[2], "scripts/apply_theme.js"));
const OUT = process.argv[3];

// Placeholder palette until Noya brand guide is supplied
const HEX = {
  deep: "0F3D47", teal: "1F7A74", mint: "E4F2EF", sand: "F2E8DC",
  coral: "D9734E", ink: "1D2B30", grey: "5F6F73", white: "FFFFFF",
};
const THEME = {
  name: "Noya OR",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.deep, lt2: HEX.mint,
    accent1: HEX.teal, accent2: HEX.coral, accent3: "6B3E75", accent4: "1F3A5F",
    accent5: "8A9199", accent6: HEX.sand, hlink: HEX.teal, folHlink: "6B3E75",
  },
};

async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.title = "Noya Clinics - OR Patient Journey and Branding";
  pres.author = "Noya Clinics";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  // ---------- Layouts ----------
  pres.defineSlideMaster({
    title: "Title Dark",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.5, w: 8.8, h: 1.6, fontSize: 40, align: "left", bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.3, w: 8.8, h: 1.0, fontSize: 18, color: C.background2, valign: "top", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "Section Dark",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.0, w: 8.8, h: 1.0, fontSize: 40, align: "left", bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.1, w: 8.8, h: 0.8, fontSize: 16, color: C.background2, valign: "top", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 9.0, y: 5.15, w: 0.5, h: 0.3, fontSize: 10, color: C.background2 },
  });
  pres.defineSlideMaster({
    title: "Content",
    background: { color: C.background1 },
    margin: [0.5, 0.5, 0.5, 0.5],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.75, fontSize: 30, align: "left", bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
      { text: { text: "Noya Clinics  |  OR Launch", options: { x: 0.5, y: 5.15, w: 5, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
    ],
    slideNumber: { x: 9.0, y: 5.15, w: 0.5, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
  });

  let n = 0;
  const T = (s, txt, o) => s.addText(txt, Object.assign({ isTextBox: true, margin: 0, objectName: "t" + (++n) }, o));
  const card = (s, x, y, w, h, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.background2 }, line: { type: "none" }, objectName: "card" + (++n),
  });
  async function iconDot(s, x, y, d, name, bg, fg) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg || C.accent1 }, line: { type: "none" }, objectName: "dot" + (++n) });
    const p = d * 0.28;
    s.addImage({ data: await icon(name, fg || HEX.white), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: "ic" + n });
  }
  const content = (sec, title) => { const s = pres.addSlide({ masterName: "Content", sectionTitle: sec }); s.addText(title, { placeholder: "title" }); return s; };
  const section = (sec, title, sub) => { pres.addSection({ title: sec }); const s = pres.addSlide({ masterName: "Section Dark", sectionTitle: sec }); s.addText(title, { placeholder: "title" }); s.addText(sub, { placeholder: "body" }); return s; };

  // ---------- 1. Title ----------
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "Title Dark", sectionTitle: "Opening" });
  s.addText("Noya Operating Rooms", { placeholder: "title" });
  s.addText("Patient journey, OR branding & uniform system", { placeholder: "body" });
  T(s, "Colours shown are placeholders until the Noya brand guide is applied", { x: 0.6, y: 4.9, w: 8.8, h: 0.3, fontSize: 11, color: C.background2, italic: true });
  s.addNotes("Purpose: align management on the patient journey for the new ORs, how we brand the space within MoH infection-control rules, and the uniform system.");

  // ---------- 2. Big idea ----------
  s = content("Opening", "Not printed doesn't mean not branded");
  T(s, "Patients spend about 10 awake minutes in the OR. We brand those minutes through people, colour, light and sound, not logos.", { x: 0.5, y: 1.15, w: 9, h: 0.6, fontSize: 16, color: C.text1 });
  const pillars = [
    ["FaRoute" in fa ? "FaRoute" : "FaChartLine", "Patient journey", "13 stages from first enquiry to recommendation, plus a parallel journey for the companion"],
    ["FaLayerGroup", "Branded space", "Three zones: full branding in public areas, colour and light only inside the OR"],
    ["FaTshirt", "Uniform system", "Role colours and printed names on reusable wear. No print on single-use items"],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.5 + i * 3.1;
    card(s, x, 2.0, 2.8, 2.85);
    await iconDot(s, x + 0.3, 2.25, 0.7, pillars[i][0]);
    T(s, pillars[i][1], { x: x + 0.3, y: 3.1, w: 2.3, h: 0.45, fontSize: 18, bold: true, color: C.text2, fontFace: "Cambria" });
    T(s, pillars[i][2], { x: x + 0.3, y: 3.55, w: 2.3, h: 1.15, fontSize: 14, color: C.text1, valign: "top" });
  }

  // ---------- 3. Ground rules ----------
  s = content("Opening", "The rules we design within");
  card(s, 0.5, 1.2, 4.35, 3.7);
  await iconDot(s, 0.75, 1.4, 0.6, "FaShieldAlt");
  T(s, "Saudi MoH infection control", { x: 1.5, y: 1.45, w: 3.2, h: 0.5, fontSize: 18, bold: true, color: C.text2, fontFace: "Cambria" });
  T(s, [
    { text: "OR walls, floor and ceiling seamless, with no decorative parts", options: { bullet: true, breakLine: true } },
    { text: "Only necessary items inside the operating theatre", options: { bullet: true, breakLine: true } },
    { text: "Scrubs facility-laundered at 71 °C for 25 min", options: { bullet: true, breakLine: true } },
    { text: "Long-sleeved warm-up jacket required in the OR", options: { bullet: true, breakLine: true } },
    { text: "Patients' hair covered once inside the surgical suite", options: { bullet: true } },
  ], { x: 0.75, y: 2.15, w: 3.95, h: 2.6, fontSize: 14, color: C.text1, valign: "top", paraSpaceAfter: 6 });
  card(s, 5.15, 1.2, 4.35, 3.7, C.accent6);
  await iconDot(s, 5.4, 1.4, 0.6, "FaBan", C.accent2);
  T(s, "Noya MoH agreement", { x: 6.15, y: 1.45, w: 3.2, h: 0.5, fontSize: 18, bold: true, color: C.text2, fontFace: "Cambria" });
  T(s, "Single-use wearables are never printed.", { x: 5.4, y: 2.15, w: 3.9, h: 0.5, fontSize: 16, bold: true, color: C.text1 });
  T(s, "Everything reusable can carry print, as long as it survives the 71 °C wash. That makes reusable items our main branding canvas.", { x: 5.4, y: 2.7, w: 3.9, h: 1.3, fontSize: 14, color: C.text1, valign: "top" });
  T(s, "Sources: MoH OR IPC Guideline v1.1 (2022); MoH IPC Auditors' Guide v6 (2025)", { x: 5.4, y: 4.25, w: 3.9, h: 0.5, fontSize: 10, color: C.text1, italic: true });

  // ---------- 4. Section: journey ----------
  section("Patient journey", "The patient journey", "From first enquiry to recommendation: 13 stages, one named coordinator throughout");

  // ---------- 5. Journey overview ----------
  s = content("Patient journey", "13 stages, 3 phases");
  const phases = [
    ["Before", ["Discover", "Consult", "Book & prepare", "D-1 call"], C.accent1],
    ["Day of surgery", ["Arrive", "Admit", "Pre-op", "Transfer", "In the OR", "Recovery"], C.text2],
    ["After", ["Discharge", "Follow-up", "Refer"], C.accent2],
  ];
  let k = 1, xs = 0.5;
  const unit = 9.0 / 13;
  for (const [name, steps, col] of phases) {
    const w = steps.length * unit;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xs + 0.04, y: 1.25, w: w - 0.08, h: 0.5, rectRadius: 0.1, fill: { color: col }, line: { type: "none" }, objectName: "ph" + k });
    T(s, name, { x: xs + 0.04, y: 1.25, w: w - 0.08, h: 0.5, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    steps.forEach((st, i) => {
      const cx = xs + i * unit + unit / 2;
      s.addShape(pres.shapes.OVAL, { x: cx - 0.24, y: 2.05, w: 0.48, h: 0.48, fill: { color: C.background1 }, line: { color: col, width: 2 }, objectName: "st" + k });
      T(s, String(k), { x: cx - 0.24, y: 2.05, w: 0.48, h: 0.48, fontSize: 14, bold: true, color: col, align: "center", valign: "middle" });
      T(s, st, { x: cx - unit / 2, y: 2.62, w: unit, h: 0.6, fontSize: 11, color: C.text1, align: "center", valign: "top" });
      k++;
    });
    xs += w;
  }
  card(s, 0.5, 3.45, 9.0, 1.45);
  const emo = [["Peak anxiety", "The night before and in pre-op. We answer with a D-1 call, the \"meet your team\" message and a warm pre-op area."],
               ["Loss of control", "Transfer and induction. We answer with the Noya Moment: names, warmth, a hand to hold."],
               ["Fear of abandonment", "At home after discharge. We answer with calls at D+1, D+7 and D+30 and a 24/7 number."]];
  emo.forEach((e, i) => {
    const x = 0.75 + i * 2.95;
    T(s, e[0], { x, y: 3.6, w: 2.7, h: 0.35, fontSize: 14, bold: true, color: C.accent2 });
    T(s, e[1], { x, y: 3.95, w: 2.7, h: 0.9, fontSize: 11, color: C.text1, valign: "top" });
  });

  // ---------- 6. Before ----------
  s = content("Patient journey", "Before surgery: earn trust, remove uncertainty");
  const before = [
    ["FaSearch", "1  Discover", "OR virtual tour, meet the team, safety standards. Ads need MoH approval and make no outcome claims"],
    ["FaUserMd", "2  Consult", "Surgeon consult and a named Surgical Coordinator from here on. Written all-inclusive quote"],
    ["FaClipboardCheck", "3  Book & prepare", "Labs and anaesthesia assessment in one visit, consent explained in plain Arabic, insurance handled by Noya"],
    ["FaPhoneAlt", "4  D-1 call", "Arrival time, fasting, parking, companion. Plus a \"meet your team\" message with photos"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + (i % 2) * 4.6, y = 1.2 + Math.floor(i / 2) * 1.9;
    card(s, x, y, 4.4, 1.7);
    await iconDot(s, x + 0.25, y + 0.3, 0.65, before[i][0]);
    T(s, before[i][1], { x: x + 1.1, y: y + 0.2, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
    T(s, before[i][2], { x: x + 1.1, y: y + 0.6, w: 3.1, h: 1.0, fontSize: 12, color: C.text1, valign: "top" });
  }

  // ---------- 7. Day of ----------
  s = content("Patient journey", "Day of surgery: expected, warm, informed");
  const day = [
    ["FaDoorOpen", "5 Arrive", "Greeted by name and walked to the floor. Companion gets a case code"],
    ["FaIdBadge", "6 Admit", "Private admission. Reusable Noya gown and cap"],
    ["FaBed", "7 Pre-op", "Warm blanket, ceiling art, music of choice, site marked with the patient"],
    ["FaProcedures", "8 Transfer", "A named nurse talks to the patient all the way. No clutter in view"],
    ["FaHandHoldingHeart", "9 In the OR", "The Noya Moment (next slide)"],
    ["FaHeartbeat", "10 Recovery", "Scripted first words. Family updated within 10 min"],
  ];
  for (let i = 0; i < 6; i++) {
    const x = 0.5 + (i % 3) * 3.05, y = 1.2 + Math.floor(i / 3) * 1.9;
    card(s, x, y, 2.85, 1.7);
    await iconDot(s, x + 0.2, y + 0.2, 0.55, day[i][0], i === 4 ? C.accent2 : C.accent1);
    T(s, day[i][1], { x: x + 0.9, y: y + 0.25, w: 1.85, h: 0.45, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
    T(s, day[i][2], { x: x + 0.2, y: y + 0.85, w: 2.5, h: 0.8, fontSize: 12, color: C.text1, valign: "top" });
  }

  // ---------- 8. Noya Moment ----------
  s = content("Patient journey", "The Noya Moment: our signature 10 minutes");
  const moment = [
    ["Set the room", "Warm ambient light, patient's music playing low"],
    ["Names first", "Every team member wears a printed name cap and introduces themselves"],
    ["Warmth", "Warm blanket on immediately"],
    ["Sign-In together", "WHO checklist read out with the patient"],
    ["A hand to hold", "\"We'll take care of you. See you in recovery.\""],
    ["Switch", "After induction: clinical white light, surgeon's music"],
  ];
  moment.forEach((m, i) => {
    const x = 0.5 + (i % 3) * 3.05, y = 1.2 + Math.floor(i / 3) * 1.85;
    T(s, String(i + 1).padStart(2, "0"), { x, y, w: 0.9, h: 0.7, fontSize: 34, bold: true, color: i === 4 ? C.accent2 : C.accent1, fontFace: "Cambria", valign: "top" });
    T(s, m[0], { x: x + 0.9, y: y + 0.05, w: 1.95, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
    T(s, m[1], { x: x + 0.9, y: y + 0.45, w: 1.95, h: 1.1, fontSize: 12, color: C.text1, valign: "top" });
  });
  T(s, "Costs almost nothing, and supports WHO safety-checklist culture.", { x: 0.5, y: 4.75, w: 9, h: 0.35, fontSize: 13, italic: true, color: C.text2 });

  // ---------- 9. After ----------
  s = content("Patient journey", "After surgery: we stay with them");
  card(s, 0.5, 1.2, 3.6, 3.7, C.text2);
  T(s, "11  Discharge", { x: 0.8, y: 1.4, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: C.background1 });
  T(s, [
    { text: "Surgeon visit before discharge", options: { bullet: true, breakLine: true } },
    { text: "Pack covering wound care, medication schedule and warning signs", options: { bullet: true, breakLine: true } },
    { text: "24/7 number", options: { bullet: true, breakLine: true } },
    { text: "Wheelchair to the car", options: { bullet: true } },
  ], { x: 0.8, y: 1.9, w: 3.1, h: 2.0, fontSize: 13, color: C.background1, valign: "top", paraSpaceAfter: 6 });
  T(s, "13  Refer", { x: 0.8, y: 3.85, w: 3.1, h: 0.35, fontSize: 16, bold: true, color: C.background2 });
  T(s, "Review request only after a good D+30 review", { x: 0.8, y: 4.2, w: 3.1, h: 0.55, fontSize: 12, color: C.background1, valign: "top" });
  T(s, "12  Follow-up timeline", { x: 4.5, y: 1.25, w: 5, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  s.addShape(pres.shapes.LINE, { x: 5.4, y: 2.35, w: 3.4, h: 0, line: { color: C.accent1, width: 2 }, objectName: "tl" });
  const fu = [["D+1", "Nurse call: pain, wound, questions"], ["D+7", "Wound check / suture removal"], ["D+30", "Surgeon review + NPS survey"]];
  fu.forEach((f, i) => {
    const cx = 5.4 + i * 1.7;
    s.addShape(pres.shapes.OVAL, { x: cx - 0.3, y: 2.05, w: 0.6, h: 0.6, fill: { color: C.accent1 }, line: { color: C.background1, width: 2 }, objectName: "fu" + i });
    T(s, f[0], { x: cx - 0.6, y: 1.7, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.text2, align: "center" });
    T(s, f[1], { x: cx - 0.8, y: 2.8, w: 1.6, h: 0.9, fontSize: 12, color: C.text1, align: "center", valign: "top" });
  });
  card(s, 4.5, 3.85, 5.0, 1.05);
  T(s, "Thank-you card signed by the surgical team. Patient stories used only with written consent and within MoH advertising rules.", { x: 4.7, y: 3.95, w: 4.6, h: 0.85, fontSize: 12, color: C.text1, valign: "middle" });

  // ---------- 10. Companion ----------
  s = content("Patient journey", "The second customer: the companion");
  T(s, "Companions wait 2 to 6 hours. That wait shapes how they talk about Noya.", { x: 0.5, y: 1.15, w: 9, h: 0.4, fontSize: 15, color: C.text1 });
  const stat = ["Pre-op", "In surgery", "Recovery", "Ready to see"];
  stat.forEach((st, i) => {
    const x = 0.5 + i * 2.3;
    s.addShape(pres.shapes.CHEVRON, { x, y: 1.85, w: 2.2, h: 0.75, fill: { color: i === 3 ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: "ch" + i });
    T(s, st, { x: x + 0.3, y: 1.85, w: 1.6, h: 0.75, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
  });
  T(s, "Status board in the family lounge, shown by case code (not by name) to protect privacy", { x: 0.5, y: 2.7, w: 9, h: 0.35, fontSize: 12, italic: true, color: C.text2 });
  const comp = [["FaMobileAlt", "WhatsApp/SMS updates", "On arrival, mid-surgery (cases over 2 h), and on reaching recovery"],
                ["FaCouch", "Branded family lounge", "Seating, charging, refreshments, status screen"],
                ["FaComments", "Private debrief", "Surgeon speaks to the family in a private room, never in the corridor"]];
  for (let i = 0; i < 3; i++) {
    const x = 0.5 + i * 3.05;
    card(s, x, 3.25, 2.85, 1.65);
    await iconDot(s, x + 0.2, 3.4, 0.5, comp[i][0]);
    T(s, comp[i][1], { x: x + 0.85, y: 3.4, w: 1.9, h: 0.5, fontSize: 13, bold: true, color: C.text2, valign: "middle" });
    T(s, comp[i][2], { x: x + 0.2, y: 4.0, w: 2.5, h: 0.85, fontSize: 12, color: C.text1, valign: "top" });
  }

  // ---------- 11. Section: space ----------
  section("Branded space", "Branding the space", "Three zones. The closer to the operating table, the quieter the brand");

  // ---------- 12. Zones ----------
  s = content("Branded space", "Three zones, three levels of branding");
  const zones = [
    ["A", "Unrestricted", "Entrance, reception, family lounge, consult rooms, pre-op admission", "Full branding: logo walls, photography, screens, ceiling art over pre-op bays", C.accent1, "100%"],
    ["B", "Semi-restricted", "Corridors to ORs, scrub sinks, sterile storage", "Paint colour, welded floor colour, cleanable wayfinding. No loose decor", "4F9D96", "40%"],
    ["C", "Restricted (OR)", "Operating theatres", "No decor (MoH D-3.3). Brand = paint, floor colour, light, sound, uniforms", C.text2, "10%"],
  ];
  zones.forEach((z, i) => {
    const y = 1.2 + i * 1.25;
    card(s, 0.5, y, 9.0, 1.1);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.65, y: y + 0.15, w: 0.8, h: 0.8, rectRadius: 0.1, fill: { color: z[4] }, line: { type: "none" }, objectName: "zone" + i });
    T(s, z[0], { x: 0.65, y: y + 0.15, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: C.background1, align: "center", valign: "middle", fontFace: "Cambria" });
    T(s, z[1], { x: 1.65, y: y + 0.12, w: 2.4, h: 0.4, fontSize: 15, bold: true, color: C.text2 });
    T(s, z[2], { x: 1.65, y: y + 0.5, w: 2.5, h: 0.55, fontSize: 11, color: C.text1, valign: "top" });
    T(s, z[3], { x: 4.35, y: y + 0.15, w: 3.9, h: 0.8, fontSize: 12, color: C.text1, valign: "middle" });
    T(s, z[5], { x: 8.3, y: y + 0.15, w: 1.05, h: 0.5, fontSize: 22, bold: true, color: z[4], align: "right", fontFace: "Cambria" });
    T(s, "brand intensity", { x: 8.0, y: y + 0.62, w: 1.35, h: 0.3, fontSize: 10, color: C.accent5, align: "right" });
  });

  // ---------- 13. Inside OR ----------
  s = content("Branded space", "Inside the OR: brand without decoration");
  const yes = [["FaPalette", "Paint colour", "Neutral walls + one brand-colour wall behind the team, in the MoH seamless epoxy-acrylic coating"],
               ["FaLayerGroup", "Floor colour", "Welded antistatic vinyl with a colour border around the table zone"],
               ["FaLightbulb", "Light", "Warm ambient light on patient entry, clinical white during surgery"],
               ["FaMusic", "Sound", "Licensed Noya calm playlist or the patient's own choice"]];
  for (let i = 0; i < 4; i++) {
    const y = 1.2 + i * 0.93;
    await iconDot(s, 0.5, y, 0.6, yes[i][0]);
    T(s, yes[i][1], { x: 1.25, y: y - 0.02, w: 4.2, h: 0.32, fontSize: 14, bold: true, color: C.text2 });
    T(s, yes[i][2], { x: 1.25, y: y + 0.3, w: 4.2, h: 0.55, fontSize: 11, color: C.text1, valign: "top" });
  }
  card(s, 5.8, 1.2, 3.7, 3.7, C.accent6);
  T(s, "Not in the OR", { x: 6.05, y: 1.35, w: 3.2, h: 0.4, fontSize: 16, bold: true, color: C.accent2 });
  T(s, [
    { text: "Wall graphics, glass art panels, logos", options: { bullet: true, breakLine: true } },
    { text: "Stickers or wraps on equipment", options: { bullet: true, breakLine: true } },
    { text: "Print on sterile drapes, gowns or wraps", options: { bullet: true, breakLine: true } },
    { text: "Changes to the laminar-flow ceiling", options: { bullet: true, breakLine: true } },
    { text: "Scent, or any non-essential item", options: { bullet: true } },
  ], { x: 6.05, y: 1.85, w: 3.25, h: 2.4, fontSize: 13, color: C.text1, valign: "top", paraSpaceAfter: 6 });
  T(s, "Keep strong warm colours away from the surgical light field", { x: 6.05, y: 4.25, w: 3.25, h: 0.55, fontSize: 11, italic: true, color: C.text2 });

  // ---------- 14. Section: uniforms ----------
  section("Uniforms", "Uniform system", "Role colours plus printed names, on everything reusable");

  // ---------- 15. Print rule ----------
  s = content("Uniforms", "One rule decides what gets printed");
  card(s, 0.5, 1.2, 4.35, 3.7, C.accent6);
  await iconDot(s, 0.75, 1.4, 0.6, "FaTrashAlt", C.accent2);
  T(s, "Single-use: no print", { x: 1.5, y: 1.45, w: 3.2, h: 0.5, fontSize: 18, bold: true, color: C.accent2, fontFace: "Cambria" });
  T(s, [
    { text: "Surgical masks", options: { bullet: true, breakLine: true } },
    { text: "Disposable caps / bouffants", options: { bullet: true, breakLine: true } },
    { text: "Disposable surgical gowns", options: { bullet: true, breakLine: true } },
    { text: "Gloves, shoe covers, aprons", options: { bullet: true, breakLine: true } },
    { text: "Disposable patient gowns", options: { bullet: true } },
  ], { x: 0.75, y: 2.15, w: 3.9, h: 1.9, fontSize: 14, color: C.text1, valign: "top", paraSpaceAfter: 5 });
  T(s, "Brand via supplier colour choice only, where options exist", { x: 0.75, y: 4.15, w: 3.9, h: 0.55, fontSize: 12, italic: true, color: C.text1 });
  card(s, 5.15, 1.2, 4.35, 3.7);
  await iconDot(s, 5.4, 1.4, 0.6, "FaRecycle");
  T(s, "Reusable: printable", { x: 6.15, y: 1.45, w: 3.2, h: 0.5, fontSize: 18, bold: true, color: C.accent1, fontFace: "Cambria" });
  T(s, [
    { text: "Scrub suits (role colour + logo)", options: { bullet: true, breakLine: true } },
    { text: "Fabric caps (name + role)", options: { bullet: true, breakLine: true } },
    { text: "Warm-up jackets", options: { bullet: true, breakLine: true } },
    { text: "Reusable patient gown + cap", options: { bullet: true, breakLine: true } },
    { text: "Blankets, clogs, white coats", options: { bullet: true } },
  ], { x: 5.4, y: 2.15, w: 3.9, h: 1.9, fontSize: 14, color: C.text1, valign: "top", paraSpaceAfter: 5 });
  T(s, "Every print must pass 50+ washes at 71 °C / 25 min", { x: 5.4, y: 4.15, w: 3.9, h: 0.55, fontSize: 12, italic: true, color: C.text2 });
  s.addNotes("Decision implication: choose reusable fabric caps (not disposable) so names can be printed, and reusable patient gowns so the patient wears the brand.");

  // ---------- 16. Role colours ----------
  s = content("Uniforms", "Role colours: recognisable at a glance");
  const roles = [
    ["1F3A5F", "Surgeons", "Name + \"Surgeon\""], ["6B3E75", "Anaesthesia", "Name + \"Anaesthesia\""],
    ["1F7A74", "Scrub nurses", "Name + \"Scrub Nurse\""], ["5FAEA5", "Circulating nurses", "Name + \"Nurse\""],
    ["8FBCD6", "Pre-op / PACU nurses", "Name + \"Nurse\""], ["8A9199", "Techs / assistants", "Name + role"],
    ["C9C1B1", "CSSD / housekeeping", "Plain cap"], ["FFFFFF", "Trainees", "Role colour + white trim"],
  ];
  roles.forEach((r, i) => {
    const x = 0.5 + (i % 4) * 2.3, y = 1.2 + Math.floor(i / 4) * 1.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.1, h: 0.85, rectRadius: 0.1, fill: { color: r[0] }, line: { color: "C9C1B1", width: 1 }, objectName: "sw" + i });
    T(s, r[1], { x, y: y + 0.95, w: 2.1, h: 0.32, fontSize: 13, bold: true, color: C.text2 });
    T(s, "Cap: " + r[2], { x, y: y + 1.27, w: 2.1, h: 0.3, fontSize: 11, color: C.text1 });
  });
  T(s, "Swatches are placeholders. Final shades come from the Noya palette. Concierge and coordinators wear a fully branded uniform outside the OR.", { x: 0.5, y: 4.7, w: 9, h: 0.4, fontSize: 11, italic: true, color: C.accent5 });

  // ---------- 17. Kit ----------
  s = content("Uniforms", "The reusable kit");
  const kit = [
    ["FaTshirt", "Scrub suit", "Low-lint poly-cotton, V-neck, 2 pockets. Small logo on the left chest + role colour"],
    ["FaIdBadge", "Name cap", "Fabric cap in role colour, name + role printed in the brand typeface. Covers hair fully"],
    ["FaShieldAlt", "Warm-up jacket", "MoH-required in the OR. Long sleeves, snap front, logo + role colour. Facility-laundered"],
    ["FaBed", "Patient gown + cap", "Reusable wrap-around gown in a Noya pattern, matching cap"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.2, 2.1, 2.85);
    await iconDot(s, x + 0.2, 1.4, 0.65, kit[i][0]);
    T(s, kit[i][1], { x: x + 0.2, y: 2.2, w: 1.75, h: 0.45, fontSize: 15, bold: true, color: C.text2 });
    T(s, kit[i][2], { x: x + 0.2, y: 2.7, w: 1.75, h: 1.3, fontSize: 12, color: C.text1, valign: "top" });
  }

  card(s, 0.5, 4.25, 9.0, 0.7, C.accent6);
  T(s, "Every printed item passes 50+ washes at 71 °C / 25 min before bulk order. Single-use items stay unprinted.", { x: 0.75, y: 4.25, w: 8.5, h: 0.7, fontSize: 13, bold: true, color: C.text1, valign: "middle" });

  // ---------- 18. Rollout ----------
  section("Rollout", "Rollout & measures", "14 weeks from inputs to launch");
  s = content("Rollout", "14-week rollout");
  const plan = [["Wk 1", "Inputs", "Brand guide, drawings, equipment list"], ["Wk 2-4", "Design", "Palette, OR finishes, uniforms, scripts"],
                ["Wk 4-6", "Approvals", "IPC committee + MoH sign-off on materials"], ["Wk 5-8", "Samples", "50-wash test on prints and fabrics"],
                ["Wk 8-12", "Build", "Finishes, uniforms (+20% stock), lounge"], ["Wk 10-12", "Training", "Scripts, Noya Moment role-play"],
                ["Wk 12-14", "Soft launch", "Reduced list, daily debriefs"]];
  plan.forEach((p, i) => {
    const x = 0.5 + i * 1.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.3 + (i % 2) * 0.15, w: 1.2, h: 0.55, rectRadius: 0.08, fill: { color: i === 2 ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: "pl" + i });
    T(s, p[0], { x, y: 1.3 + (i % 2) * 0.15, w: 1.2, h: 0.55, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, p[1], { x, y: 2.15, w: 1.2, h: 0.35, fontSize: 13, bold: true, color: C.text2 });
    T(s, p[2], { x, y: 2.5, w: 1.2, h: 1.2, fontSize: 11, color: C.text1, valign: "top" });
  });
  card(s, 0.5, 3.6, 9.0, 1.0, C.accent6);
  T(s, "Critical gate: no uniform or finish is ordered before written approval of the materials list (week 6).", { x: 0.75, y: 3.6, w: 8.5, h: 1.0, fontSize: 14, bold: true, color: C.text1, valign: "middle" });

  // ---------- 19. KPIs ----------
  s = content("Rollout", "How we'll know it works");
  const kpis = [["<3%", "Same-day cancellations"], ["100%", "WHO checklist compliance"], ["<10 min", "Family update after surgery"], ["100%", "Patients called on D-1"]];
  kpis.forEach((kp, i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.3, 2.1, 2.0);
    T(s, kp[0], { x: x + 0.1, y: 1.45, w: 1.9, h: 1.0, fontSize: 36, bold: true, color: i % 2 ? C.accent2 : C.accent1, align: "center", valign: "middle", fontFace: "Cambria" });
    T(s, kp[1], { x: x + 0.15, y: 2.5, w: 1.8, h: 0.65, fontSize: 13, color: C.text1, align: "center", valign: "top" });
  });
  T(s, "Also tracked: pre-op anxiety score (1-5), first-case on-time start, D+30 NPS, and surgical site infection rate (the brand promise depends on it).", { x: 0.5, y: 3.65, w: 9, h: 0.8, fontSize: 14, color: C.text1 });

  // ---------- 20. Decisions ----------
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "Title Dark", sectionTitle: "Close" });
  s.addText("Decisions needed today", { placeholder: "title" });
  T(s, [
    { text: "Approve the three-zone branding approach and the Noya Moment script", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Switch to reusable caps and patient gowns so they can carry print", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Share the Noya brand guide to finalise colours", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "Name owners: Surgical Coordinator lead, IPC sign-off, procurement", options: { bullet: { type: "number" } } },
  ], { x: 0.6, y: 3.25, w: 8.8, h: 1.9, fontSize: 16, color: C.background1, valign: "top", paraSpaceAfter: 6 });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
