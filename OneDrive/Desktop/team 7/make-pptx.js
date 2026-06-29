const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();

pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 inches

// ── Brand colours ──────────────────────────────────────────
const BLACK  = '0A0A0A';
const DARK   = '1A1200';
const GOLD   = 'D4AF37';
const GOLD2  = 'F0C040';
const WHITE  = 'FFFFFF';
const GREY   = 'AAAAAA';
const GREEN  = '00C853';
const RED    = 'FF5252';
const PINK   = 'FF69B4';

// ── Reusable helpers ───────────────────────────────────────
function darkBg(slide) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: '100%', h: '100%',
    fill: { color: BLACK }
  });
  // subtle grid lines
  for (let i = 0; i < 14; i++) {
    slide.addShape(pptx.shapes.LINE, {
      x: i * 0.95, y: 0, w: 0, h: 7.5,
      line: { color: '1E1600', width: 0.5 }
    });
  }
  for (let j = 0; j < 8; j++) {
    slide.addShape(pptx.shapes.LINE, {
      x: 0, y: j * 0.95, w: 13.33, h: 0,
      line: { color: '1E1600', width: 0.5 }
    });
  }
}

function goldBar(slide, y) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: y, w: '100%', h: 0.04,
    fill: { color: GOLD }
  });
}

function badge(slide, num, x, y) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: 0.55, h: 0.55,
    fill: { type: 'none' },
    line: { color: GOLD, width: 1.5 }
  });
  slide.addText(num, {
    x, y, w: 0.55, h: 0.55,
    fontSize: 14, bold: true, color: GOLD,
    align: 'center', valign: 'middle',
    fontFace: 'Arial'
  });
}

function sectionTitle(slide, text) {
  darkBg(slide);
  goldBar(slide, 0.15);
  goldBar(slide, 7.3);
  slide.addText(text, {
    x: 0.5, y: 2.5, w: 12.33, h: 2.5,
    fontSize: 52, bold: true, italic: true,
    color: GOLD, align: 'center', valign: 'middle',
    fontFace: 'Arial'
  });
}

function upNextLabel(slide) {
  slide.addText('Up Next:', {
    x: 0.5, y: 1.5, w: 12.33, h: 0.6,
    fontSize: 20, color: GREY, align: 'center',
    fontFace: 'Arial', charSpacing: 4
  });
}

function infoBox(slide, x, y, w, h, title, body, icon) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: '1A1200' },
    line: { color: GOLD, width: 0.8 },
    rectRadius: 0.05
  });
  if (icon) {
    slide.addText(icon, {
      x, y: y + 0.08, w, h: 0.5,
      fontSize: 18, align: 'center'
    });
  }
  slide.addText(title, {
    x: x + 0.15, y: y + (icon ? 0.55 : 0.12), w: w - 0.3, h: 0.35,
    fontSize: 11, bold: true, color: GOLD,
    fontFace: 'Arial', charSpacing: 1
  });
  slide.addText(body, {
    x: x + 0.15, y: y + (icon ? 0.88 : 0.45), w: w - 0.3, h: h - (icon ? 1.0 : 0.55),
    fontSize: 10, color: 'CCCCCC',
    fontFace: 'Arial', wrap: true
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 1 — TEAM 7: THE INVINCIBLES
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  goldBar(s, 7.28);
  badge(s, '01', 0.4, 0.3);
  s.addText('TEAM 7', {
    x: 0.5, y: 1.1, w: 12.33, h: 0.7,
    fontSize: 22, bold: true, color: GOLD,
    align: 'center', charSpacing: 8, fontFace: 'Arial'
  });
  s.addText('THE INVINCIBLES', {
    x: 0.5, y: 1.7, w: 12.33, h: 1.8,
    fontSize: 72, bold: true, italic: true, color: GOLD2,
    align: 'center', fontFace: 'Arial'
  });
  s.addText('DIFFERENT STRENGTHS  ·  ONE MISSION  ·  UNSTOPPABLE TOGETHER', {
    x: 0.5, y: 3.6, w: 12.33, h: 0.5,
    fontSize: 13, color: GREY,
    align: 'center', charSpacing: 3, fontFace: 'Arial'
  });
  // icon strip
  const icons = ['🛡️','👁️','📊','</>','🎨','⚙️'];
  icons.forEach((ic, i) => {
    s.addText(ic, {
      x: 1.5 + i * 1.7, y: 4.5, w: 1.4, h: 0.8,
      fontSize: 24, align: 'center'
    });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 2 — PRICILLA
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  badge(s, '02', 0.4, 0.3);
  s.addText('PRICILLA', {
    x: 0.5, y: 1.1, w: 6, h: 1.2,
    fontSize: 54, bold: true, color: GOLD2, fontFace: 'Arial'
  });
  s.addText('STRATEGIC LEAD', {
    x: 0.5, y: 2.2, w: 6, h: 0.5,
    fontSize: 18, bold: true, color: WHITE,
    charSpacing: 4, fontFace: 'Arial'
  });
  s.addText('🎯', { x: 0.5, y: 3.0, w: 0.6, h: 0.6, fontSize: 22 });
  s.addText('SUPERPOWER:', {
    x: 1.1, y: 3.0, w: 5, h: 0.3,
    fontSize: 11, bold: true, color: GOLD, charSpacing: 2, fontFace: 'Arial'
  });
  s.addText('THE MASTER PLAN', {
    x: 1.1, y: 3.3, w: 5, h: 0.4,
    fontSize: 16, bold: true, color: WHITE, fontFace: 'Arial'
  });
  s.addText('Sees the big picture and maps\nthe path to victory.', {
    x: 0.5, y: 3.9, w: 5.5, h: 0.9,
    fontSize: 13, color: GREY, fontFace: 'Arial'
  });
  infoBox(s, 7.2, 1.0, 5.6, 1.5, 'ROLE', 'Leads strategy, sets direction, and ensures all team goals align with the mission.', null);
  infoBox(s, 7.2, 2.7, 5.6, 1.5, 'KEY STRENGTH', 'Big-picture thinking combined with decisive action planning.', null);
  infoBox(s, 7.2, 4.4, 5.6, 1.5, 'IMPACT', 'Ensures the team never loses sight of what matters most — delivering results.', null);
}

// ══════════════════════════════════════════════════════════
// SLIDE 3 — NKOSI
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  badge(s, '03', 0.4, 0.3);
  s.addText('NKOSI', {
    x: 0.5, y: 1.1, w: 6, h: 1.2,
    fontSize: 54, bold: true, color: GOLD2, fontFace: 'Arial'
  });
  s.addText('CREATIVE LEAD', {
    x: 0.5, y: 2.2, w: 6, h: 0.5,
    fontSize: 18, bold: true, color: WHITE,
    charSpacing: 4, fontFace: 'Arial'
  });
  s.addText('👁️', { x: 0.5, y: 3.0, w: 0.6, h: 0.6, fontSize: 22 });
  s.addText('SUPERPOWER:', {
    x: 1.1, y: 3.0, w: 5, h: 0.3,
    fontSize: 11, bold: true, color: GOLD, charSpacing: 2, fontFace: 'Arial'
  });
  s.addText('THE VISION', {
    x: 1.1, y: 3.3, w: 5, h: 0.4,
    fontSize: 16, bold: true, color: WHITE, fontFace: 'Arial'
  });
  s.addText('Turns ideas into inspiring concepts\nthat spark change.', {
    x: 0.5, y: 3.9, w: 5.5, h: 0.9,
    fontSize: 13, color: GREY, fontFace: 'Arial'
  });
  infoBox(s, 7.2, 1.0, 5.6, 1.5, 'ROLE', 'Drives creative direction, visual identity and conceptual innovation for the team.', null);
  infoBox(s, 7.2, 2.7, 5.6, 1.5, 'KEY STRENGTH', 'Transforming abstract ideas into tangible, impactful design concepts.', null);
  infoBox(s, 7.2, 4.4, 5.6, 1.5, 'IMPACT', 'Sets the creative tone that makes the product unforgettable and user-loved.', null);
}

// ══════════════════════════════════════════════════════════
// SLIDE 4 — PHUNYEZWA
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  badge(s, '04', 0.4, 0.3);
  s.addText('PHUNYEZWA', {
    x: 0.5, y: 1.1, w: 6.5, h: 1.0,
    fontSize: 44, bold: true, color: GOLD2, fontFace: 'Arial'
  });
  s.addText('FEATURES & LAYOUT SUPERHERO', {
    x: 0.5, y: 2.0, w: 6.5, h: 0.5,
    fontSize: 14, bold: true, color: WHITE, charSpacing: 3, fontFace: 'Arial'
  });
  s.addText('📋', { x: 0.5, y: 2.8, w: 0.6, h: 0.6, fontSize: 22 });
  s.addText('SUPERPOWER:', {
    x: 1.1, y: 2.8, w: 5, h: 0.3,
    fontSize: 11, bold: true, color: GOLD, charSpacing: 2, fontFace: 'Arial'
  });
  s.addText('DESIGN THAT EMPOWERS', {
    x: 1.1, y: 3.1, w: 5.5, h: 0.5,
    fontSize: 15, bold: true, color: WHITE, fontFace: 'Arial'
  });
  s.addText('Creates experiences that are\nbeautiful, intuitive and bold.', {
    x: 0.5, y: 3.75, w: 5.5, h: 0.9,
    fontSize: 13, color: GREY, fontFace: 'Arial'
  });
  infoBox(s, 7.2, 1.0, 5.6, 1.5, 'ROLE', 'Owns features definition, layout architecture and user experience flow design.', null);
  infoBox(s, 7.2, 2.7, 5.6, 1.5, 'KEY STRENGTH', 'Bridging user needs and design solutions with clarity and precision.', null);
  infoBox(s, 7.2, 4.4, 5.6, 1.5, 'IMPACT', 'Every screen users see is shaped by Phunyezwa\'s bold and empowering design thinking.', null);
}

// ══════════════════════════════════════════════════════════
// SLIDE 5 — WANDILE
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  badge(s, '05', 0.4, 0.3);
  s.addText('WANDILE', {
    x: 0.5, y: 1.1, w: 6, h: 1.2,
    fontSize: 54, bold: true, color: GOLD2, fontFace: 'Arial'
  });
  s.addText('APP BUILDER', {
    x: 0.5, y: 2.2, w: 6, h: 0.5,
    fontSize: 18, bold: true, color: WHITE, charSpacing: 4, fontFace: 'Arial'
  });
  s.addText('</>', { x: 0.5, y: 3.0, w: 0.8, h: 0.6, fontSize: 20, color: GOLD, bold: true, fontFace: 'Courier New' });
  s.addText('SUPERPOWER:', {
    x: 1.3, y: 3.0, w: 5, h: 0.3,
    fontSize: 11, bold: true, color: GOLD, charSpacing: 2, fontFace: 'Arial'
  });
  s.addText('BUILDS THE DREAM', {
    x: 1.3, y: 3.3, w: 5, h: 0.4,
    fontSize: 16, bold: true, color: WHITE, fontFace: 'Arial'
  });
  s.addText('Turns concepts into powerful,\nworking solutions.', {
    x: 0.5, y: 3.9, w: 5.5, h: 0.9,
    fontSize: 13, color: GREY, fontFace: 'Arial'
  });
  infoBox(s, 7.2, 1.0, 5.6, 1.5, 'ROLE', 'Leads development, builds the application and brings all designs to life technically.', null);
  infoBox(s, 7.2, 2.7, 5.6, 1.5, 'KEY STRENGTH', 'Turning complex requirements into clean, functional, scalable code.', null);
  infoBox(s, 7.2, 4.4, 5.6, 1.5, 'IMPACT', 'Without Wandile, there is no product. He makes the dream real.', null);
}

// ══════════════════════════════════════════════════════════
// SLIDE 6 — LLOYD
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  badge(s, '06', 0.4, 0.3);
  s.addText('LLOYD', {
    x: 0.5, y: 1.1, w: 6, h: 1.2,
    fontSize: 54, bold: true, color: GOLD2, fontFace: 'Arial'
  });
  s.addText('FIGMA GURU', {
    x: 0.5, y: 2.2, w: 6, h: 0.5,
    fontSize: 18, bold: true, color: WHITE, charSpacing: 4, fontFace: 'Arial'
  });
  s.addText('🎨', { x: 0.5, y: 3.0, w: 0.6, h: 0.6, fontSize: 22 });
  s.addText('SUPERPOWER:', {
    x: 1.1, y: 3.0, w: 5, h: 0.3,
    fontSize: 11, bold: true, color: GOLD, charSpacing: 2, fontFace: 'Arial'
  });
  s.addText('HARMONY & BALANCE', {
    x: 1.1, y: 3.3, w: 5, h: 0.4,
    fontSize: 16, bold: true, color: WHITE, fontFace: 'Arial'
  });
  s.addText('Brings clarity to chaos and\ndesigns with precision and flow.', {
    x: 0.5, y: 3.9, w: 5.5, h: 0.9,
    fontSize: 13, color: GREY, fontFace: 'Arial'
  });
  infoBox(s, 7.2, 1.0, 5.6, 1.5, 'ROLE', 'Masters Figma prototyping, UI components and design system consistency.', null);
  infoBox(s, 7.2, 2.7, 5.6, 1.5, 'KEY STRENGTH', 'Creating pixel-perfect, harmonious interfaces that communicate clearly.', null);
  infoBox(s, 7.2, 4.4, 5.6, 1.5, 'IMPACT', 'The product looks and feels professional because of Lloyd\'s eye for detail.', null);
}

// ══════════════════════════════════════════════════════════
// SLIDE 7 — LWANDO
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  badge(s, '07', 0.4, 0.3);
  s.addText('LWANDO', {
    x: 0.5, y: 1.1, w: 6, h: 1.2,
    fontSize: 54, bold: true, color: GOLD2, fontFace: 'Arial'
  });
  s.addText('MISSION COMMANDER', {
    x: 0.5, y: 2.2, w: 6, h: 0.5,
    fontSize: 18, bold: true, color: WHITE, charSpacing: 4, fontFace: 'Arial'
  });
  s.addText('⭐', { x: 0.5, y: 3.0, w: 0.6, h: 0.6, fontSize: 22 });
  s.addText('SUPERPOWER:', {
    x: 1.1, y: 3.0, w: 5, h: 0.3,
    fontSize: 11, bold: true, color: GOLD, charSpacing: 2, fontFace: 'Arial'
  });
  s.addText('KEEPS US ALIGNED', {
    x: 1.1, y: 3.3, w: 5, h: 0.4,
    fontSize: 16, bold: true, color: WHITE, fontFace: 'Arial'
  });
  s.addText('Guides the team, manages the mission\nand shields us from distractions.', {
    x: 0.5, y: 3.9, w: 5.5, h: 0.9,
    fontSize: 13, color: GREY, fontFace: 'Arial'
  });
  infoBox(s, 7.2, 1.0, 5.6, 1.5, 'ROLE', 'Project management, sprint coordination and team alignment throughout all phases.', null);
  infoBox(s, 7.2, 2.7, 5.6, 1.5, 'KEY STRENGTH', 'Keeping focus, managing timelines and ensuring delivery without losing team energy.', null);
  infoBox(s, 7.2, 4.4, 5.6, 1.5, 'IMPACT', 'The team stays on track and delivers because Lwando keeps everyone mission-focused.', null);
}

// ══════════════════════════════════════════════════════════
// SLIDE 8 — PRESENTATION CHARACTERS
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  s.addText('PRESENTATION CHARACTERS', {
    x: 0.5, y: 0.4, w: 12.33, h: 0.7,
    fontSize: 26, bold: true, color: GOLD,
    align: 'center', charSpacing: 4, fontFace: 'Arial'
  });
  const members = [
    { name:'Pricilla', role:'Strategic Lead', icon:'🎯' },
    { name:'Nkosi', role:'Creative Lead', icon:'👁️' },
    { name:'Phunyezwa', role:'Features Lead', icon:'📋' },
    { name:'Wandile', role:'App Builder', icon:'</>' },
    { name:'Loyd', role:'Figma Lead', icon:'🎨' },
    { name:'Lwando', role:'Project Mom', icon:'⭐' },
  ];
  members.forEach((m, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 4.3;
    const y = 1.4 + row * 2.7;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w: 4.0, h: 2.4, fill: { color: '1A1200' }, line: { color: GOLD, width: 1 } });
    s.addText(m.icon, { x, y: y + 0.1, w: 4.0, h: 0.6, fontSize: 22, align: 'center' });
    s.addText(m.name, { x, y: y + 0.7, w: 4.0, h: 0.5, fontSize: 16, bold: true, color: GOLD2, align: 'center', fontFace: 'Arial' });
    s.addText(m.role, { x, y: y + 1.2, w: 4.0, h: 0.4, fontSize: 11, color: GREY, align: 'center', fontFace: 'Arial', charSpacing: 2 });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 9 — UP NEXT: ELEVATOR PITCH
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  sectionTitle(s, 'Our Elevator Pitch');
  upNextLabel(s);
  s.addText('🛗', { x: 5.5, y: 4.8, w: 2.33, h: 1.2, fontSize: 48, align: 'center' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 10 — SURVIVING TO THRIVING
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  s.addText('[ Surviving to Thriving ]', {
    x: 0.5, y: 0.3, w: 12.33, h: 0.7,
    fontSize: 28, bold: true, color: GOLD,
    align: 'center', fontFace: 'Arial'
  });
  const stories = [
    { icon:'😰', title:'Behind the Scenes Chaos', text:'Nomsa frantically counting ingredients, trying to serve customers while managing scattered orders.' },
    { icon:'😤', title:'The Breaking Point', text:'Long queue at the food truck. "Sorry — Out of Chicken & Buns". Disappointed customers walk away.' },
    { icon:'😔', title:'Nighttime Struggle', text:'Nomsa at her kitchen table, surrounded by notebooks and receipts. Clearly overwhelmed.' },
    { icon:'💡', title:'The Discovery', text:'Nomsa on her couch at night. A message from her friend: "Try this inventory app!"' },
    { icon:'😊', title:'First Light Relief', text:'Nomsa smiling at her phone, using a clean inventory management app with real-time stock.' },
    { icon:'🚀', title:'Smooth Operation', text:'Back at the food truck, Nomsa confidently serving customers. Everything organised — no stress.' },
  ];
  stories.forEach((st, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.3 + col * 4.35;
    const y = 1.2 + row * 2.85;
    s.addShape(pptx.shapes.RECTANGLE, { x, y, w: 4.1, h: 2.6, fill: { color: '1A1200' }, line: { color: GOLD, width: 0.8 } });
    s.addText(st.icon, { x, y: y + 0.1, w: 4.1, h: 0.55, fontSize: 20, align: 'center' });
    s.addText(st.title, { x: x + 0.1, y: y + 0.65, w: 3.9, h: 0.4, fontSize: 10, bold: true, color: GOLD2, fontFace: 'Arial' });
    s.addText(st.text, { x: x + 0.1, y: y + 1.05, w: 3.9, h: 1.4, fontSize: 9, color: 'BBBBBB', fontFace: 'Arial', wrap: true });
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 11 — UP NEXT: DESIGN THINKING
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  sectionTitle(s, 'The Design\nThinking Process');
  upNextLabel(s);
  s.addText('💡', { x: 5.5, y: 5.0, w: 2.33, h: 1.0, fontSize: 42, align: 'center' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 12 — NOMSA'S USER PROFILE
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  s.addText("NOMSA'S USER PROFILE", {
    x: 0.5, y: 0.3, w: 12.33, h: 0.65,
    fontSize: 26, bold: true, color: GOLD,
    align: 'center', charSpacing: 4, fontFace: 'Arial'
  });
  infoBox(s, 0.4, 1.2, 6.0, 1.55, 'WHO SHE IS', 'Nomsa, 32, food truck owner in Johannesburg. Passionate and driven but overwhelmed by manual processes.', '📍');
  infoBox(s, 0.4, 2.9, 6.0, 1.55, 'PAIN POINTS', 'Runs out of stock unexpectedly. Loses track of sales. Spends hours doing admin with no real-time insights.', '😤');
  infoBox(s, 6.8, 1.2, 6.0, 1.55, 'HER GOAL', 'To manage her inventory, payments and business analytics from one simple, affordable mobile app.', '🎯');
  infoBox(s, 6.8, 2.9, 6.0, 1.55, 'HER WORDS', '"I just need to know what I have, what I sold and what I made — without spending hours on spreadsheets."', '💬');
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 4.65, w: 12.4, h: 1.8,
    fill: { color: '1A1200' }, line: { color: GOLD, width: 0.8 }
  });
  s.addText('KEY INSIGHT', { x: 0.6, y: 4.75, w: 12, h: 0.35, fontSize: 11, bold: true, color: GOLD, fontFace: 'Arial', charSpacing: 2 });
  s.addText('SME owners need an all-in-one mobile solution that unifies inventory tracking, payment processing and sales analytics with zero technical complexity.', {
    x: 0.6, y: 5.1, w: 12, h: 1.2, fontSize: 12, color: WHITE, fontFace: 'Arial', wrap: true
  });
}

// ══════════════════════════════════════════════════════════
// SLIDE 13 — THE IDENTIFIED PROBLEM
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  s.addText('🔍', { x: 5.5, y: 0.6, w: 2.33, h: 1.0, fontSize: 42, align: 'center' });
  s.addText('THE IDENTIFIED PROBLEM', {
    x: 0.5, y: 1.5, w: 12.33, h: 1.2,
    fontSize: 40, bold: true, italic: true, color: GOLD,
    align: 'center', fontFace: 'Arial'
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 1.0, y: 3.0, w: 11.33, h: 2.4,
    fill: { color: '1A1200' }, line: { color: GOLD, width: 1 }
  });
  s.addText('Small and medium enterprises (SMEs) like food truck operators lack an integrated, affordable digital tool to manage inventory, payments, and sales analytics in real time — forcing them to rely on fragmented, manual processes that cost time and revenue.', {
    x: 1.3, y: 3.2, w: 10.8, h: 2.0,
    fontSize: 15, color: WHITE, fontFace: 'Arial',
    wrap: true, align: 'center', valign: 'middle'
  });
  goldBar(s, 5.6);
}

// ══════════════════════════════════════════════════════════
// SLIDE 14 — HOW MIGHT WE
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  s.addText('HOW MIGHT WE', {
    x: 0.5, y: 0.35, w: 12.33, h: 0.9,
    fontSize: 42, bold: true, color: GOLD,
    align: 'center', fontFace: 'Arial'
  });
  infoBox(s, 0.4, 1.5, 3.9, 3.5, 'INVENTORY', 'How might we help Nomsa track stock levels in real time without manual counting or unexpected stockouts?', '📦');
  infoBox(s, 4.7, 1.5, 3.9, 3.5, 'PAYMENTS', 'How might we simplify payment collection so Nomsa never loses a sale or misses a transaction?', '💳');
  infoBox(s, 9.0, 1.5, 3.9, 3.5, 'ANALYTICS', 'How might we give Nomsa clear, daily insights into her business performance with zero accounting knowledge?', '📊');
  goldBar(s, 5.2);
}

// ══════════════════════════════════════════════════════════
// SLIDE 15 — BRAINSTORMING & IDEATE
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  s.addText('💡', { x: 5.8, y: 0.3, w: 1.8, h: 0.9, fontSize: 36, align: 'center' });
  s.addText('Brainstorming & Ideate', {
    x: 0.5, y: 1.0, w: 12.33, h: 0.8,
    fontSize: 36, bold: true, color: GOLD,
    align: 'center', fontFace: 'Arial'
  });
  infoBox(s, 0.4, 2.1, 3.9, 2.8, 'RAPID IDEATION', '60-minute sprint generating 30+ feature ideas per round across inventory, payments and analytics.', '🔄');
  infoBox(s, 4.7, 2.1, 3.9, 2.8, 'DOT VOTING', 'Team voted on top ideas based on impact vs. feasibility matrix. Top 5 ideas advanced to prototype.', '🗳️');
  infoBox(s, 9.0, 2.1, 3.9, 2.8, 'TOP 3 FEATURES', '1. Unified dashboard  2. Smart low-stock alerts  3. One-tap payment integration', '🏆');
  goldBar(s, 5.1);
}

// ══════════════════════════════════════════════════════════
// SLIDE 16 — UP NEXT: FINAL OUTCOME
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  sectionTitle(s, 'The Final Outcome');
  upNextLabel(s);
  s.addText('🏆', { x: 5.5, y: 5.0, w: 2.33, h: 1.0, fontSize: 42, align: 'center' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 17 — PROTOTYPE
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  s.addText('PROTOTYPE', {
    x: 0.5, y: 0.3, w: 12.33, h: 0.85,
    fontSize: 44, bold: true, color: GOLD,
    align: 'center', fontFace: 'Arial'
  });
  infoBox(s, 0.4, 1.4, 5.9, 2.2, 'WIREFRAMES', 'Low-fidelity sketches mapping all core screens and user flows, built collaboratively in Figma during sprint 1.', '📐');
  infoBox(s, 7.0, 1.4, 5.9, 2.2, 'HIGH-FIDELITY UI', 'Full colour, pixel-perfect screens designed with the SME user in mind — clean, bold and accessible.', '🎨');
  infoBox(s, 0.4, 3.8, 5.9, 2.2, 'CLICKABLE PROTOTYPE', 'Interactive Figma prototype demonstrating the complete user journey from onboarding to sales report.', '🔗');
  infoBox(s, 7.0, 3.8, 5.9, 2.2, 'USER TESTING', 'Tested with 3 SME users. Collected feedback, iterated on 12 UI elements. Final version approved.', '✅');
  goldBar(s, 6.2);
}

// ══════════════════════════════════════════════════════════
// SLIDE 18 — UP NEXT: CONCLUSION
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  sectionTitle(s, 'The Conclusion');
  upNextLabel(s);
  s.addText('🎯', { x: 5.5, y: 5.0, w: 2.33, h: 1.0, fontSize: 42, align: 'center' });
}

// ══════════════════════════════════════════════════════════
// SLIDE 19 — DESIGN JOURNEY: REFLECTION & GROWTH
// ══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  darkBg(s);
  goldBar(s, 0.18);
  s.addText('Our Design Journey: Reflection and Growth', {
    x: 0.4, y: 0.3, w: 12.53, h: 0.65,
    fontSize: 22, bold: true, color: GOLD,
    align: 'center', fontFace: 'Arial'
  });
  // PROJECT HIGHS
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 1.1, w: 6.1, h: 2.6, fill: { color: '001A0A' }, line: { color: GREEN, width: 1 } });
  s.addText('⬆  PROJECT HIGHS', { x: 0.5, y: 1.2, w: 5.9, h: 0.4, fontSize: 13, bold: true, color: GREEN, fontFace: 'Arial', charSpacing: 2 });
  s.addText('1.  Successful Prototype Development\n     (Concept to Reality)\n2.  Integrated Final Product', {
    x: 0.6, y: 1.7, w: 5.7, h: 1.8, fontSize: 12, color: WHITE, fontFace: 'Arial'
  });
  // PROJECT LOWS
  s.addShape(pptx.shapes.RECTANGLE, { x: 6.9, y: 1.1, w: 6.0, h: 2.6, fill: { color: '1A0000' }, line: { color: RED, width: 1 } });
  s.addText('⬇  PROJECT LOWS', { x: 7.0, y: 1.2, w: 5.8, h: 0.4, fontSize: 13, bold: true, color: RED, fontFace: 'Arial', charSpacing: 2 });
  s.addText('1.  Tight Sprint Timeline Management\n2.  Time Constraint Challenges', {
    x: 7.1, y: 1.7, w: 5.6, h: 1.8, fontSize: 12, color: WHITE, fontFace: 'Arial'
  });
  // KEY LEARNINGS
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 3.9, w: 12.53, h: 2.7, fill: { color: '1A1200' }, line: { color: GOLD, width: 1 } });
  s.addText('🧭  KEY TEAM LEARNINGS', { x: 0.6, y: 4.0, w: 12.1, h: 0.4, fontSize: 13, bold: true, color: GOLD, fontFace: 'Arial', charSpacing: 2 });
  s.addText('1.  The Crucial Role of Effective Teamwork\n2.  Value of Rapid Ideation + Cohesive Collaboration\n3.  Key Emphasis: Constant User-Centric Focus (compass pointing to core user)', {
    x: 0.6, y: 4.5, w: 12.1, h: 1.9, fontSize: 12, color: WHITE, fontFace: 'Arial'
  });
}

// ══════════════════════════════════════════════════════════
// SAVE
// ══════════════════════════════════════════════════════════
pptx.writeFile({ fileName: 'Team7_Invincibles_Presentation.pptx' })
  .then(() => console.log('\n  ✓  Saved:  Team7_Invincibles_Presentation.pptx\n'))
  .catch(e => console.error('Error:', e));
