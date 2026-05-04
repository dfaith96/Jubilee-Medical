const fs = require("fs");
const path = require("path");
const PptxGenJS = require(path.join(process.env.APPDATA, "npm", "node_modules", "pptxgenjs"));

const root = __dirname;
const outPath = path.join(root, "Jubilee_Med_Pitch_Deck.pptx");
const heroPath = path.join(root, "assets", "images", "jubilee-hero.png");

if (!fs.existsSync(heroPath)) {
  throw new Error(`Missing hero image: ${heroPath}`);
}

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Jubilee Medical";
pptx.company = "Jubilee Medical";
pptx.subject = "Jubilee Medical digital health platform pitch deck";
pptx.title = "Jubilee Medical Pitch Deck";
pptx.lang = "en-US";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "en-US",
};
pptx.defineSlideMaster({
  title: "JUBILEE_MASTER",
  background: { color: "F6FAF8" },
  objects: [],
  slideNumber: { x: 12.35, y: 7.06, color: "7B8C87" },
});

const C = {
  ink: "132421",
  muted: "5C6F6A",
  paper: "FFFFFF",
  bg: "F6FAF8",
  line: "D7E6DF",
  teal: "0F766E",
  tealDark: "0B504B",
  green: "17633C",
  mint: "DFF4EC",
  coral: "F26F5B",
  gold: "F6BD43",
  blue: "3B82F6",
};

function addBrand(slide, x = 0.55, y = 0.34, dark = false) {
  const textColor = dark ? C.paper : C.ink;
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w: 0.46,
    h: 0.46,
    rectRadius: 0.07,
    fill: { color: dark ? C.paper : C.tealDark },
    line: { color: dark ? C.paper : C.tealDark },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 0.09,
    y: y + 0.08,
    w: 0.28,
    h: 0.28,
    fill: { color: dark ? C.tealDark : C.mint },
    line: { color: dark ? C.tealDark : C.mint, width: 1 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.2,
    y: y + 0.15,
    w: 0.07,
    h: 0.18,
    fill: { color: C.coral },
    line: { color: C.coral },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.145,
    y: y + 0.205,
    w: 0.18,
    h: 0.07,
    fill: { color: C.coral },
    line: { color: C.coral },
  });
  slide.addText("Jubilee Medical", {
    x: x + 0.58,
    y: y + 0.06,
    w: 1.95,
    h: 0.3,
    fontFace: "Aptos Display",
    fontSize: 18,
    bold: true,
    color: textColor,
    margin: 0,
  });
}

function addFooter(slide) {
  slide.addShape(pptx.ShapeType.line, {
    x: 0.55,
    y: 6.9,
    w: 12.2,
    h: 0,
    line: { color: C.line, width: 1 },
  });
  slide.addText("Jubilee Medical | Digital health made clear", {
    x: 0.55,
    y: 7.05,
    w: 4.5,
    h: 0.2,
    fontSize: 8,
    color: C.muted,
    margin: 0,
  });
}

function title(slide, eyebrow, heading, body) {
  slide.addText(eyebrow.toUpperCase(), {
    x: 0.72,
    y: 0.95,
    w: 3.8,
    h: 0.22,
    fontSize: 9,
    bold: true,
    color: C.coral,
    charSpace: 0,
    margin: 0,
  });
  slide.addText(heading, {
    x: 0.7,
    y: 1.28,
    w: 5.55,
    h: 1.15,
    fontFace: "Aptos Display",
    fontSize: 30,
    bold: true,
    color: C.ink,
    breakLine: false,
    fit: "shrink",
    margin: 0,
  });
  if (body) {
    slide.addText(body, {
      x: 0.72,
      y: 2.56,
      w: 5.35,
      h: 0.64,
      fontSize: 13,
      color: C.muted,
      fit: "shrink",
      margin: 0,
      breakLine: false,
    });
  }
}

function card(slide, x, y, w, h, heading, body, color = C.teal) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.06,
    fill: { color: C.paper },
    line: { color: C.line, width: 1 },
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.18,
    y: y + 0.2,
    w: 0.34,
    h: 0.34,
    rectRadius: 0.04,
    fill: { color },
    line: { color },
  });
  slide.addText(heading, {
    x: x + 0.18,
    y: y + 0.68,
    w: w - 0.36,
    h: 0.32,
    fontSize: 15,
    bold: true,
    color: C.ink,
    margin: 0,
  });
  slide.addText(body, {
    x: x + 0.18,
    y: y + 1.08,
    w: w - 0.36,
    h: h - 1.25,
    fontSize: 10.2,
    color: C.muted,
    fit: "shrink",
    margin: 0,
    breakLine: false,
  });
}

function pill(slide, x, y, text, color = C.mint, textColor = C.tealDark) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w: 1.85,
    h: 0.38,
    rectRadius: 0.05,
    fill: { color },
    line: { color },
  });
  slide.addText(text, {
    x: x + 0.12,
    y: y + 0.1,
    w: 1.6,
    h: 0.16,
    fontSize: 8.5,
    bold: true,
    color: textColor,
    align: "center",
    margin: 0,
    fit: "shrink",
  });
}

function addThreeSteps(slide, steps, startY = 3.5) {
  steps.forEach((step, index) => {
    const x = 0.75 + index * 4.15;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: startY,
      w: 3.55,
      h: 1.56,
      rectRadius: 0.06,
      fill: { color: C.paper },
      line: { color: C.line },
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.18,
      y: startY + 0.22,
      w: 0.48,
      h: 0.48,
      fill: { color: index === 1 ? C.coral : C.tealDark },
      line: { color: index === 1 ? C.coral : C.tealDark },
    });
    slide.addText(String(index + 1), {
      x: x + 0.31,
      y: startY + 0.34,
      w: 0.22,
      h: 0.18,
      fontSize: 10,
      bold: true,
      color: C.paper,
      align: "center",
      margin: 0,
    });
    slide.addText(step[0], {
      x: x + 0.82,
      y: startY + 0.22,
      w: 2.45,
      h: 0.25,
      fontSize: 14,
      bold: true,
      color: C.ink,
      margin: 0,
    });
    slide.addText(step[1], {
      x: x + 0.82,
      y: startY + 0.6,
      w: 2.38,
      h: 0.62,
      fontSize: 9.6,
      color: C.muted,
      fit: "shrink",
      margin: 0,
      breakLine: false,
    });
  });
}

// 1. Cover
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  slide.background = { color: C.tealDark };
  slide.addImage({ path: heroPath, x: 6.95, y: 0, w: 6.38, h: 7.5 });
  slide.addShape(pptx.ShapeType.rect, {
    x: 6.7,
    y: 0,
    w: 1.3,
    h: 7.5,
    fill: { color: C.tealDark, transparency: 4 },
    line: { color: C.tealDark, transparency: 100 },
  });
  addBrand(slide, 0.7, 0.55, true);
  slide.addText("Jubilee Medical", {
    x: 0.75,
    y: 1.72,
    w: 5.4,
    h: 0.75,
    fontFace: "Aptos Display",
    fontSize: 42,
    bold: true,
    color: C.paper,
    margin: 0,
  });
  slide.addText("Digital healthcare that connects patients, providers, pharmacy, labs, records, and payments.", {
    x: 0.78,
    y: 2.58,
    w: 5.45,
    h: 0.7,
    fontSize: 17,
    color: "DDF7EF",
    fit: "shrink",
    margin: 0,
    breakLine: false,
  });
  pill(slide, 0.78, 3.62, "Beta launch: Aug 2026", C.coral, C.paper);
  pill(slide, 2.8, 3.62, "Nigeria first", C.mint, C.tealDark);
  slide.addText("Pitch deck", {
    x: 0.78,
    y: 6.78,
    w: 2.4,
    h: 0.28,
    fontSize: 12,
    bold: true,
    color: "DDF7EF",
    margin: 0,
  });
}

// 2. Problem
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(
    slide,
    "Problem",
    "Healthcare journeys are still too fragmented.",
    "Patients move between phone calls, paper records, cash payments, pharmacies, and labs. Providers lose time because the same care journey is split across disconnected tools.",
  );
  card(slide, 0.75, 3.55, 2.8, 1.7, "Access delays", "Patients struggle to find trusted services, book visits, and receive follow-up instructions.", C.coral);
  card(slide, 3.85, 3.55, 2.8, 1.7, "Scattered records", "Medical history, lab results, prescriptions, and payment receipts rarely stay in one reliable profile.", C.teal);
  card(slide, 6.95, 3.55, 2.8, 1.7, "Provider admin load", "Clinics spend too much time on manual queue, billing, inventory, and reporting tasks.", C.gold);
  card(slide, 10.05, 3.55, 2.55, 1.7, "Poor visibility", "Owners need clearer data on revenue, patient demand, and care operations.", C.blue);
  addFooter(slide);
}

// 3. Solution
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(
    slide,
    "Solution",
    "A simple connected health platform.",
    "Jubilee Medical gives patients one front door for care and gives providers a practical workspace for appointments, EHR, billing, pharmacy, labs, and reports.",
  );
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 3.55,
    w: 11.75,
    h: 1.9,
    rectRadius: 0.08,
    fill: { color: C.paper },
    line: { color: C.line },
  });
  ["Patient app", "Provider workspace", "Pharmacy and labs", "Payments and reports"].forEach((label, i) => {
    const x = 1.1 + i * 2.85;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 4.05,
      w: 2.2,
      h: 0.74,
      rectRadius: 0.06,
      fill: { color: [C.mint, "FFF0EC", "F9F3DA", "EFF6FF"][i] },
      line: { color: "FFFFFF", transparency: 100 },
    });
    slide.addText(label, {
      x,
      y: 4.28,
      w: 2.2,
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: [C.tealDark, "A33D2F", C.green, "1D4ED8"][i],
      align: "center",
      margin: 0,
    });
    if (i < 3) {
      slide.addShape(pptx.ShapeType.chevron, {
        x: x + 2.35,
        y: 4.23,
        w: 0.28,
        h: 0.28,
        fill: { color: C.line },
        line: { color: C.line },
      });
    }
  });
  addFooter(slide);
}

// 4. Product modules
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(slide, "Product", "The everyday modules providers need.", "Each module is understandable on its own, but stronger when connected to the same patient record.");
  card(slide, 0.75, 3.0, 3.7, 1.45, "Appointments and queue", "Online booking, visit reminders, check-in, triage, and staff calendars.", C.teal);
  card(slide, 4.85, 3.0, 3.7, 1.45, "EHR and consultation", "Vitals, diagnosis, notes, documents, e-prescriptions, and follow-up tasks.", C.coral);
  card(slide, 8.95, 3.0, 3.7, 1.45, "Pharmacy and lab", "Orders, fulfilment status, lab requests, sample tracking, and linked results.", C.gold);
  card(slide, 0.75, 4.85, 3.7, 1.45, "Payments", "Consultation, medication, lab, subscription, and facility billing receipts.", C.green);
  card(slide, 4.85, 4.85, 3.7, 1.45, "Patient mobile app", "Bookings, records, prescriptions, results, reminders, and care messages.", C.blue);
  card(slide, 8.95, 4.85, 3.7, 1.45, "Analytics", "Revenue, visits, demand, staff activity, and operations reporting.", C.tealDark);
  addFooter(slide);
}

// 5. Patient journey
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(
    slide,
    "Patient journey",
    "From need to follow-up in fewer steps.",
    "The patient experience is designed around plain choices: find care, book, pay, receive results, and keep the record.",
  );
  addThreeSteps(slide, [
    ["Create profile", "Patients add basic details, allergies, emergency contacts, and insurance information."],
    ["Book or consult", "They choose doctor, pharmacy, lab, teleconsultation, or follow-up service."],
    ["Keep the record", "Prescriptions, lab results, receipts, and care notes stay in one account."],
  ]);
  addFooter(slide);
}

// 6. Provider workflow
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(
    slide,
    "Provider workflow",
    "One care journey, one operational view.",
    "Jubilee Medical supports the way a clinic actually works from reception to consultation, billing, fulfilment, and reporting.",
  );
  const steps = ["Booking", "Check-in", "Consult", "Orders", "Payment", "Follow-up"];
  steps.forEach((step, i) => {
    const x = 0.85 + i * 2.0;
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: 3.72,
      w: 1.08,
      h: 1.08,
      fill: { color: i % 2 ? C.coral : C.tealDark },
      line: { color: i % 2 ? C.coral : C.tealDark },
    });
    slide.addText(String(i + 1), {
      x: x + 0.39,
      y: 4.02,
      w: 0.3,
      h: 0.22,
      fontSize: 14,
      bold: true,
      color: C.paper,
      align: "center",
      margin: 0,
    });
    slide.addText(step, {
      x: x - 0.2,
      y: 5.04,
      w: 1.48,
      h: 0.28,
      fontSize: 12,
      bold: true,
      color: C.ink,
      align: "center",
      margin: 0,
    });
    if (i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.line, {
        x: x + 1.08,
        y: 4.26,
        w: 0.86,
        h: 0,
        line: { color: C.line, width: 2, beginArrowType: "none", endArrowType: "triangle" },
      });
    }
  });
  addFooter(slide);
}

// 7. Market
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(
    slide,
    "Opportunity",
    "A practical digital layer for local healthcare.",
    "Jubilee Medical starts with Nigeria and focuses on providers and patients who need affordable, easy-to-understand tools that fit daily care delivery.",
  );
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 3.35,
    w: 5.55,
    h: 2.25,
    rectRadius: 0.06,
    fill: { color: C.tealDark },
    line: { color: C.tealDark },
  });
  slide.addText("Primary users", {
    x: 1.12,
    y: 3.72,
    w: 2.2,
    h: 0.28,
    fontSize: 16,
    bold: true,
    color: C.paper,
    margin: 0,
  });
  slide.addText("Patients, clinics, hospitals, pharmacies, laboratories, employer health programs, and health partners.", {
    x: 1.12,
    y: 4.2,
    w: 4.8,
    h: 0.75,
    fontSize: 13,
    color: "DDF7EF",
    fit: "shrink",
    margin: 0,
    breakLine: false,
  });
  card(slide, 6.75, 3.35, 2.75, 2.25, "Demand drivers", "Convenience, continuity of records, faster billing, remote care, and stronger provider operations.", C.coral);
  card(slide, 9.85, 3.35, 2.75, 2.25, "Expansion path", "Start with pilots, grow provider network, then layer payments, insurance, and partner services.", C.gold);
  addFooter(slide);
}

// 8. Business model
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(slide, "Business model", "Multiple revenue lines from one platform.", "The product can begin with provider subscriptions and expand into transaction and partner revenue as usage grows.");
  card(slide, 0.75, 3.35, 2.8, 1.65, "Provider SaaS", "Monthly plans for clinics, hospitals, pharmacies, and labs based on users and modules.", C.teal);
  card(slide, 3.85, 3.35, 2.8, 1.65, "Transactions", "Small platform fee on paid bookings, pharmacy orders, lab orders, and teleconsultations.", C.coral);
  card(slide, 6.95, 3.35, 2.8, 1.65, "Setup and training", "Onboarding, migration support, facility setup, staff training, and workflow configuration.", C.gold);
  card(slide, 10.05, 3.35, 2.55, 1.65, "Partners", "Employer care programs, insurance workflows, referral networks, and integrations.", C.blue);
  addFooter(slide);
}

// 9. Go to market
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(slide, "Go to market", "Win locally with focused pilots.", "The fastest route is to onboard care teams that already feel the pain of scattered tools and patient follow-up gaps.");
  addThreeSteps(slide, [
    ["Pilot facilities", "Start with selected clinics, pharmacies, and labs in one city to validate workflows."],
    ["Patient acquisition", "Use provider referrals, community health campaigns, and clear app onboarding."],
    ["Partner growth", "Add employer groups, diagnostics partners, pharmacies, and payment partnerships."],
  ]);
  addFooter(slide);
}

// 10. Roadmap and ask
{
  const slide = pptx.addSlide("JUBILEE_MASTER");
  addBrand(slide);
  title(
    slide,
    "Roadmap and ask",
    "Build, pilot, measure, and scale.",
    "Jubilee Medical needs product completion, compliance review, pilot partners, and launch support to move from beta concept to live healthcare operations.",
  );
  card(slide, 0.75, 3.25, 3.7, 1.9, "Next 90 days", "Finalize beta, prepare provider onboarding, test appointment, EHR, pharmacy, lab, payment, and reporting flows.", C.teal);
  card(slide, 4.85, 3.25, 3.7, 1.9, "Pilot goals", "Launch with early providers, monitor appointment volume, retention, payment success, and care follow-up.", C.coral);
  card(slide, 8.95, 3.25, 3.7, 1.9, "What we need", "Pilot partners, technical build support, compliance guidance, payment partnerships, and launch capital.", C.gold);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.75,
    y: 5.72,
    w: 11.9,
    h: 0.62,
    rectRadius: 0.06,
    fill: { color: C.tealDark },
    line: { color: C.tealDark },
  });
  slide.addText("hello@jubileemedical.health  |  +234 800 000 0000  |  Lagos, Nigeria", {
    x: 1.05,
    y: 5.95,
    w: 11.2,
    h: 0.18,
    fontSize: 12,
    bold: true,
    color: C.paper,
    align: "center",
    margin: 0,
  });
  addFooter(slide);
}

pptx
  .writeFile({ fileName: outPath })
  .then(() => {
    console.log(outPath);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
