const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
 
// Icon helper
const { FaBrain, FaDatabase, FaCogs, FaChartBar, FaRocket, FaExclamationTriangle, FaLightbulb, FaSearch, FaCode, FaCheckCircle } = require("react-icons/fa");
const { MdSentimentSatisfied, MdSentimentDissatisfied } = require("react-icons/md");
 
async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}
 
const makeShadow = () => ({ type: "outer", blur: 8, offset: 2, angle: 135, color: "000000", opacity: 0.12 });
 
// Color palette: Deep Navy + Teal accent + White
const C = {
  navy: "0D1B2A",
  navyMid: "1B2A3B",
  teal: "00B4D8",
  tealDark: "0077B6",
  white: "FFFFFF",
  offWhite: "F0F4F8",
  lightGray: "CBD5E1",
  darkText: "1E293B",
  accent: "F59E0B",  // amber for highlights
  red: "EF4444",
  green: "10B981",
};
 
async function createPresentation() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Muhammad Haroon & Mohammad Sami";
  pres.title = "NLP Sentiment Analysis - Movie Reviews";
 
  // ─────────────────────────────────────────
  // SLIDE 1: TITLE SLIDE
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
 
    // Teal accent bar left
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.25, h: 5.625, fill: { color: C.teal } });
 
    // Course tag
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.45, w: 3.0, h: 0.38, fill: { color: C.tealDark }, shadow: makeShadow() });
    s.addText("NLP — PROJECT ASSIGNMENT 4", { x: 0.5, y: 0.45, w: 3.0, h: 0.38, fontSize: 9, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
 
    // Main title
    s.addText("Sentiment Analysis", { x: 0.5, y: 1.05, w: 9, h: 0.9, fontSize: 44, color: C.white, bold: true, fontFace: "Calibri", margin: 0 });
    s.addText("of Movie Reviews", { x: 0.5, y: 1.9, w: 9, h: 0.75, fontSize: 38, color: C.teal, bold: true, fontFace: "Calibri", margin: 0 });
 
    s.addText("with Streamlit GUI", { x: 0.5, y: 2.65, w: 9, h: 0.5, fontSize: 22, color: C.lightGray, fontFace: "Calibri", margin: 0 });
 
    // Divider
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.3, w: 9, h: 0.03, fill: { color: C.teal } });
 
    // Authors and info
    s.addText([
      { text: "Muhammad Haroon", options: { bold: true } },
      { text: "  (01-134231-052)     " },
      { text: "Mohammad Sami", options: { bold: true } },
      { text: "  (01-134231-063)" }
    ], { x: 0.5, y: 3.5, w: 9, h: 0.35, fontSize: 13, color: C.lightGray, fontFace: "Calibri" });
 
    s.addText("Bahria University, Islamabad  |  Dept. of Computer Science  |  Instructor: Ma'am Saira Hameed  |  May 2026",
      { x: 0.5, y: 3.95, w: 9, h: 0.3, fontSize: 10, color: "64748B", fontFace: "Calibri" });
 
    // Brain icon placeholder (decorative circles)
    s.addShape(pres.shapes.OVAL, { x: 7.2, y: 0.8, w: 2.4, h: 2.4, fill: { color: C.tealDark, transparency: 70 } });
    s.addShape(pres.shapes.OVAL, { x: 7.6, y: 1.2, w: 1.6, h: 1.6, fill: { color: C.teal, transparency: 50 } });
    s.addText("🎬", { x: 7.8, y: 1.55, w: 1.2, h: 1.0, fontSize: 36, align: "center" });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 2: PROBLEM STATEMENT
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
 
    // Header band
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
    s.addText("The Problem", { x: 0.4, y: 0, w: 9, h: 1.0, fontSize: 30, color: C.white, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.0, w: 10, h: 0.06, fill: { color: C.teal } });
 
    // Left column: problem
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.3, w: 4.4, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.3, w: 4.4, h: 0.5, fill: { color: C.navyMid } });
    s.addText("CHALLENGE", { x: 0.35, y: 1.3, w: 4.4, h: 0.5, fontSize: 13, color: C.teal, bold: true, align: "center", valign: "middle", margin: 0 });
 
    s.addText([
      { text: "Millions of online movie reviews exist — but reading them manually is:", options: { breakLine: true } },
    ], { x: 0.55, y: 1.9, w: 4.0, h: 0.5, fontSize: 12, color: C.darkText, fontFace: "Calibri" });
 
    const probs = [
      ["⏱", "Time-consuming at scale"],
      ["🎭", "Subjective & inconsistent"],
      ["🔍", "Fails to detect sarcasm & negation"],
      ["📉", "Keyword methods are inaccurate"],
    ];
    probs.forEach(([icon, txt], i) => {
      s.addShape(pres.shapes.RECTANGLE, { x: 0.55, y: 2.5 + i * 0.6, w: 4.0, h: 0.5, fill: { color: C.offWhite } });
      s.addText(icon + "  " + txt, { x: 0.55, y: 2.5 + i * 0.6, w: 4.0, h: 0.5, fontSize: 12, color: C.darkText, valign: "middle", margin: 8 });
    });
 
    // Right column: solution / stakeholders
    s.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.3, w: 4.4, h: 3.8, fill: { color: C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.3, w: 4.4, h: 0.5, fill: { color: C.tealDark } });
    s.addText("SOLUTION & IMPACT", { x: 5.25, y: 1.3, w: 4.4, h: 0.5, fontSize: 13, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
 
    s.addText("Automated ML-based sentiment classifier delivers:", { x: 5.45, y: 1.9, w: 4.0, h: 0.45, fontSize: 12, color: C.darkText });
    const sols = [
      ["🎬", "Studios: gauge audience reaction"],
      ["👤", "Viewers: make informed choices"],
      ["📊", "Platforms: automate recommendations"],
      ["🔬", "Researchers: study opinion trends"],
    ];
    sols.forEach(([icon, txt], i) => {
      s.addShape(pres.shapes.RECTANGLE, { x: 5.45, y: 2.5 + i * 0.6, w: 4.0, h: 0.5, fill: { color: "E0F2FE" } });
      s.addText(icon + "  " + txt, { x: 5.45, y: 2.5 + i * 0.6, w: 4.0, h: 0.5, fontSize: 12, color: C.tealDark, valign: "middle", margin: 8, bold: true });
    });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 3: PIPELINE / ARCHITECTURE
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
 
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navyMid } });
    s.addText("System Pipeline", { x: 0.4, y: 0, w: 9, h: 1.0, fontSize: 30, color: C.white, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.0, w: 10, h: 0.06, fill: { color: C.teal } });
 
    const steps = [
      { num: "1", label: "Data\nCollection", sub: "IMDB / Kaggle\n13K+ labeled reviews", color: "0077B6" },
      { num: "2", label: "Preprocessing", sub: "Remove punct, stopwords,\nstemming (NLTK)", color: "0096C7" },
      { num: "3", label: "Feature\nEngineering", sub: "Bag-of-Words\nTF-IDF, N-grams", color: "00B4D8" },
      { num: "4", label: "Naive Bayes\nTraining", sub: "MultinomialNB\n80/20 train-test split", color: "0096C7" },
      { num: "5", label: "Prediction &\nEvaluation", sub: "F1-Score, Confusion\nMatrix, Accuracy", color: "0077B6" },
      { num: "6", label: "Streamlit\nGUI", sub: "Interactive input\n& visualization", color: "005F9E" },
    ];
 
    const boxW = 1.35, boxH = 2.0, startX = 0.3, y = 1.35;
    steps.forEach((st, i) => {
      const x = startX + i * (boxW + 0.2);
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: boxW, h: boxH, fill: { color: st.color }, shadow: makeShadow() });
      // Number circle
      s.addShape(pres.shapes.OVAL, { x: x + 0.45, y: y + 0.12, w: 0.45, h: 0.45, fill: { color: C.accent } });
      s.addText(st.num, { x: x + 0.45, y: y + 0.12, w: 0.45, h: 0.45, fontSize: 14, color: C.navy, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(st.label, { x: x + 0.05, y: y + 0.65, w: boxW - 0.1, h: 0.65, fontSize: 11, color: C.white, bold: true, align: "center", fontFace: "Calibri", margin: 0 });
      s.addText(st.sub, { x: x + 0.05, y: y + 1.3, w: boxW - 0.1, h: 0.65, fontSize: 8.5, color: "BDE0FE", align: "center", fontFace: "Calibri", margin: 0 });
 
      // Arrow
      if (i < steps.length - 1) {
        s.addShape(pres.shapes.LINE, { x: x + boxW, y: y + boxH / 2, w: 0.2, h: 0, line: { color: C.teal, width: 2 } });
      }
    });
 
    // Math formula
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 3.55, w: 9.4, h: 1.6, fill: { color: "0D2137" } });
    s.addText("Naive Bayes Formula:", { x: 0.55, y: 3.65, w: 3, h: 0.35, fontSize: 11, color: C.teal, bold: true });
    s.addText("P(C | W) = P(C) × ∏ P(wᵢ | C) / P(W)", { x: 0.55, y: 4.0, w: 5.5, h: 0.45, fontSize: 14, color: C.white, fontFace: "Courier New" });
    s.addText("Where C = class (pos/neg), W = {w₁, w₂, ..., wₙ} = review words. Laplace smoothing handles unseen words.", { x: 0.55, y: 4.55, w: 9.0, h: 0.45, fontSize: 9.5, color: C.lightGray, fontFace: "Calibri" });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 4: DATA & PREPROCESSING
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
 
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
    s.addText("Data & Preprocessing", { x: 0.4, y: 0, w: 9, h: 1.0, fontSize: 30, color: C.white, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.0, w: 10, h: 0.06, fill: { color: C.teal } });
 
    // Dataset card
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 4.2, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 4.2, h: 0.5, fill: { color: C.navy } });
    s.addText("📦  DATASET", { x: 0.3, y: 1.2, w: 4.2, h: 0.5, fontSize: 13, color: C.teal, bold: true, align: "center", valign: "middle", margin: 0 });
 
    const dsRows = [
      ["Source", "IMDB Movie Reviews"],
      ["Size", "~13,000+ labeled reviews"],
      ["Classes", "Positive / Negative (binary)"],
      ["Split", "80% Train — 20% Test"],
      ["Format", "CSV — text + label columns"],
    ];
    dsRows.forEach(([k, v], i) => {
      const bg = i % 2 === 0 ? C.offWhite : C.white;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.8 + i * 0.58, w: 4.2, h: 0.55, fill: { color: bg } });
      s.addText(k, { x: 0.45, y: 1.82 + i * 0.58, w: 1.3, h: 0.48, fontSize: 11, color: C.tealDark, bold: true, valign: "middle", margin: 0 });
      s.addText(v, { x: 1.85, y: 1.82 + i * 0.58, w: 2.5, h: 0.48, fontSize: 11, color: C.darkText, valign: "middle", margin: 0 });
    });
 
    // Preprocessing steps
    s.addShape(pres.shapes.RECTANGLE, { x: 4.85, y: 1.2, w: 4.8, h: 4.1, fill: { color: C.white }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.85, y: 1.2, w: 4.8, h: 0.5, fill: { color: C.tealDark } });
    s.addText("⚙️  PREPROCESSING STEPS", { x: 4.85, y: 1.2, w: 4.8, h: 0.5, fontSize: 13, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
 
    const steps = [
      ["1.", "Text Cleaning", "Remove punct, numbers, lowercase"],
      ["2.", "Stop Word Removal", "Eliminate 'the', 'is', 'and'... (NLTK)"],
      ["3.", "Tokenization", "Split text into individual word tokens"],
      ["4.", "Stemming", '"running" → "run" (PorterStemmer)'],
      ["5.", "Feature Extraction", "BoW / TF-IDF / N-grams (Scikit-Learn)"],
    ];
    steps.forEach(([num, title, desc], i) => {
      s.addShape(pres.shapes.OVAL, { x: 5.05, y: 1.83 + i * 0.63, w: 0.32, h: 0.32, fill: { color: C.teal } });
      s.addText(num, { x: 5.05, y: 1.83 + i * 0.63, w: 0.32, h: 0.32, fontSize: 9, color: C.white, align: "center", valign: "middle", margin: 0 });
      s.addText(title, { x: 5.45, y: 1.82 + i * 0.63, w: 4.0, h: 0.28, fontSize: 11, color: C.darkText, bold: true, margin: 0 });
      s.addText(desc, { x: 5.45, y: 2.08 + i * 0.63, w: 4.0, h: 0.25, fontSize: 9.5, color: "64748B", margin: 0 });
    });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 5: MODEL & TRAINING
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
 
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navyMid } });
    s.addText("Model: Naive Bayes Classifier", { x: 0.4, y: 0, w: 9, h: 1.0, fontSize: 28, color: C.white, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.0, w: 10, h: 0.06, fill: { color: C.teal } });
 
    // Why Naive Bayes
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.15, w: 4.5, h: 4.1, fill: { color: "0D2137" }, shadow: makeShadow() });
    s.addText("WHY NAIVE BAYES?", { x: 0.3, y: 1.15, w: 4.5, h: 0.45, fontSize: 12, color: C.teal, bold: true, align: "center", valign: "middle", margin: 0 });
 
    const reasons = [
      ["⚡", "Fast & computationally efficient"],
      ["📐", "Works well with high-dim text features"],
      ["🎯", "Strong baseline for text classification"],
      ["🔢", "Probabilistic — gives confidence scores"],
      ["📊", "Effective even with moderate dataset sizes"],
    ];
    reasons.forEach(([icon, txt], i) => {
      s.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 1.72 + i * 0.65, w: 4.2, h: 0.55, fill: { color: "132232" } });
      s.addText(icon + "  " + txt, { x: 0.45, y: 1.72 + i * 0.65, w: 4.2, h: 0.55, fontSize: 12, color: C.white, valign: "middle", margin: 10 });
    });
 
    // Tools used
    s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.15, w: 4.5, h: 4.1, fill: { color: "0D2137" }, shadow: makeShadow() });
    s.addText("TOOLS & LIBRARIES", { x: 5.2, y: 1.15, w: 4.5, h: 0.45, fontSize: 12, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
 
    const tools = [
      ["🐍", "Python", "Core language"],
      ["🔤", "NLTK", "Text preprocessing"],
      ["🤖", "Scikit-Learn", "MultinomialNB, TF-IDF"],
      ["🐼", "Pandas", "Data manipulation"],
      ["🌐", "Streamlit", "GUI deployment"],
    ];
    tools.forEach(([icon, name, desc], i) => {
      s.addShape(pres.shapes.RECTANGLE, { x: 5.35, y: 1.72 + i * 0.65, w: 4.2, h: 0.55, fill: { color: "132232" } });
      s.addText(icon + "  ", { x: 5.35, y: 1.72 + i * 0.65, w: 0.5, h: 0.55, fontSize: 16, valign: "middle", margin: 8 });
      s.addText(name, { x: 5.75, y: 1.74 + i * 0.65, w: 1.3, h: 0.25, fontSize: 11, color: C.teal, bold: true, margin: 0 });
      s.addText(desc, { x: 5.75, y: 1.99 + i * 0.65, w: 3.5, h: 0.22, fontSize: 9.5, color: C.lightGray, margin: 0 });
    });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 6: RESULTS — EXPECTED VS ACTUAL
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
 
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
    s.addText("Results: Expected vs. Actual", { x: 0.4, y: 0, w: 9, h: 1.0, fontSize: 28, color: C.white, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.0, w: 10, h: 0.06, fill: { color: C.teal } });
 
    // Metrics table header
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.2, w: 9.4, h: 0.5, fill: { color: C.navyMid } });
    ["Metric", "Expected", "Actual (Test Set)", "Status"].forEach((h, i) => {
      const xs = [0.35, 2.85, 5.35, 7.85];
      const ws = [2.4, 2.4, 2.4, 1.8];
      s.addText(h, { x: xs[i], y: 1.2, w: ws[i], h: 0.5, fontSize: 12, color: C.teal, bold: true, valign: "middle", margin: 8 });
    });
 
    const rows = [
      ["Overall Accuracy", "~85%", "91.25%", "✅ Exceeded"],
      ["Precision (Pos)", "~90%", "94.13%", "✅ Exceeded"],
      ["Recall (Neg)", "~85%", "96%", "✅ Exceeded"],
      ["F1-Score (Pos)", "~0.93", "0.82*", "⚠️ Near target"],
      ["F1-Score (Neg)", "~0.72", "0.96", "✅ Exceeded"],
    ];
 
    rows.forEach(([metric, exp, act, status], i) => {
      const bg = i % 2 === 0 ? C.white : "EEF2F7";
      s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 1.75 + i * 0.55, w: 9.4, h: 0.52, fill: { color: bg } });
      s.addText(metric, { x: 0.35, y: 1.75 + i * 0.55, w: 2.4, h: 0.52, fontSize: 12, color: C.darkText, bold: true, valign: "middle", margin: 8 });
      s.addText(exp, { x: 2.85, y: 1.75 + i * 0.55, w: 2.4, h: 0.52, fontSize: 12, color: "64748B", valign: "middle", margin: 8 });
      s.addText(act, { x: 5.35, y: 1.75 + i * 0.55, w: 2.4, h: 0.52, fontSize: 12, color: C.tealDark, bold: true, valign: "middle", margin: 8 });
      const sc = status.includes("✅") ? C.green : C.accent;
      s.addText(status, { x: 7.85, y: 1.75 + i * 0.55, w: 1.8, h: 0.52, fontSize: 11, color: sc, bold: true, valign: "middle", margin: 8 });
    });
 
    s.addText("* Positive F1 lower due to class imbalance in recall (0.72 recall vs 0.96 precision). Confusion matrix: [[1618, 124], [222, 1989]]",
      { x: 0.3, y: 4.7, w: 9.4, h: 0.4, fontSize: 9.5, color: "64748B", italic: true });
 
    // Confusion matrix mini visual
    s.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 5.05, w: 9.4, h: 0.4, fill: { color: "E0F2FE" } });
    s.addText("Confusion Matrix: TN=1618  FP=124  FN=222  TP=1989   →   Strong true positive/negative detection", {
      x: 0.4, y: 5.05, w: 9.2, h: 0.4, fontSize: 11, color: C.tealDark, bold: true, valign: "middle", margin: 0
    });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 7: PERFORMANCE CHARTS
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
 
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navyMid } });
    s.addText("Performance Visualization", { x: 0.4, y: 0, w: 9, h: 1.0, fontSize: 28, color: C.white, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.0, w: 10, h: 0.06, fill: { color: C.teal } });
 
    // Bar chart — F1 by class
    s.addChart(pres.charts.BAR, [
      { name: "Precision", labels: ["Positive", "Negative"], values: [0.96, 0.97] },
      { name: "Recall", labels: ["Positive", "Negative"], values: [0.72, 0.96] },
      { name: "F1-Score", labels: ["Positive", "Negative"], values: [0.82, 0.96] },
    ], {
      x: 0.3, y: 1.15, w: 5.8, h: 3.8, barDir: "col",
      chartColors: ["00B4D8", "F59E0B", "10B981"],
      chartArea: { fill: { color: "0D2137" } },
      catAxisLabelColor: "BDE0FE",
      valAxisLabelColor: "BDE0FE",
      valGridLine: { color: "1E3A5F", size: 0.5 },
      catGridLine: { style: "none" },
      showValue: true,
      dataLabelColor: "FFFFFF",
      dataLabelFontSize: 9,
      showLegend: true,
      legendColor: "FFFFFF",
      legendPos: "b",
      legendFontSize: 10,
      showTitle: true,
      title: "Precision / Recall / F1 by Class",
      titleColor: "00B4D8",
      titleFontSize: 12,
    });
 
    // Pie chart — accuracy
    s.addChart(pres.charts.PIE, [{
      name: "Accuracy", labels: ["Correct", "Incorrect"], values: [91.25, 8.75]
    }], {
      x: 6.3, y: 1.15, w: 3.5, h: 3.8,
      chartColors: ["10B981", "EF4444"],
      showPercent: true,
      showTitle: true,
      title: "Overall Accuracy (91.25%)",
      titleColor: "00B4D8",
      titleFontSize: 12,
      chartArea: { fill: { color: "0D2137" } },
      legendColor: "FFFFFF",
      legendPos: "b",
      legendFontSize: 10,
    });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 8: LIMITATIONS
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.offWhite };
 
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navy } });
    s.addText("Limitations of Our System", { x: 0.4, y: 0, w: 9, h: 1.0, fontSize: 28, color: C.white, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.0, w: 10, h: 0.06, fill: { color: C.red } });
 
    const lims = [
      {
        icon: "🎭", title: "Cannot Detect Sarcasm & Negation",
        desc: "\"This movie was NOT bad at all\" → misclassified as negative. The word independence assumption ignores sentence structure entirely."
      },
      {
        icon: "📌", title: "Domain Dependency",
        desc: "Trained only on movie reviews. Performance degrades significantly on product reviews, tweets, or news without retraining."
      },
      {
        icon: "⚖️", title: "Class Imbalance Impact",
        desc: "Positive class has lower recall (0.72) because the dataset has more negative samples, biasing predictions toward negative."
      },
      {
        icon: "🔤", title: "Feature Simplicity",
        desc: "Bag-of-Words ignores word order, context, and semantics. 'Great but boring' is treated the same as two separate bag words."
      },
      {
        icon: "🌍", title: "English Only",
        desc: "No multilingual support. International reviews in other languages are completely unhandled by the current pipeline."
      },
      {
        icon: "📉", title: "Static Model",
        desc: "Language evolves. Slang, new terminology, and shifting sentiment patterns require periodic retraining to maintain accuracy."
      },
    ];
 
    lims.forEach((lim, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.3 + col * 4.85;
      const y = 1.2 + row * 1.4;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.6, h: 1.25, fill: { color: C.white }, shadow: makeShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.08, h: 1.25, fill: { color: C.red } });
      s.addText(lim.icon + "  " + lim.title, { x: x + 0.18, y: y + 0.1, w: 4.3, h: 0.35, fontSize: 12, color: C.darkText, bold: true, margin: 0 });
      s.addText(lim.desc, { x: x + 0.18, y: y + 0.45, w: 4.3, h: 0.75, fontSize: 9.5, color: "475569", margin: 0 });
    });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 9: ADVANCED IMPROVEMENTS (BERT/RAG)
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
 
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: C.navyMid } });
    s.addText("Proposed Improvements", { x: 0.4, y: 0, w: 9, h: 1.0, fontSize: 28, color: C.white, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 1.0, w: 10, h: 0.06, fill: { color: C.accent } });
 
    // BERT Card
    s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 1.15, w: 3.0, h: 4.1, fill: { color: "0A1628" }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.25, y: 1.15, w: 3.0, h: 0.55, fill: { color: "1D4ED8" } });
    s.addText("🤖  BERT", { x: 0.25, y: 1.15, w: 3.0, h: 0.55, fontSize: 14, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText("Bidirectional Encoder Representations from Transformers", { x: 0.35, y: 1.78, w: 2.8, h: 0.5, fontSize: 9, color: "93C5FD", italic: true, align: "center" });
 
    const bertPoints = ["Context-aware word embeddings", "Handles negation & sarcasm", "Pre-trained on 3.3B words", "Fine-tune on movie reviews", "State-of-the-art NLP accuracy"];
    bertPoints.forEach((p, i) => {
      s.addText("▸  " + p, { x: 0.35, y: 2.38 + i * 0.52, w: 2.8, h: 0.45, fontSize: 10.5, color: C.white, margin: 4 });
    });
 
    // Word2Vec Card
    s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 1.15, w: 3.0, h: 4.1, fill: { color: "0A1628" }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: 1.15, w: 3.0, h: 0.55, fill: { color: "047857" } });
    s.addText("📐  Word2Vec / GloVe", { x: 3.5, y: 1.15, w: 3.0, h: 0.55, fontSize: 13, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText("Dense vector embeddings capturing semantic similarity", { x: 3.6, y: 1.78, w: 2.8, h: 0.5, fontSize: 9, color: "6EE7B7", italic: true, align: "center" });
 
    const w2vPoints = ["Words with similar meaning cluster together", "Replaces sparse BoW with rich 300-d vectors", "Captures analogy relationships", "Integrate with LSTM / CNN", "Massive reduction in feature noise"];
    w2vPoints.forEach((p, i) => {
      s.addText("▸  " + p, { x: 3.6, y: 2.38 + i * 0.52, w: 2.8, h: 0.45, fontSize: 10.5, color: C.white, margin: 4 });
    });
 
    // RAG Card
    s.addShape(pres.shapes.RECTANGLE, { x: 6.75, y: 1.15, w: 3.0, h: 4.1, fill: { color: "0A1628" }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.75, y: 1.15, w: 3.0, h: 0.55, fill: { color: "7C3AED" } });
    s.addText("🔍  RAG System", { x: 6.75, y: 1.15, w: 3.0, h: 0.55, fontSize: 14, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText("Retrieval-Augmented Generation for contextual analysis", { x: 6.85, y: 1.78, w: 2.8, h: 0.5, fontSize: 9, color: "C4B5FD", italic: true, align: "center" });
 
    const ragPoints = ["Retrieve similar past reviews", "Augment context before classification", "LLM generates nuanced explanation", "Handles slang & evolving language", "Real-time, no retraining needed"];
    ragPoints.forEach((p, i) => {
      s.addText("▸  " + p, { x: 6.85, y: 2.38 + i * 0.52, w: 2.8, h: 0.45, fontSize: 10.5, color: C.white, margin: 4 });
    });
  }
 
  // ─────────────────────────────────────────
  // SLIDE 10: CONCLUSION
  // ─────────────────────────────────────────
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };
 
    // Decorative circle BG
    s.addShape(pres.shapes.OVAL, { x: 6.5, y: -1, w: 5, h: 5, fill: { color: C.tealDark, transparency: 80 } });
    s.addShape(pres.shapes.OVAL, { x: -1, y: 3, w: 4, h: 4, fill: { color: "1D4ED8", transparency: 85 } });
 
    s.addText("Conclusion", { x: 0.5, y: 0.5, w: 9, h: 0.75, fontSize: 38, color: C.white, bold: true, fontFace: "Calibri", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.3, w: 4.0, h: 0.04, fill: { color: C.teal } });
 
    const bullets = [
      ["✅", "Naive Bayes achieved 91.25% accuracy on 13K+ reviews"],
      ["✅", "Effective text pipeline: cleaning → TF-IDF → probabilistic model"],
      ["✅", "Interactive Streamlit GUI for real-time prediction"],
      ["⚠️", "Positive recall (0.72) limited by class imbalance & no context-awareness"],
      ["🚀", "Next: BERT for contextual embeddings, RAG for dynamic retrieval"],
      ["🌍", "Future: Multilingual support & real-time social media analytics"],
    ];
 
    bullets.forEach(([icon, txt], i) => {
      const color = icon === "⚠️" ? C.accent : icon === "🚀" || icon === "🌍" ? C.teal : C.white;
      s.addText(icon + "  " + txt, { x: 0.5, y: 1.5 + i * 0.62, w: 9, h: 0.55, fontSize: 13.5, color, fontFace: "Calibri", margin: 0 });
    });
 
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 5.1, w: 9, h: 0.04, fill: { color: C.teal, transparency: 50 } });
    s.addText("Muhammad Haroon (052)  &  Mohammad Sami (063)  |  Bahria University  |  NLP — 2026",
      { x: 0.5, y: 5.18, w: 9, h: 0.35, fontSize: 9.5, color: "64748B", fontFace: "Calibri" });
  }
 
await pres.writeFile({ fileName: "NLP_Sentiment_Analysis_Presentation.pptx" });
  console.log("✅ Presentation created successfully!");
}
 
createPresentation().catch(console.error);
