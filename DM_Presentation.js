const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// Icon imports
const { FaHospital, FaShieldAlt, FaGraduationCap, FaChartLine, FaDatabase,
        FaBrain, FaLock, FaNetworkWired, FaSearch, FaLightbulb, FaExclamationTriangle,
        FaCheckCircle, FaArrowRight, FaUniversity, FaRobot } = require("react-icons/fa");
const { MdSecurity, MdSpeed, MdSchool, MdMemory } = require("react-icons/md");

// ── PALETTE ────────────────────────────────────────────────────────────────
const C = {
  darkBg:   "0D1B2A",   // deep navy
  midBg:    "1B2A3B",   // slate
  accent1:  "00B4D8",   // cyan
  accent2:  "0096C7",   // teal-blue
  accent3:  "023E8A",   // deep blue
  white:    "FFFFFF",
  offWhite: "E0F4FF",
  muted:    "8ECAE6",
  gold:     "F4A261",
  red:      "E63946",
  green:    "52B788",
  lightBg:  "F0F8FF",
  cardBg:   "FFFFFF",
  textDark: "0D1B2A",
  cardBorder: "00B4D8",
};

async function iconPng(IconComp, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComp, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ── HELPERS ────────────────────────────────────────────────────────────────
function makeShadow() {
  return { type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.18 };
}

function sectionTag(slide, label, color, x = 0.32, y = 1.02) {
  slide.addShape("rect", { x, y, w: 1.1, h: 0.22, fill: { color } });
  slide.addText(label, { x, y, w: 1.1, h: 0.22, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
}

// ── PRESENTATION ──────────────────────────────────────────────────────────
async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Abdul Moeez";
  pres.title = "Data Mining – Literature Review";

  // ═══════════════════════════════════════════════════════════
  // SLIDE 1 – TITLE
  // ═══════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };

    // Left accent stripe
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.accent1 } });

    // Stats band – top right
    const stats = [["11", "Research Papers"], ["5", "Domains"], ["2021–25", "Time Range"]];
    stats.forEach(([num, lbl], i) => {
      const bx = 7.2 + i * 0.92;
      s.addShape("rect", { x: bx, y: 0.18, w: 0.82, h: 0.72, fill: { color: C.accent3 }, shadow: makeShadow() });
      s.addText(num, { x: bx, y: 0.18, w: 0.82, h: 0.44, fontSize: 18, bold: true, color: C.accent1, align: "center", valign: "bottom", margin: 0 });
      s.addText(lbl, { x: bx, y: 0.55, w: 0.82, h: 0.3, fontSize: 6.5, color: C.muted, align: "center", valign: "top", margin: 0 });
    });

    // Main title
    s.addText("Literature Review on", { x: 0.55, y: 1.2, w: 9, h: 0.55, fontSize: 22, color: C.muted, italic: true, margin: 0 });
    s.addText("Data Mining", { x: 0.55, y: 1.75, w: 9, h: 1.0, fontSize: 54, bold: true, color: C.white, margin: 0 });
    s.addText("Emerging Issues & Solutions Across Domains", {
      x: 0.55, y: 2.75, w: 9, h: 0.5, fontSize: 18, color: C.accent1, italic: false, margin: 0,
    });

    // Divider
    s.addShape("rect", { x: 0.55, y: 3.32, w: 6.5, h: 0.03, fill: { color: C.accent2 } });

    // Author block
    s.addText([
      { text: "Abdul Moeez", options: { bold: true, color: C.white } },
      { text: "  ·  01-134231-002  ·  BSCS 7B  ·  Bahria University, Islamabad  ·  Assignment 3", options: { color: C.muted } },
    ], { x: 0.55, y: 3.48, w: 9, h: 0.35, fontSize: 11, margin: 0 });

    // Visual hint
    s.addText("🔍  IEEE  ·  Springer  ·  Wiley  ·  Elsevier  ·  ACM", {
      x: 0.55, y: 4.9, w: 6, h: 0.35, fontSize: 9.5, color: C.muted, italic: true, margin: 0,
    });

    // IMAGE GUIDE (watermark text)
    s.addText("[ Visual: Flat vector of interconnected data nodes – right watermark, 15% opacity ]", {
      x: 5.5, y: 4.2, w: 4.2, h: 0.5, fontSize: 7, color: "2A4060", italic: true, align: "right", margin: 0,
    });
  }

  // ═══════════════════════════════════════════════════════════
  // SLIDE 2 – OVERVIEW: 5 DOMAINS
  // ═══════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.accent1 } });

    s.addText("Overview", { x: 0.35, y: 0.2, w: 9, h: 0.52, fontSize: 28, bold: true, color: C.white, margin: 0 });
    s.addText("5 Domains · 11 Papers · 2021–2025", { x: 0.35, y: 0.72, w: 9, h: 0.3, fontSize: 12, color: C.accent1, italic: true, margin: 0 });

    const domains = [
      { label: "Healthcare", sub: "Privacy · EHR · Ethics", refs: "J1, J2, J3", color: C.green },
      { label: "Finance", sub: "Fraud Detection · Fintech", refs: "J4, J5", color: C.gold },
      { label: "Scalability", sub: "GNNs · Distributed", refs: "J6, J7", color: C.accent1 },
      { label: "Privacy", sub: "GDPR · Unlearning", refs: "J8, J9", color: C.red },
      { label: "Edu & Survey", sub: "EdTech · IoT Buildings", refs: "J10, J11", color: "A78BFA" },
    ];

    domains.forEach((d, i) => {
      const bx = 0.28 + i * 1.9;
      s.addShape("rect", { x: bx, y: 1.12, w: 1.75, h: 3.9, fill: { color: C.midBg }, shadow: makeShadow() });
      s.addShape("rect", { x: bx, y: 1.12, w: 1.75, h: 0.08, fill: { color: d.color } });
      s.addText(d.label, { x: bx, y: 1.24, w: 1.75, h: 0.42, fontSize: 13, bold: true, color: C.white, align: "center", margin: 0 });
      s.addText(d.sub, { x: bx, y: 1.7, w: 1.75, h: 0.55, fontSize: 9, color: C.muted, align: "center", margin: 0 });
      s.addShape("rect", { x: bx + 0.4, y: 2.35, w: 0.95, h: 0.01, fill: { color: d.color } });
      s.addText(d.refs, { x: bx, y: 2.45, w: 1.75, h: 0.38, fontSize: 12, bold: true, color: d.color, align: "center", margin: 0 });
    });

    s.addText("[ Visual: Radar/hub-and-spoke diagram showing 5 domains – center of slide ]", {
      x: 0.3, y: 5.2, w: 9, h: 0.28, fontSize: 7, color: "2A4060", italic: true, margin: 0,
    });
  }

  // ═══════════════════════════════════════════════════════════
  // SLIDES 3–6: JOURNAL SUMMARIES (J1–J11)
  // ═══════════════════════════════════════════════════════════

  const journalSlides = [
    {
      title: "Research Articles – J1 · J2 · J3  (Healthcare)",
      accentColor: C.green,
      papers: [
        {
          tag: "J1 · Healthcare",
          title: "Data Mining in Healthcare: Overview, Techniques & Challenges",
          cite: "Ahmed et al. (2024) – IJAIMS",
          bottleneck: "Curse of Dimensionality in EHR / medical imaging",
          technique: "Clinical Decision Support + Feature Engineering",
          bullets: [
            "Missing data → life-threatening diagnostic errors",
            "High-dim genomic data: features >> samples",
            "EHR heterogeneity across hospital systems",
          ],
          visual: "Flat icon: medical cross on a circuit board – top-right corner [Search: 'flat vector medical data mining icon']",
        },
        {
          tag: "J2 · Healthcare",
          title: "The Ethics of Data Mining in Healthcare",
          cite: "Ahmed et al. (2025) – Springer",
          bottleneck: "Consent-by-default inadequacy & breach risk",
          technique: "Ethical AI Frameworks + Transparent Auditing",
          bullets: [
            "Data breaches: privacy risk beyond anonymisation",
            "Black-box models unacceptable in clinical settings",
            "Calls for GDPR-aligned, patient-centric governance",
          ],
          visual: "Icon: scales of justice over a shield – right third [Search: 'flat vector ethics AI healthcare icon']",
        },
        {
          tag: "J3 · Healthcare",
          title: "Privacy-Preserving Data Mining & ML in Healthcare",
          cite: "Naresh & Thamarai (2023) – Wiley WIDM",
          bottleneck: "Centralising sensitive patient records",
          technique: "Federated Learning + Homomorphic Encryption",
          def: "FL: models train locally on devices — raw data never leaves the source.\nHE: computes on encrypted data without ever decrypting it.",
          bullets: [
            "Adversarial data poisoning degrades AI diagnosis",
            "HE overhead: significant latency vs. plaintext mining",
            "FL enables multi-hospital collaboration without data sharing",
          ],
          visual: "Icon: padlock over hospital building – top-right [Search: 'federated learning decentralised hospital flat icon']",
        },
      ],
    },
    {
      title: "Research Articles – J4 · J5 · J6  (Finance & Scalability)",
      accentColor: C.gold,
      papers: [
        {
          tag: "J4 · Finance",
          title: "Intelligent Fraud Detection in Financial Statements",
          cite: "Jafari et al. (2021) – IEEE Access",
          bottleneck: "Extreme class imbalance — fraud is rare",
          technique: "Hybrid Models (RF + Neural Nets) + SMOTE",
          def: "SMOTE: generates synthetic minority samples to balance training set.",
          bullets: [
            "Conventional classifiers achieve high accuracy by ignoring fraud",
            "SMOTE boosts recall on minority fraud class [4]",
            "Risk: synthetic data may not reflect novel fraud patterns",
          ],
          visual: "Icon: bar chart with anomaly spike – right third [Search: 'flat vector fraud detection financial anomaly icon']",
        },
        {
          tag: "J5 · Finance",
          title: "Bibliometric Analysis of DM & ML in Fintech",
          cite: "Kumar et al. (2021) – DSFE",
          bottleneck: "Volatility & non-linearity of Fintech data",
          technique: "Ensemble Learning + Credit Scoring Models",
          bullets: [
            "Models overfit to noise rather than market signals",
            "Real-time risk assessment requires online learning",
            "Over-reliance on synthetic data → brittle generalization",
          ],
          visual: "Icon: network graph with dollar sign – right [Search: 'flat vector fintech machine learning network icon']",
        },
        {
          tag: "J6 · Scalability",
          title: "Deep Learning on Graphs: A Survey",
          cite: "Zhang et al. (2022) – ACM TIST",
          bottleneck: "Neighbourhood Explosion — exponential cost per GNN layer",
          technique: "Graph Convolutional Networks (GCNs)",
          bullets: [
            "Traditional DL cannot handle billions of nodes/edges",
            "GNNs capture relational patterns in financial/social nets",
            "Polynomial/exponential complexity limits real-time use",
          ],
          visual: "Icon: interconnected nodes / graph structure – right third [Search: 'graph neural network flat vector icon blue']",
        },
      ],
    },
    {
      title: "Research Articles – J7 · J8 · J9  (Scalability & Privacy)",
      accentColor: C.red,
      papers: [
        {
          tag: "J7 · Scalability",
          title: "Scalable DM: Parallel & Distributed Approaches",
          cite: "Bekkerman et al. (2021) – IEEE",
          bottleneck: "Inter-node Communication Bottleneck in clusters",
          technique: "MapReduce + Apache Spark Parallelisation",
          bullets: [
            "Synchronisation delay grows with cluster size",
            "Spark handles datasets exceeding single-machine RAM",
            "Communication overhead often outweighs compute savings",
          ],
          visual: "Icon: server rack with connecting arrows – right [Search: 'flat vector distributed computing cluster server icon']",
        },
        {
          tag: "J8 · Privacy",
          title: "Privacy-Preserving Data Stream Mining",
          cite: "Hewage et al. (2023) – AI Review (Springer)",
          bottleneck: "Privacy–Utility Trade-off on streaming data",
          technique: "k-Anonymity + Differential Privacy (DP)",
          def: "DP: adds calibrated noise to outputs providing mathematical privacy guarantee.",
          bullets: [
            "Streaming data: noise removal must not sacrifice velocity",
            "Every dB of privacy protection measurably drops accuracy [8]",
            "k-Anonymity fails against background-knowledge attacks",
          ],
          visual: "Icon: stream of data with a shield overlay – right [Search: 'flat vector data stream privacy protection icon']",
        },
        {
          tag: "J9 · Privacy",
          title: "Machine Learning & the Right to Be Forgotten",
          cite: "Voigt et al. (2022) – CLSR (Elsevier)",
          bottleneck: "Technical impossibility of selective model unlearning",
          technique: "SISA Training (Sharded, Isolated, Sliced, Aggregated)",
          def: "SISA: partitions data into shards — only the affected shard is retrained to 'forget' a user.",
          bullets: [
            "GDPR mandates deletion, but weights bake in user influence [9]",
            "SISA enables fast deletion without full retraining",
            "Limitation: significant storage overhead for shard maintenance",
          ],
          visual: "Icon: eraser over a neural network diagram – right third [Search: 'machine unlearning GDPR right to forget flat icon']",
        },
      ],
    },
    {
      title: "Research Articles – J10 · J11  (Survey & Education)",
      accentColor: "A78BFA",
      papers: [
        {
          tag: "J10 · Survey / Automation",
          title: "AI & Big Data for Building Automation Systems",
          cite: "Himeur et al. (2022) – Journal of Big Data",
          bottleneck: "Heterogeneous IoT sensor integration & edge constraints",
          technique: "Edge / Fog Computing + Explainable AI (XAI)",
          bullets: [
            "Inconsistent sampling rates create dataset gaps",
            "Edge devices: limited battery & processing power",
            "XAI required for facility managers to trust AI decisions",
          ],
          visual: "Icon: smart building with IoT signal waves – right third [Search: 'smart building IoT sensor flat vector icon teal']",
        },
        {
          tag: "J11 · Education",
          title: "DM & ML in the Educational Sector",
          cite: "Arkhypova & Svydenko (2024) – KNUTD",
          bottleneck: "Data sparsity in student engagement logs",
          technique: "Decision Trees + Rule-Based Classifiers",
          bullets: [
            "Sparse logs → incomplete student behaviour profiles",
            "Rule-based models: interpretable for teacher intervention",
            "Decision Trees weaker than deep nets on raw accuracy",
            "Personalised learning paths improve student retention",
          ],
          visual: "Icon: graduation cap with a decision tree branching below – right [Search: 'education data mining decision tree flat icon']",
        },
      ],
    },
  ];

  for (const jSlide of journalSlides) {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: jSlide.accentColor } });
    s.addText(jSlide.title, { x: 0.25, y: 0.1, w: 9.5, h: 0.45, fontSize: 16, bold: true, color: C.white, margin: 0 });
    s.addShape("rect", { x: 0.25, y: 0.55, w: 9.5, h: 0.025, fill: { color: jSlide.accentColor } });

    const count = jSlide.papers.length;
    const cardW = count === 2 ? 4.55 : 3.0;
    const gap = count === 2 ? 0.4 : 0.22;

    jSlide.papers.forEach((p, i) => {
      const bx = 0.22 + i * (cardW + gap);
      const by = 0.65;
      const bh = 4.8;

      // Card background
      s.addShape("rect", { x: bx, y: by, w: cardW, h: bh, fill: { color: C.midBg }, shadow: makeShadow() });
      s.addShape("rect", { x: bx, y: by, w: cardW, h: 0.06, fill: { color: jSlide.accentColor } });

      // Tag
      s.addShape("rect", { x: bx + 0.08, y: by + 0.1, w: 1.4, h: 0.2, fill: { color: jSlide.accentColor } });
      s.addText(p.tag, { x: bx + 0.08, y: by + 0.1, w: 1.4, h: 0.2, fontSize: 7, bold: true, color: C.darkBg, align: "center", valign: "middle", margin: 0 });

      // Title
      s.addText(p.title, { x: bx + 0.08, y: by + 0.38, w: cardW - 0.16, h: 0.65, fontSize: 9.5, bold: true, color: C.white, margin: 0 });

      // Citation
      s.addText(p.cite, { x: bx + 0.08, y: by + 1.02, w: cardW - 0.16, h: 0.22, fontSize: 7.5, color: C.muted, italic: true, margin: 0 });

      // Bottleneck row
      s.addText("PRIMARY BOTTLENECK", { x: bx + 0.08, y: by + 1.28, w: cardW - 0.16, h: 0.18, fontSize: 6.5, bold: true, color: jSlide.accentColor, margin: 0 });
      s.addText(p.bottleneck, { x: bx + 0.08, y: by + 1.45, w: cardW - 0.16, h: 0.3, fontSize: 8.5, bold: true, color: C.white, margin: 0 });

      // Key technique row
      s.addText("KEY TECHNIQUE", { x: bx + 0.08, y: by + 1.78, w: cardW - 0.16, h: 0.18, fontSize: 6.5, bold: true, color: C.gold, margin: 0 });
      s.addText(p.technique, { x: bx + 0.08, y: by + 1.95, w: cardW - 0.16, h: 0.3, fontSize: 8.5, bold: true, color: C.gold, margin: 0 });

      // Mechanical definition
      if (p.def) {
        s.addShape("rect", { x: bx + 0.08, y: by + 2.3, w: cardW - 0.16, h: 0.02, fill: { color: C.accent3 } });
        s.addText(p.def, { x: bx + 0.08, y: by + 2.35, w: cardW - 0.16, h: 0.48, fontSize: 7.2, color: C.offWhite, italic: true, margin: 0 });
      }

      // Bullets
      const bulletY = p.def ? by + 2.88 : by + 2.38;
      const bullets = p.bullets.map((b, bi) => ({
        text: b,
        options: { bullet: true, color: C.muted, fontSize: 8, breakLine: bi < p.bullets.length - 1 },
      }));
      s.addText(bullets, { x: bx + 0.08, y: bulletY, w: cardW - 0.16, h: bh - bulletY + by - 0.15, margin: 0 });

      // Visual guide at bottom of card
      s.addText(`🖼 ${p.visual}`, {
        x: bx + 0.08, y: by + bh - 0.32, w: cardW - 0.16, h: 0.28,
        fontSize: 5.8, color: "2A5070", italic: true, margin: 0,
      });
    });
  }

  // ═══════════════════════════════════════════════════════════
  // SLIDE 7 – KEY ISSUES
  // ═══════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.accent1 } });

    s.addText("Key Data Mining Issues", { x: 0.25, y: 0.1, w: 7, h: 0.48, fontSize: 24, bold: true, color: C.white, margin: 0 });
    s.addText("Identified across all 11 papers  [1–11]", { x: 0.25, y: 0.57, w: 7, h: 0.28, fontSize: 10, color: C.accent1, italic: true, margin: 0 });

    const issues = [
      { num: "01", title: "Data Preprocessing & Cleaning", color: C.green,
        desc: "EHR heterogeneity · missing values · streaming noise · IoT inconsistencies · student log sparsity" },
      { num: "02", title: "High-Dimensional & Large-Scale Data", color: C.accent1,
        desc: "Curse of Dimensionality in genomics · web-scale graph bottlenecks · cluster communication lag" },
      { num: "03", title: "Model Accuracy & Overfitting", color: C.gold,
        desc: "Class imbalance in fraud datasets · Fintech market volatility · Privacy–Utility trade-off" },
      { num: "04", title: "Privacy & Security Concerns", color: C.red,
        desc: "Data breaches · GDPR Right-to-Forget · adversarial data poisoning in healthcare AI" },
      { num: "05", title: "Computational Complexity", color: "A78BFA",
        desc: "Polynomial GNN cost · expensive model retraining for unlearning · edge device constraints" },
    ];

    issues.forEach((iss, i) => {
      const bx = 0.22;
      const by = 0.95 + i * 0.9;
      s.addShape("rect", { x: bx, y: by, w: 0.5, h: 0.65, fill: { color: iss.color } });
      s.addText(iss.num, { x: bx, y: by, w: 0.5, h: 0.65, fontSize: 16, bold: true, color: C.darkBg, align: "center", valign: "middle", margin: 0 });
      s.addShape("rect", { x: bx + 0.5, y: by, w: 9.0, h: 0.65, fill: { color: C.midBg }, shadow: makeShadow() });
      s.addText(iss.title, { x: bx + 0.6, y: by + 0.02, w: 4.5, h: 0.3, fontSize: 11, bold: true, color: C.white, valign: "middle", margin: 0 });
      s.addText(iss.desc, { x: bx + 0.6, y: by + 0.32, w: 8.8, h: 0.28, fontSize: 8, color: C.muted, margin: 0 });
    });

    s.addText("[ Visual: Search 'five layers data pipeline flat icon' – right watermark 10% opacity ]", {
      x: 0.25, y: 5.3, w: 9, h: 0.22, fontSize: 6.5, color: "2A4060", italic: true, margin: 0,
    });
  }

  // ═══════════════════════════════════════════════════════════
  // SLIDE 8 – PROPOSED SOLUTIONS
  // ═══════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.accent1 } });

    s.addText("Proposed Solutions by Domain", { x: 0.25, y: 0.1, w: 9.5, h: 0.48, fontSize: 24, bold: true, color: C.white, margin: 0 });
    s.addShape("rect", { x: 0.25, y: 0.58, w: 9.5, h: 0.025, fill: { color: C.accent1 } });

    const solutions = [
      { domain: "Healthcare [1, 2, 3]", color: C.green,
        tech: "Federated Learning + Homomorphic Encryption",
        pro: "Multi-hospital training without sharing patient records",
        con: "High encryption overhead slows mining throughput" },
      { domain: "Finance [4, 5]", color: C.gold,
        tech: "Hybrid Models (RF + Neural Nets) + SMOTE",
        pro: "SMOTE boosts fraud recall on imbalanced datasets",
        con: "Synthetic data risks overfitting to known fraud patterns" },
      { domain: "Scalability [6, 7]", color: C.accent1,
        tech: "Graph Convolutional Networks + MapReduce / Spark",
        pro: "Handles billions of nodes in social & financial graphs",
        con: "Neighbourhood Explosion & synchronisation overhead" },
      { domain: "Privacy [8, 9]", color: C.red,
        tech: "Differential Privacy + SISA Unlearning",
        pro: "Math. privacy guarantee; fast GDPR-compliant deletion",
        con: "Noise reduces accuracy; SISA requires extra shard storage" },
      { domain: "Edu & Survey [10, 11]", color: "A78BFA",
        tech: "Decision Trees + Explainable AI (XAI) + Edge Computing",
        pro: "Interpretable student predictions; low-latency IoT inference",
        con: "Decision Trees weaker than deep nets; edge power limited" },
    ];

    solutions.forEach((sol, i) => {
      const bx = 0.22;
      const by = 0.68 + i * 0.93;
      s.addShape("rect", { x: bx, y: by, w: 1.7, h: 0.8, fill: { color: sol.color } });
      s.addText(sol.domain, { x: bx, y: by, w: 1.7, h: 0.8, fontSize: 8.5, bold: true, color: C.darkBg, align: "center", valign: "middle", margin: 4 });

      s.addShape("rect", { x: bx + 1.7, y: by, w: 7.85, h: 0.8, fill: { color: C.midBg } });
      s.addText(sol.tech, { x: bx + 1.82, y: by + 0.02, w: 7.6, h: 0.3, fontSize: 9.5, bold: true, color: C.white, margin: 0 });
      s.addText([
        { text: "✔ ", options: { bold: true, color: C.green } },
        { text: sol.pro + "  ", options: { color: C.muted } },
        { text: "  ⚠ ", options: { bold: true, color: C.gold } },
        { text: sol.con, options: { color: C.muted } },
      ], { x: bx + 1.82, y: by + 0.36, w: 7.6, h: 0.35, fontSize: 8, margin: 0 });
    });

    s.addText("[ Visual: Search 'solutions toolkit flat vector icon' – right margin small icon ]", {
      x: 0.25, y: 5.36, w: 9, h: 0.22, fontSize: 6.5, color: "2A4060", italic: true, margin: 0,
    });
  }

  // ═══════════════════════════════════════════════════════════
  // SLIDE 9 – COMPARATIVE ANALYSIS
  // ═══════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.accent1 } });

    s.addText("Comparative Analysis", { x: 0.25, y: 0.1, w: 9.5, h: 0.46, fontSize: 24, bold: true, color: C.white, margin: 0 });

    // Table header
    const cols = ["Domain", "Primary Issue", "Key Solution", "Risk Tolerance"];
    const colW = [1.6, 2.4, 2.8, 2.0];
    const hdrX = [0.22, 1.82, 4.22, 7.02];

    hdrX.forEach((x, i) => {
      s.addShape("rect", { x, y: 0.65, w: colW[i] - 0.05, h: 0.32, fill: { color: C.accent3 } });
      s.addText(cols[i], { x, y: 0.65, w: colW[i] - 0.05, h: 0.32, fontSize: 9, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0 });
    });

    const rows = [
      ["Healthcare [1, 2, 3]", "Data Privacy & Ethics", "Federated Learning + Encryption", "Very High – patient lives", C.green],
      ["Finance [4, 5]", "Class Imbalance / Fraud", "Hybrid ML + SMOTE", "High – financial loss", C.gold],
      ["Scalability [6, 7]", "Computational Complexity", "GNNs + Distributed Compute", "Medium", C.accent1],
      ["Privacy [8, 9]", "GDPR Right-to-Forget", "Differential Privacy + SISA", "High – legal compliance", C.red],
      ["Edu & Survey [10, 11]", "Data Sparsity / IoT Noise", "Decision Trees + Edge AI + XAI", "Medium – retention", "A78BFA"],
    ];

    rows.forEach((row, ri) => {
      const by = 1.01 + ri * 0.72;
      const rowColor = ri % 2 === 0 ? C.midBg : "162438";
      hdrX.forEach((x, ci) => {
        s.addShape("rect", { x, y: by, w: colW[ci] - 0.05, h: 0.66, fill: { color: rowColor } });
        if (ci === 0) {
          s.addShape("rect", { x, y: by, w: 0.045, h: 0.66, fill: { color: row[4] } });
        }
        s.addText(row[ci], { x: x + (ci === 0 ? 0.08 : 0.06), y: by, w: colW[ci] - 0.16, h: 0.66, fontSize: 8.5, color: C.white, valign: "middle", margin: 0 });
      });
    });

    // Similarities / Differences
    s.addText([
      { text: "SIMILARITIES:  ", options: { bold: true, color: C.accent1 } },
      { text: "All domains grapple with missingness & noise while seeking interpretable, scalable models.", options: { color: C.muted } },
    ], { x: 0.22, y: 4.66, w: 9.5, h: 0.28, fontSize: 8.5, margin: 0 });
    s.addText([
      { text: "DIFFERENCES:  ", options: { bold: true, color: C.gold } },
      { text: "Finance → adversarial robustness  ·  Education → teacher feedback  ·  Healthcare → clinical verifiability", options: { color: C.muted } },
    ], { x: 0.22, y: 4.95, w: 9.5, h: 0.28, fontSize: 8.5, margin: 0 });

    s.addText("[ Visual: Search 'Venn diagram three overlapping circles flat icon' – right corner small ]", {
      x: 0.22, y: 5.3, w: 9, h: 0.22, fontSize: 6.5, color: "2A4060", italic: true, margin: 0,
    });
  }

  // ═══════════════════════════════════════════════════════════
  // SLIDE 10 – RESEARCH GAPS & FUTURE DIRECTIONS
  // ═══════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.accent1 } });

    s.addText("Research Gaps & Future Directions", { x: 0.25, y: 0.1, w: 9, h: 0.46, fontSize: 22, bold: true, color: C.white, margin: 0 });
    s.addShape("rect", { x: 0.25, y: 0.55, w: 9, h: 0.025, fill: { color: C.accent1 } });

    const gaps = [
      {
        icon: "⚙", color: C.green, num: "01",
        title: "Standardisation of Heterogeneous Data [1, 10]",
        body: "No global format for IoT & medical data — causes preprocessing rot before any mining begins. CV models (e.g. YOLO) trained on one sensor spec fail on another device's stream.",
      },
      {
        icon: "🧠", color: C.red, num: "02",
        title: "Machine Unlearning at Scale [9]",
        body: "SISA handles single-user deletion, but mass unlearning for thousands of simultaneous GDPR requests remains unsolved and prohibitively expensive.",
      },
      {
        icon: "🔗", color: C.gold, num: "03",
        title: "Cross-Domain Transferability [6, 11]",
        body: "Models are brittle — a fraud detector trained at one bank or a student-success predictor at one university fails elsewhere without expensive retraining.",
      },
      {
        icon: "🌱", color: "A78BFA", num: "04",
        title: "Energy-Efficient 'Green AI' [10]",
        body: "As GNNs and deep models scale, carbon footprint and energy costs become unsustainable for long-term IoT or mobile deployments.",
      },
      {
        icon: "📡", color: C.accent1, num: "05",
        title: "Real-Time Processing & Data Heterogeneity [1, 10] ← NEW",
        body: "Smartphone-based sensing (e.g. soil organic matter estimation via hyperspectral phone cameras) and YOLO-based CV pipelines demand sub-second inference on heterogeneous, unstructured data streams — a gap no current framework fully bridges.",
      },
    ];

    gaps.forEach((g, i) => {
      const bx = 0.22;
      const by = 0.65 + i * 0.96;
      s.addShape("rect", { x: bx, y: by, w: 0.55, h: 0.8, fill: { color: g.color } });
      s.addText(g.num, { x: bx, y: by, w: 0.55, h: 0.8, fontSize: 18, bold: true, color: C.darkBg, align: "center", valign: "middle", margin: 0 });
      s.addShape("rect", { x: bx + 0.55, y: by, w: 9.0, h: 0.8, fill: { color: C.midBg } });
      s.addText(g.title, { x: bx + 0.65, y: by + 0.02, w: 8.8, h: 0.28, fontSize: 9.5, bold: true, color: C.white, margin: 0 });
      s.addText(g.body, { x: bx + 0.65, y: by + 0.3, w: 8.75, h: 0.45, fontSize: 8, color: C.muted, margin: 0 });
    });

    s.addText("[ Visual: Search 'research roadmap future directions flat icon' – top-right corner ]", {
      x: 0.22, y: 5.38, w: 9, h: 0.2, fontSize: 6.5, color: "2A4060", italic: true, margin: 0,
    });
  }

  // ═══════════════════════════════════════════════════════════
  // SLIDE 11 – CONCLUSION
  // ═══════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.accent1 } });

    s.addText("Conclusion", { x: 0.25, y: 0.25, w: 9, h: 0.55, fontSize: 28, bold: true, color: C.white, margin: 0 });

    const concPoints = [
      { bold: "Hybrid techniques (FL + GNNs + Explainable AI (XAI))", rest: " are the frontier for privacy-aware, scalable data mining." },
      { bold: "Unified frameworks", rest: " bridging privacy and interpretability simultaneously remain the critical unmet need." },
      { bold: "Energy-efficient 'Green AI'", rest: " must underpin next-generation deployments at scale." },
      { bold: "Domain-specific risk tolerance", rest: " dictates solution priority — healthcare demands clinical verifiability above all." },
    ];

    concPoints.forEach((pt, i) => {
      const by = 0.95 + i * 1.0;
      s.addShape("rect", { x: 0.22, y: by, w: 0.06, h: 0.72, fill: { color: C.accent1 } });
      s.addText([
        { text: pt.bold, options: { bold: true, color: C.white } },
        { text: pt.rest, options: { color: C.muted } },
      ], { x: 0.4, y: by, w: 9.35, h: 0.72, fontSize: 11.5, valign: "middle", margin: 0 });
    });

    s.addText("[ Visual: Search 'data mining convergence funnel flat icon' – center-right background ]", {
      x: 0.22, y: 5.35, w: 9, h: 0.22, fontSize: 6.5, color: "2A4060", italic: true, margin: 0,
    });
  }

  // ═══════════════════════════════════════════════════════════
  // SLIDE 12 – THANK YOU
  // ═══════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: C.accent1 } });
    s.addShape("rect", { x: 0, y: 5.1, w: 10, h: 0.525, fill: { color: C.accent3 } });

    s.addText("Thank You", { x: 0.5, y: 1.2, w: 9, h: 1.1, fontSize: 58, bold: true, color: C.white, align: "center", margin: 0 });
    s.addText("Questions & Discussion", { x: 0.5, y: 2.38, w: 9, h: 0.45, fontSize: 18, color: C.accent1, align: "center", italic: true, margin: 0 });
    s.addShape("rect", { x: 3.5, y: 2.9, w: 3, h: 0.04, fill: { color: C.accent2 } });

    s.addText("Abdul Moeez  ·  01-134231-002  ·  BSCS 7B", { x: 0.5, y: 3.08, w: 9, h: 0.38, fontSize: 13, color: C.white, align: "center", margin: 0 });
    s.addText("Data Mining – Assignment 3  |  Bahria University, Islamabad", { x: 0.5, y: 3.5, w: 9, h: 0.32, fontSize: 11, color: C.muted, align: "center", margin: 0 });
    s.addText("Ma'am Erum Ashraf", { x: 0.5, y: 3.88, w: 9, h: 0.3, fontSize: 10, color: C.muted, align: "center", italic: true, margin: 0 });

    s.addText("IEEE  ·  Springer  ·  Wiley  ·  Elsevier  ·  ACM  ·  11 papers cited", {
      x: 0.5, y: 5.14, w: 9, h: 0.32, fontSize: 9, color: C.muted, align: "center", margin: 0,
    });

    s.addText("[ Visual: Search 'open book with data nodes flat icon' – center background watermark 8% opacity ]", {
      x: 0.25, y: 4.65, w: 9, h: 0.3, fontSize: 7, color: "2A4060", italic: true, align: "center", margin: 0,
    });
  }

  await pres.writeFile({ fileName: "DM3_Moeez_002_Revised.pptx" });
  console.log("Done ✓");
}

build().catch(console.error);
