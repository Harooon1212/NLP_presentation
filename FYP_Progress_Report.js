const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Header, Footer, TabStopType, TabStopPosition,
  VerticalAlign, PageBreak
} = require('docx');
const fs = require('fs');

// ─── Colors ───
const DARK_BLUE = "1B3A5C";
const MID_BLUE  = "2E75B6";
const LIGHT_BLUE = "D5E8F4";
const ACCENT    = "C0392B";
const GOLD      = "B7860B";
const DARK_GRAY = "2C2C2C";
const MED_GRAY  = "555555";
const LIGHT_GRAY= "F2F2F2";
const TABLE_HDR = "1B3A5C";
const WHITE     = "FFFFFF";
const BORDER_CLR= "AAAAAA";

// ─── Helper: cell border ───
const cellBorder = (color = BORDER_CLR) => ({
  top:    { style: BorderStyle.SINGLE, size: 1, color },
  bottom: { style: BorderStyle.SINGLE, size: 1, color },
  left:   { style: BorderStyle.SINGLE, size: 1, color },
  right:  { style: BorderStyle.SINGLE, size: 1, color },
});

// ─── Helper: header cell ───
function hdrCell(text, w) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { fill: TABLE_HDR, type: ShadingType.CLEAR },
    borders: cellBorder(TABLE_HDR),
    margins: { top: 100, bottom: 100, left: 130, right: 130 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: WHITE, size: 20, font: "Arial" })]
    })]
  });
}

// ─── Helper: data cell ───
function dataCell(text, w, shade = WHITE, bold = false, center = false) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { fill: shade, type: ShadingType.CLEAR },
    borders: cellBorder(),
    margins: { top: 90, bottom: 90, left: 130, right: 130 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, bold, size: 19, font: "Arial", color: DARK_GRAY })]
    })]
  });
}

// ─── Helper: normal paragraph ───
function para(text, { bold=false, size=22, color=DARK_GRAY, spacing=160, italic=false, indent=0 } = {}) {
  return new Paragraph({
    indent: indent ? { left: indent } : undefined,
    spacing: { after: spacing },
    children: [new TextRun({ text, bold, size, color, italics: italic, font: "Arial" })]
  });
}

// ─── Helper: mixed-run paragraph ───
function mixedPara(runs, spacing = 160) {
  return new Paragraph({
    spacing: { after: spacing },
    children: runs.map(r => new TextRun({ font: "Arial", size: 22, color: DARK_GRAY, ...r }))
  });
}

// ─── Helper: bullet ───
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { after: 100 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: DARK_GRAY })]
  });
}

// ─── Helper: section heading (H1) ───
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: DARK_BLUE })]
  });
}

// ─── Helper: section heading (H2) ───
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: MID_BLUE })]
  });
}

// ─── Helper: H3 ───
function h3(text) {
  return new Paragraph({
    spacing: { before: 220, after: 120 },
    children: [new TextRun({ text, bold: true, size: 23, font: "Arial", color: ACCENT })]
  });
}

// ─── Helper: divider ───
function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 1 } },
    children: []
  });
}

// ─── Helper: label box row (diagram substitute using table) ───
function boxRow(cells) {
  // cells: [{text, shade, w}]
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: cells.map(c => c.w),
    rows: [
      new TableRow({
        children: cells.map(c => new TableCell({
          width: { size: c.w, type: WidthType.DXA },
          shading: { fill: c.shade, type: ShadingType.CLEAR },
          borders: cellBorder(c.border || MID_BLUE),
          margins: { top: 120, bottom: 120, left: 160, right: 160 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: c.text, bold: c.bold || false, size: c.size || 20,
              font: "Arial", color: c.color || DARK_GRAY })]
          })]
        }))
      })
    ]
  });
}

// ─── Helper: arrow-row paragraph ───
function arrowRow(label) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [new TextRun({ text: label, size: 20, font: "Arial", color: MED_GRAY, bold: true })]
  });
}

function spacer(pts = 120) {
  return new Paragraph({ spacing: { after: pts }, children: [] });
}

// ─── Caption ───
function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 260 },
    children: [new TextRun({ text, italics: true, size: 18, font: "Arial", color: MED_GRAY })]
  });
}

// ═══════════════════════════════════════════
// DIAGRAM 1 — DFS Data Access / Caching Flow
// ═══════════════════════════════════════════
function buildDiagram1() {
  const W = 9360;
  const col = Math.floor(W / 3);

  function flowRow(label, shade, border, bold = false) {
    return [
      new Table({
        width: { size: W, type: WidthType.DXA },
        columnWidths: [W],
        rows: [new TableRow({ children: [
          new TableCell({
            width: { size: W, type: WidthType.DXA },
            shading: { fill: shade, type: ShadingType.CLEAR },
            borders: cellBorder(border),
            margins: { top: 110, bottom: 110, left: 200, right: 200 },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: label, bold, size: 20, font: "Arial",
                color: shade === DARK_BLUE || shade === "1A5276" ? WHITE : DARK_GRAY })]
            })]
          })
        ]})]
      }),
      spacer(4)
    ];
  }

  return [
    h3("Diagram 1 — DFS Caching & Data Access Flow"),
    spacer(60),
    ...flowRow("CLIENT REQUEST: Access File Data", DARK_BLUE, DARK_BLUE, true),
    arrowRow("▼  Step 1"),
    ...flowRow("Check Client Cache  →  Hit? → Return Data to Application", LIGHT_BLUE, MID_BLUE),
    arrowRow("▼  Miss: Step 2"),
    ...flowRow("Check Client Local Disk (if any)  →  Hit? → Load to Client Cache", "D6EAF8", MID_BLUE),
    arrowRow("▼  Miss: Step 3"),
    ...flowRow("Send Request over Network to File Server", "FDEBD0", "E59866"),
    arrowRow("▼  Step 4"),
    ...flowRow("Check Server Cache  →  Hit? → Return Data", "FEF9E7", GOLD),
    arrowRow("▼  Miss: Step 5"),
    ...flowRow("Issue Disk Read on Server  →  Load Server Cache", "EAF4E8", "27AE60"),
    arrowRow("▼  Step 6"),
    ...flowRow("Return Data to Client → Load into Client Cache", "1A5276", "1A5276", true),
    spacer(8),
    caption("Figure 1: Step-by-step DFS data access with multi-level caching (Lecture, Slide 6)"),
  ];
}

// ═══════════════════════════════════════════
// DIAGRAM 2 — Write / Cache Consistency Policies
// ═══════════════════════════════════════════
function buildDiagram2() {
  const W = 9360;
  const c1 = 2200, c2 = 3580, c3 = 3580;

  const rows = [
    ["Policy", "Mechanism", "Trade-off"],
    ["Write-Through", "Immediate write to server on cache modify", "High reliability; high network traffic"],
    ["Delayed Write", "Write to server after a time delay", "Low traffic; risk of data loss on crash"],
    ["Write-on-Close", "Flush modified cache when file is closed", "Balanced; suitable for sequential sharing"],
    ["Server-Initiated", "Server notifies client caches on change", "Proactive; scales better"],
    ["Client-Initiated", "Client checks freshness before every read", "Accurate; per-access overhead"],
    ["Concurrent-Write", "Server purges caches of other clients", "Strong consistency; complex management"],
  ];

  return [
    h3("Diagram 2 — Write & Cache Consistency Policy Comparison Table"),
    spacer(60),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [c1, c2, c3],
      rows: rows.map((r, i) => new TableRow({
        children: r.map((cell, j) => {
          const isHdr = i === 0;
          const shade = isHdr ? TABLE_HDR : (i % 2 === 0 ? WHITE : LIGHT_GRAY);
          const color = isHdr ? WHITE : DARK_GRAY;
          const w = j === 0 ? c1 : j === 1 ? c2 : c3;
          return new TableCell({
            width: { size: w, type: WidthType.DXA },
            shading: { fill: shade, type: ShadingType.CLEAR },
            borders: cellBorder(isHdr ? TABLE_HDR : BORDER_CLR),
            margins: { top: 90, bottom: 90, left: 130, right: 130 },
            children: [new Paragraph({
              alignment: j === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
              children: [new TextRun({ text: cell, bold: isHdr || j === 0, size: 19,
                font: "Arial", color })]
            })]
          });
        })
      }))
    }),
    spacer(8),
    caption("Figure 2: Write and cache consistency policies in DFS (Lecture, Slides 20–21)"),
  ];
}

// ═══════════════════════════════════════════
// DIAGRAM 3 — NFS vs GFS Architecture (side-by-side)
// ═══════════════════════════════════════════
function buildDiagram3() {
  const W = 9360;
  const half = 4680;
  const inner = 4360;

  function sideBox(title, lines, shade, borderColor, titleColor) {
    const children = [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
        children: [new TextRun({ text: title, bold: true, size: 22, font: "Arial",
          color: titleColor || WHITE })]
      }),
      ...lines.map(l => new Paragraph({
        indent: { left: 100 },
        spacing: { after: 80 },
        children: [new TextRun({ text: `• ${l}`, size: 18, font: "Arial",
          color: shade === WHITE || shade === LIGHT_GRAY || shade === "EAF4FB" ? DARK_GRAY : WHITE })]
      }))
    ];
    return new TableCell({
      width: { size: half, type: WidthType.DXA },
      shading: { fill: shade, type: ShadingType.CLEAR },
      borders: cellBorder(borderColor),
      margins: { top: 140, bottom: 140, left: 160, right: 160 },
      children
    });
  }

  return [
    h3("Diagram 3 — NFS vs GFS Architecture Comparison"),
    spacer(60),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [half, half],
      rows: [
        new TableRow({ children: [
          sideBox("SUN NFS", [
            "Client–Server model via RPC",
            "VFS / vnode abstraction layer",
            "Stateless server design",
            "Per-client mount tables",
            "Small file optimized (8KB blocks)",
            "Iterative name resolution",
            "Timestamp-based cache freshness",
            "Delayed write + write-on-close",
          ], DARK_BLUE, DARK_BLUE, WHITE),
          sideBox("GOOGLE FILE SYSTEM (GFS)", [
            "Single Master + Chunk Servers",
            "64 MB fixed-size chunks",
            "3x replication across racks",
            "In-memory metadata at Master",
            "HeartBeat failure detection",
            "Lease-based mutation ordering",
            "Relaxed consistency model",
            "Atomic record append",
          ], MID_BLUE, MID_BLUE, WHITE),
        ]})
      ]
    }),
    spacer(8),
    caption("Figure 3: Architectural comparison of NFS and GFS (Lecture, Slides 30–65)"),
  ];
}

// ═══════════════════════════════════════════
// DIAGRAM 4 — GFS Fault Tolerance Mechanisms
// ═══════════════════════════════════════════
function buildDiagram4() {
  const W = 9360;
  const c1 = 2800, c2 = 3280, c3 = 3280;

  const rows = [
    ["Mechanism", "Description", "Benefit"],
    ["Chunk Replication", "Each 64MB chunk replicated on 3+ chunkservers across racks", "Survives individual disk/server failure"],
    ["Master Replication", "Operation log & checkpoints replicated on multiple remote machines", "No single point of failure for metadata"],
    ["HeartBeat Messages", "Periodic signals between Master and chunkservers", "Detects failed nodes quickly"],
    ["Operation Log", "Historical record of all metadata changes, flushed to disk before visible", "Enables crash recovery & replay"],
    ["Checkpointing", "Compact snapshots of namespace to limit log replay size", "Fast recovery startup time"],
    ["Checksumming", "Each chunkserver independently verifies data integrity", "Detects silent data corruption"],
    ["Stale Replica GC", "Chunk version numbers track stale replicas; garbage collected", "Prevents stale reads post-failure"],
  ];

  return [
    h3("Diagram 4 — GFS Fault Tolerance Mechanisms"),
    spacer(60),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [c1, c2, c3],
      rows: rows.map((r, i) => new TableRow({
        children: r.map((cell, j) => {
          const isHdr = i === 0;
          const shade = isHdr ? "1A5276" : (i % 2 === 0 ? WHITE : "EBF5FB");
          const color = isHdr ? WHITE : DARK_GRAY;
          const w = j === 0 ? c1 : j === 1 ? c2 : c3;
          return new TableCell({
            width: { size: w, type: WidthType.DXA },
            shading: { fill: shade, type: ShadingType.CLEAR },
            borders: cellBorder(isHdr ? "1A5276" : BORDER_CLR),
            margins: { top: 90, bottom: 90, left: 130, right: 130 },
            children: [new Paragraph({
              children: [new TextRun({ text: cell, bold: isHdr || j === 0, size: 19,
                font: "Arial", color })]
            })]
          });
        })
      }))
    }),
    spacer(8),
    caption("Figure 4: GFS fault tolerance mechanisms (Lecture, Slides 59, 63, 78–79)"),
  ];
}

// ═══════════════════════════════════════════
// BUILD DOCUMENT
// ═══════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "\u25E6",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
        ]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22, color: DARK_GRAY } }
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: DARK_BLUE },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: MID_BLUE },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            spacing: { after: 0 },
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 1 } },
            children: [
              new TextRun({ text: "Parallel and Distributed Computing  |  DC Lecture 04: Distributed File Systems", size: 18, font: "Arial", color: MED_GRAY }),
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 1 } },
            spacing: { before: 80 },
            children: [
              new TextRun({ text: "Page ", size: 18, font: "Arial", color: MED_GRAY }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: MED_GRAY }),
            ]
          })
        ]
      })
    },
    children: [

      // ────────────── TITLE PAGE ──────────────
      spacer(400),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [new TextRun({ text: "DISTRIBUTED FILE SYSTEMS", bold: true, size: 48, font: "Arial", color: DARK_BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: "Exam Answer Sheet  —  Questions 2 & 3", size: 26, font: "Arial", color: MED_GRAY, italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } },
        spacing: { after: 240 },
        children: []
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Course: Parallel and Distributed Computing", size: 22, font: "Arial", color: MED_GRAY })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Reference: DC-Lecture-04: Distributed File Systems", size: 22, font: "Arial", color: MED_GRAY })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Instructors: Mrs. Maryam Khalid  |  Mrs. Saima Jawad", size: 22, font: "Arial", color: MED_GRAY })]
      }),
      spacer(600),
      // Page break
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // QUESTION 2
      // ═══════════════════════════════════════
      new Paragraph({
        spacing: { after: 60 },
        shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
        children: []
      }),
      new Paragraph({
        spacing: { after: 220 },
        children: [new TextRun({ text: "QUESTION 2", bold: true, size: 36, font: "Arial", color: DARK_BLUE })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text: "Discuss the design choices in Distributed File Systems, focusing on:", bold: true, size: 24, font: "Arial", color: DARK_GRAY })]
      }),
      bullet("Caching strategies and consistency issues"),
      bullet("Replication for availability"),
      bullet("Write and cache consistency policies"),
      para("Explain how these choices affect scalability, fault tolerance, and system semantics.", { italic: true, size: 20, color: MED_GRAY }),
      divider(),

      // ── Introduction ──
      h1("Introduction to DFS Design Choices"),
      para("A Distributed File System (DFS) is a file system spread over multiple, autonomous computers that provides users with transparent access to files regardless of their physical location. According to the lecture, a DFS must achieve network transparency (hiding where a file is stored) and high availability (ensuring accessibility irrespective of physical location). Achieving both simultaneously demands careful design choices, particularly in caching, replication, and consistency management."),
      para("The key design axes are: Naming, Authentication & Access Control, Batched Operations, Caching, Concurrency Control, and Locking. This answer focuses deeply on caching, replication, and write/consistency policies — the three most consequential design dimensions."),

      // ── Section 1: Caching ──
      h1("1. Caching Strategies and Consistency Issues"),

      h2("1.1 Purpose and Role of Caching"),
      para("The lecture defines caching as bringing a copy of remote data to the client so that subsequent accesses are served locally, dramatically reducing response time and network traffic. In a DFS, files may reside on geographically distributed servers. Without caching, every read or write would traverse the network, creating unacceptable latency."),
      para("Client caches can reside in main memory or on disk. Main memory caches are faster and work for diskless workstations, but compete with virtual memory. Disk caches are slower but support larger files and allow the client to disconnect from the network when needed — an important feature for mobile clients."),
      para("An important variation is Hints: when cached data need not be completely accurate, the address of a server is stored as a hint. If the hint fails, the cache is purged and a name server is consulted to find the actual location. Unlike strict caching, hints are neither updated nor invalidated when the source changes — sacrificing strict accuracy for performance."),

      h2("1.2 Cache Location: Main Memory vs Disk"),
      ...buildDiagram1(),

      para("As shown in Figure 1, data access flows through multiple cache levels before a disk read is issued. The diagram mirrors the lecture's data access flow on Slide 6. At each step, a cache hit short-circuits the path, saving time. The client cache check comes first, followed by local disk (if available), then the server cache, and finally an actual disk read on the server."),

      h2("1.3 Cache Consistency"),
      para("Caching introduces a fundamental consistency challenge: what happens when the source data changes? The lecture identifies three consistency strategies:"),
      bullet("Server-initiated policy: The server's cache manager informs client cache managers whenever a file changes. Clients then retrieve fresh data. This proactively prevents stale reads and scales better (as noted in the lecture's scalability discussion)."),
      bullet("Client-initiated policy: The client checks data freshness before every delivery to the user. This involves comparing timestamps of the cached copy against the server. Sun NFS uses this approach — cached file attributes expire after 3 seconds and directory attributes after 30 seconds. The overhead is per-access."),
      bullet("Concurrent-write sharing policy: When multiple clients have a file open and at least one is writing, the server purges the cached copies of the file from all other clients to maintain consistency. This is the strongest guarantee but most expensive."),
      bullet("Sequential-write sharing policy: When a client opens a file recently closed after writing by another client, timestamps are compared. The server may force the previous writing client to flush its cache."),
      spacer(100),

      // ── Section 2: Replication ──
      h1("2. Replication for Availability"),

      h2("2.1 Why Replication?"),
      para("The lecture states that the objective of high availability is difficult to achieve because a DFS is vulnerable to network failures and server crashes. Replication — maintaining copies of files at multiple servers — is the primary solution. It ensures that even if one server fails, another copy remains accessible."),
      para("However, replication introduces new problems: maintaining consistency across replicas. If a write updates one replica but not others before a crash, inconsistency arises. The lecture acknowledges that 'inconsistency problems and their recovery may reduce the benefit of replication.'"),

      h2("2.2 Unit of Replication"),
      bullet("File-level replication: the most common approach. Each file may be replicated at different servers, requiring extra name resolution to locate replicas."),
      bullet("Group-of-files replication: files related by directory or access pattern are replicated together. This reduces name resolution overhead but wastes disk space if only a few files in the group are frequently accessed."),

      h2("2.3 Replica Management Schemes"),
      para("The lecture describes several replica management protocols:"),
      bullet("Two-phase commit: updates are sent to all replicas atomically. Either all replicas commit the change or none do, ensuring strong consistency."),
      bullet("Weighted votes: a quorum system where a read requires r votes and a write requires w votes from replicas. This allows flexible trade-offs between read and write performance."),
      bullet("Current Synchronization Site (CSS): one designated process controls all modifications. File open/close operations go through the CSS, ensuring total ordering of updates — but introducing a bottleneck."),
      spacer(100),

      // ── Section 3: Write & Cache Consistency ──
      h1("3. Write and Cache Consistency Policies"),

      h2("3.1 Writing Policy — When to Push Changes to Server"),
      para("The lecture identifies three writing policies that govern when a modified cache must be synchronized back to the server:"),
      ...buildDiagram2(),
      para("Write-through is the most reliable: every modification to the client cache is immediately written to the server. This guarantees that a client crash does not result in data loss. However, even small, frequent changes cause network writes, increasing traffic significantly."),
      para("Delayed writing defers the server update, reducing traffic for applications that make many small changes. The trade-off is reduced reliability — a client crash before the delayed write occurs results in lost updates."),
      para("Write-on-close is a balanced approach: modified data is flushed to the server when the file is closed. Sun NFS uses this policy for sequential-write sharing scenarios, where one client writes and another subsequently opens the same file."),

      h2("3.2 Impact on System Semantics"),
      para("The semantics of a DFS operation — what a read is guaranteed to return — depends directly on which consistency and writing policies are used. The lecture asks: does read() return the result of the most recent write()? In a single-machine file system, yes. In a DFS with delayed writing and client-initiated consistency checks, the answer is 'not necessarily immediately.'"),
      para("Unix semantics (one-copy semantics) guarantee that a read always returns the result of the latest write. Enforcing this in a caching DFS requires either write-through or immediate cache invalidation, both of which are expensive. Session semantics (as in NFS) relax this: consistency is enforced at file close/open boundaries rather than on every operation."),

      // ── Section 4: Impact on Scalability, FT, Semantics ──
      h1("4. Impact on Scalability, Fault Tolerance, and Semantics"),

      h2("4.1 Scalability"),
      para("The lecture directly addresses scalability: server-initiated cache invalidation scales better than client-initiated polling because the server only generates messages when data actually changes. For client scaling, the lecture proposes a hierarchical approach: after the first X clients are served directly by the server, new clients are directed to existing clients (forming a chain or hierarchy) who act as mini-file-servers. Cache misses and invalidations propagate up and down this hierarchy."),
      para("Stateless server design (as in NFS) also aids scalability — servers that don't track open files or file positions can handle more concurrent connections without maintaining per-client state."),

      h2("4.2 Fault Tolerance"),
      para("Replication directly improves fault tolerance by ensuring multiple copies of data exist. However, the replication protocol determines how much protection is actually provided. Two-phase commit is strongly fault-tolerant but slow. Weighted votes allow a degraded-mode read even when some replicas are offline."),
      para("The lecture notes that stateless servers in NFS recover from crashes simply by restarting — no need to restore transaction records or renegotiate file state with clients. Clients simply resend unanswered requests. GFS (discussed in Question 3) uses replication, operation logs, and checksumming for deep fault tolerance."),

      h2("4.3 System Semantics"),
      para("Design choices create a spectrum of semantic guarantees. Strong consistency (write-through + server-invalidation) approaches one-copy semantics but is expensive. Relaxed models (delayed write + hint-based caching) offer performance at the cost of temporary inconsistency. Applications must be designed knowing which semantics the DFS provides."),
      divider(),
      para("In summary, DFS design is an exercise in balancing competing forces: caching improves performance but requires consistency management; replication improves availability but requires replica synchronization; strong write policies improve reliability but hurt scalability. The right balance depends on the workload — interactive small-file access vs. large sequential data processing — a point that becomes concrete in the NFS vs. GFS comparison below."),
      spacer(200),

      // Page break before Q3
      new Paragraph({ children: [new PageBreak()] }),

      // ═══════════════════════════════════════
      // QUESTION 3
      // ═══════════════════════════════════════
      new Paragraph({
        spacing: { after: 220 },
        children: [new TextRun({ text: "QUESTION 3", bold: true, size: 36, font: "Arial", color: DARK_BLUE })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        children: [new TextRun({ text: "Compare Sun NFS and Google File System (GFS) as DFS case studies:", bold: true, size: 24, font: "Arial", color: DARK_GRAY })]
      }),
      bullet("Naming and metadata management"),
      bullet("Consistency and fault tolerance mechanisms"),
      bullet("Suitability for different types of workloads"),
      para("Conclude by stating which system is more appropriate for large-scale data-intensive applications and why.", { italic: true, size: 20, color: MED_GRAY }),
      divider(),

      // ── Intro Q3 ──
      h1("Introduction"),
      para("Sun NFS (Network File System) and Google File System (GFS) are two landmark distributed file system implementations that address the same fundamental challenge — transparent, reliable access to files across a network — from very different philosophical starting points. NFS was designed in the 1980s to provide general-purpose, OS-independent file sharing across heterogeneous systems. GFS was designed in the early 2000s specifically for Google's large-scale, data-intensive workloads. Comparing them illuminates the central design trade-offs in any DFS."),

      // ── Side-by-side diagram ──
      h1("1. Architectural Overview"),
      ...buildDiagram3(),

      // ── Naming & Metadata ──
      h1("2. Naming and Metadata Management"),

      h2("2.1 Sun NFS — Naming"),
      para("NFS uses a Virtual File System (VFS) interface with vnode (virtual node) abstractions. Each vnode is globally unique within NFS and carries a mount table pointer linking it to its parent file system and the remote system it is mounted over. This architecture allows NFS to be mapped onto different underlying file systems — Unix, DOS, or others — providing hardware and OS independence, which was NFS's primary design goal."),
      para("Naming in NFS is client-specific. Each client independently mounts remote directories into its own local namespace. Two clients can mount the same remote directory at different local paths, so the name /students/jack on Client 1 may map to /usr/students/jack on Client 2. The lecture describes this as 'different clients can see different name spaces'. Name resolution is iterative: to resolve /a/b/c/d, NFS looks up each component step-by-step, potentially contacting multiple servers, which makes it slow."),
      para("NFS also supports tilde context (~name) to provide user-level name space portability: ~john/t resolves differently depending on the client's mount configuration, but always to John's home directory — a useful device for systems where clients mount in different ways."),

      h2("2.2 GFS — Naming and Metadata"),
      para("GFS takes a radically different approach. All metadata is centralized in a single Master server that keeps the following in memory: the file and chunk namespace, the mapping from files to 64-bit chunk handles, and the locations of all chunk replicas. Having metadata in memory allows extremely fast lookups. The lecture notes that a 64MB chunk requires only 64 bytes of metadata, so even terabytes of data produce manageable metadata volume."),
      para("The namespace is represented as a flat lookup table mapping full pathnames (e.g., /foo/bar) to metadata — not a true directory hierarchy. Read/write locks are used over namespace regions to serialize concurrent operations. The Master's operation log records all metadata changes and is replicated on multiple remote machines, providing durability."),
      para("Unlike NFS's iterative resolution, GFS clients contact the Master once to obtain the chunk handle and replica locations, cache this information, and then communicate directly with chunkservers for data — minimizing Master involvement."),

      h2("2.3 Key Differences in Naming"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2200, 3580, 3580],
        rows: [
          new TableRow({ children: [hdrCell("Aspect", 2200), hdrCell("Sun NFS", 3580), hdrCell("GFS", 3580)] }),
          new TableRow({ children: [dataCell("Namespace", 2200, LIGHT_GRAY, true), dataCell("Per-client mount tables; heterogeneous", 3580), dataCell("Single global lookup table; flat pathnames", 3580)] }),
          new TableRow({ children: [dataCell("Name Resolution", 2200, LIGHT_GRAY, true), dataCell("Iterative; each component resolved separately (slow)", 3580), dataCell("One Master lookup → direct chunkserver contact (fast)", 3580)] }),
          new TableRow({ children: [dataCell("Metadata Location", 2200, LIGHT_GRAY, true), dataCell("Distributed; servers are stateless (no central metadata)", 3580), dataCell("Centralized in-memory at Master; replicated for durability", 3580)] }),
          new TableRow({ children: [dataCell("Consistency of Namespace", 2200, LIGHT_GRAY, true), dataCell("Eventual; clients see different paths", 3580), dataCell("Strong; namespace locking with read/write locks", 3580)] }),
        ]
      }),
      spacer(8),
      caption("Table 1: Naming and metadata management comparison"),

      // ── Consistency & FT ──
      h1("3. Consistency and Fault Tolerance Mechanisms"),

      h2("3.1 NFS Consistency"),
      para("NFS uses a client-initiated, timestamp-based consistency model. Client caches are validated by comparing file timestamps against the server. File attributes are considered stale after 3 seconds and directory attributes after 30 seconds. The lecture describes this as a form of hint-based caching: cached directory entries (vnode lookups) are also treated as hints — if a hint fails (the looked-up entry no longer exists), the cache is purged and resolution retried."),
      para("For writes, NFS uses a delayed writing policy with flush-on-close for sequential-write sharing scenarios. This means two clients writing to the same file may briefly see inconsistent views. NFS servers are explicitly designed to be stateless: they maintain no record of which clients have files open, what the file pointer position is, or what transactions are in progress. If a server crashes and restarts, clients simply resend their requests. Statelessness simplifies crash recovery but limits the server's ability to enforce strong consistency or provide useful information for concurrent file sharing."),

      h2("3.2 GFS Consistency — Relaxed Model"),
      para("GFS uses a relaxed consistency model, acknowledging that in a large distributed system running on commodity hardware, component failures are the norm rather than the exception. The model defines four states for a file region after a mutation:"),
      bullet("Consistent: all clients see the same data regardless of which replica they read from."),
      bullet("Defined: consistent and all clients see exactly what the mutation wrote in its entirety."),
      bullet("Consistent but Undefined: all replicas are identical, but the data may be a mix of fragments from concurrent writes — concurrent writes succeed but produce undefined interleaving."),
      bullet("Inconsistent: clients see different data at different times — only occurs on failure."),
      spacer(100),
      para("GFS uses leases to enforce mutation ordering. The Master grants a lease to one chunkserver, making it the Primary for that chunk. The Primary assigns serial numbers to all mutation requests, and all secondaries follow this order. This ensures that even concurrent writes from multiple clients result in identical replicas, even if the regions are undefined. Leases expire after 60 seconds but can be extended or revoked."),
      para("For atomic record appends — GFS's signature operation — the system guarantees at-least-once atomicity. The client specifies data; GFS chooses the offset; all replicas receive the append in the same order. If an append fails at any replica, the client retries, potentially producing duplicate records. Applications are designed to handle duplicates using checksums and unique identifiers."),

      h2("3.3 Fault Tolerance"),
      ...buildDiagram4(),

      para("As shown in Figure 4, GFS implements a layered fault tolerance strategy rooted in the lecture's observation that on commodity hardware, 'component failure is the norm rather than the exception.' Every critical component has a backup mechanism: chunk data is replicated across multiple chunkservers on different physical racks, the Master's metadata is replicated remotely via the operation log, and data corruption is caught at the chunkserver level through checksumming."),
      para("NFS's fault tolerance is simpler: stateless servers recover by restarting, and clients retry requests. There is no built-in chunk replication, no operation log, and no checksumming. Availability depends on external mechanisms such as RAID storage or OS-level mirroring. The lecture explicitly notes NFS's vulnerability to underlying network problems and server crashes."),

      // ── Workload Suitability ──
      h1("4. Suitability for Different Workloads"),

      h2("4.1 NFS Workloads"),
      para("NFS was designed for general-purpose file sharing across heterogeneous systems — a university lab where students and staff share files across Unix and Windows machines, or a corporate environment sharing documents over a LAN. It excels at:"),
      bullet("Small-file interactive workloads: editing source code, reading configuration files, sharing office documents."),
      bullet("Heterogeneous environments: NFS can be mapped to DOS, NFS, Unix, and other file systems via the VFS layer."),
      bullet("Environments requiring familiar Unix semantics: standard file operations (open, read, write, close) work transparently."),
      bullet("Campus or enterprise LANs: where network reliability is high and file sizes are manageable."),
      spacer(100),
      para("NFS's iterative name resolution, stateless servers, and per-client mounting are liabilities at large scale but are entirely acceptable for dozens to hundreds of clients sharing moderately sized files."),

      h2("4.2 GFS Workloads"),
      para("GFS was built for Google's specific data processing pipeline. The lecture states GFS assumptions explicitly: multi-GB files are common, workloads consist of large streaming reads, small random reads, and large sequential appends. Throughput is prioritized over latency. Multiple clients concurrently appending to one log file (a common MapReduce pattern) are directly supported."),
      bullet("Large-scale batch data processing (MapReduce, indexing pipelines)."),
      bullet("Log aggregation: many producers appending to a shared file atomically."),
      bullet("Sequential scan of terabyte datasets."),
      bullet("Environments with hundreds of chunkservers and thousands of disks."),
      spacer(100),
      para("GFS's 64MB chunk size is a liability for small files — it produces wasted space (though lazy allocation mitigates internal fragmentation) and potential hot spots if many clients access the same small file. GFS is not suitable as a general-purpose NFS replacement."),

      h2("4.3 Head-to-Head Workload Comparison"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2800, 3280, 3280],
        rows: [
          new TableRow({ children: [hdrCell("Workload Dimension", 2800), hdrCell("Sun NFS", 3280), hdrCell("GFS", 3280)] }),
          new TableRow({ children: [dataCell("File Size", 2800, LIGHT_GRAY, true), dataCell("Small to medium files", 3280), dataCell("Multi-GB files (optimized)", 3280)] }),
          new TableRow({ children: [dataCell("Access Pattern", 2800, LIGHT_GRAY, true), dataCell("Random reads & writes, interactive", 3280), dataCell("Large sequential reads, appends", 3280)] }),
          new TableRow({ children: [dataCell("Number of Clients", 2800, LIGHT_GRAY, true), dataCell("Tens to hundreds", 3280), dataCell("Thousands", 3280)] }),
          new TableRow({ children: [dataCell("Consistency Model", 2800, LIGHT_GRAY, true), dataCell("Timestamp freshness; near-Unix semantics", 3280), dataCell("Relaxed; defined/consistent/undefined/inconsistent", 3280)] }),
          new TableRow({ children: [dataCell("Fault Recovery", 2800, LIGHT_GRAY, true), dataCell("Server restart; client retry", 3280), dataCell("Replication, operation log, checksums, fast recovery", 3280)] }),
          new TableRow({ children: [dataCell("Heterogeneity", 2800, LIGHT_GRAY, true), dataCell("Excellent (VFS maps to many OS/FS)", 3280), dataCell("Homogeneous commodity Linux clusters", 3280)] }),
          new TableRow({ children: [dataCell("Scalability", 2800, LIGHT_GRAY, true), dataCell("Limited; per-client state difficult to scale", 3280), dataCell("Designed for 100s of servers, 1000s of disks", 3280)] }),
          new TableRow({ children: [dataCell("Metadata Management", 2800, LIGHT_GRAY, true), dataCell("Distributed; iterative resolution", 3280), dataCell("Centralized in-memory master; fast lookups", 3280)] }),
        ]
      }),
      spacer(8),
      caption("Table 2: NFS vs GFS workload suitability comparison"),

      // ── Conclusion ──
      h1("5. Conclusion — Which System for Large-Scale Data-Intensive Applications?"),
      new Paragraph({
        spacing: { after: 180 },
        shading: { fill: "EBF5FB", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: MID_BLUE, space: 8 } },
        indent: { left: 240 },
        children: [new TextRun({ text: "For large-scale, data-intensive applications, GFS is significantly more appropriate than NFS.", bold: true, size: 24, font: "Arial", color: DARK_BLUE })]
      }),
      para("The reasoning, grounded directly in lecture content, is as follows:"),
      mixedPara([
        { text: "1. Scale: ", bold: true },
        { text: "GFS was architecturally designed for hundreds of chunkservers and thousands of disks. Its in-memory metadata, 64MB chunks that reduce Master involvement, and pipelined TCP data transfer all optimize for terabytes to petabytes of data. NFS has no equivalent infrastructure." }
      ]),
      spacer(60),
      mixedPara([
        { text: "2. Fault Tolerance: ", bold: true },
        { text: "On commodity hardware clusters, failures are frequent. GFS treats failure as normal — every chunk is replicated 3x across racks, the Master is replicated via operation logs, and checksumming detects silent corruption. NFS relies on server restart and client retry with no native replication." }
      ]),
      spacer(60),
      mixedPara([
        { text: "3. Workload Match: ", bold: true },
        { text: "Large data-intensive applications read sequentially across huge files and write by appending, matching GFS's streaming read optimization and atomic record append exactly. NFS's random-access, small-file optimizations are mismatched." }
      ]),
      spacer(60),
      mixedPara([
        { text: "4. Throughput vs Latency: ", bold: true },
        { text: "The lecture states GFS prioritizes high sustained bandwidth over low latency. Data-intensive pipelines — processing logs, building indexes, training models — are throughput-bound, not latency-bound. GFS's design aligns perfectly." }
      ]),
      spacer(60),
      mixedPara([
        { text: "5. Consistency for Append Workloads: ", bold: true },
        { text: "GFS's relaxed consistency model, while unintuitive, is well-suited to append-heavy workloads where applications already use checksums and unique identifiers. The cost of strong consistency at scale would be prohibitive." }
      ]),
      spacer(120),
      para("That said, NFS remains the superior choice for general-purpose, heterogeneous, interactive file sharing environments — corporate networks, university labs, and OS-agnostic sharing. GFS's design is, as the lecture itself notes, 'only viable in a specific environment' with 'limited security' — it is not a universal replacement."),
      para("In conclusion, the contrast between NFS and GFS perfectly illustrates the core lesson of DFS design: there is no universally optimal system. The right design emerges from understanding the target workload, the failure model, and the acceptable consistency trade-offs — a theme that runs through every design choice discussed in this lecture."),
      divider(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
        children: [new TextRun({ text: "— End of Answer Sheet —", italics: true, size: 20, font: "Arial", color: MED_GRAY })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('DFS_Answers_Q2_Q3.docx',buffer);
  console.log('Done');
});



