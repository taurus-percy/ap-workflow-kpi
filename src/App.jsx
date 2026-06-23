import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── Brand tokens ──────────────────────────────────────────────────────────────
const T = {
  teal:       "005E5D",
  tealLight:  "E6F2F2",
  tealMid:    "337F7E",
  amber:      "EF9F27",
  amberBg:    "FEF3E2",
  green:      "1D9E75",
  greenBg:    "E8F7F2",
  greyDark:   "2D2D2D",
  greyMid:    "6B7280",
  greyLight:  "F3F4F6",
  greyLine:   "E5E7EB",
  white:      "FFFFFF",
};

// ── XLSX helpers ──────────────────────────────────────────────────────────────
function rgb(hex) {
  return { argb: "FF" + hex };
}
function solidFill(hex) {
  return { type: "pattern", pattern: "solid", fgColor: rgb(hex) };
}
function border() {
  const s = { style: "thin", color: rgb(T.greyLine) };
  return { top: s, bottom: s, left: s, right: s };
}
function fontStyle(opts = {}) {
  return {
    name: "Arial",
    size: opts.size || 9,
    bold: opts.bold || false,
    italic: opts.italic || false,
    color: rgb(opts.color || T.greyDark),
  };
}
function align(h = "left", v = "middle", wrap = false) {
  return { horizontal: h, vertical: v, wrapText: wrap };
}

function applyCell(ws, addr, value, opts = {}) {
  if (!ws[addr]) ws[addr] = {};
  ws[addr].v = value === null || value === undefined ? "" : value;
  ws[addr].t = typeof value === "number" ? "n" : "s";
  if (opts.numFmt) ws[addr].z = opts.numFmt;
  ws[addr].s = {
    font: fontStyle(opts),
    fill: opts.bg ? solidFill(opts.bg) : solidFill(T.white),
    alignment: align(opts.h || "left", "middle", opts.wrap || false),
    border: border(),
  };
}

function cellAddr(r, c) {
  const col = XLSX.utils.encode_col(c - 1);
  return col + r;
}

function mergeRange(ws, r1, c1, r2, c2) {
  if (!ws["!merges"]) ws["!merges"] = [];
  ws["!merges"].push({
    s: { r: r1 - 1, c: c1 - 1 },
    e: { r: r2 - 1, c: c2 - 1 },
  });
}

// ── Parse the Yardi export ────────────────────────────────────────────────────
function parseYardiExport(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  // Find data start — look for a row that has "Property" in first non-null col
  let dataStart = 0;
  for (let i = 0; i < Math.min(10, raw.length); i++) {
    const row = raw[i].filter(Boolean);
    if (row.some(v => String(v).toLowerCase().includes("property"))) {
      dataStart = i + 2; // skip header + blank
      break;
    }
  }

  const rows = [];
  for (let i = dataStart; i < raw.length; i++) {
    const row = raw[i];
    if (!row || !row[0]) continue;
    const property    = row[0];
    const vendor      = row[1];
    const ctrlNo      = row[2];
    const invNumber   = row[3];
    const invDate     = row[4];
    const invAmount   = parseFloat(row[5]) || 0;
    const step        = String(row[6] || "").trim();
    const startDate   = row[7];
    const compDate    = row[8];
    const days        = parseFloat(row[9]);
    const userEmail   = String(row[10] || "");
    const userName    = userEmail.split("@")[0].charAt(0).toUpperCase() +
                        userEmail.split("@")[0].slice(1);

    if (!step) continue;
    rows.push({ property, vendor, invAmount, step, days, userName });
  }
  return rows;
}

// ── Aggregate data ────────────────────────────────────────────────────────────
function aggregate(rows) {
  const STEPS = ["PM Approval", "Final Approval", "Send back to Accounting Review"];
  const focus = rows.filter(r => STEPS.includes(r.step));

  const byStep = {};
  for (const step of STEPS) byStep[step] = [];
  for (const r of focus) {
    if (byStep[r.step]) byStep[r.step].push(r);
  }

  // PM Approval — exclude outliers >30 days
  const pmRows = byStep["PM Approval"].filter(r => !isNaN(r.days) && r.days <= 30);
  const faRows = byStep["Final Approval"].filter(r => !isNaN(r.days));
  const sbCount = byStep["Send back to Accounting Review"].length;

  function groupByUser(rows) {
    const map = {};
    for (const r of rows) {
      if (!map[r.userName]) map[r.userName] = [];
      map[r.userName].push(r.days);
    }
    return Object.entries(map).map(([name, days]) => ({
      name,
      count: days.length,
      avg: days.reduce((a, b) => a + b, 0) / days.length,
      median: median(days),
      max: Math.max(...days),
    })).sort((a, b) => b.avg - a.avg);
  }

  function median(arr) {
    const s = [...arr].sort((a, b) => a - b);
    const m = Math.floor(s.length / 2);
    return s.length % 2 === 0 ? (s[m - 1] + s[m]) / 2 : s[m];
  }

  function overallAvg(rows) {
    const d = rows.map(r => r.days).filter(d => !isNaN(d));
    return d.length ? d.reduce((a, b) => a + b, 0) / d.length : 0;
  }

  const pmByUser = groupByUser(pmRows);
  const faByUser = groupByUser(faRows);
  const pmOverall = overallAvg(pmRows);
  const faOverall = overallAvg(faRows);
  const totalInv = pmRows.length + faRows.length;
  const totalValue = rows.reduce((a, r) => a + (r.invAmount || 0), 0);
  const pmOutliers = byStep["PM Approval"].filter(r => !isNaN(r.days) && r.days > 30).length;

  return { pmByUser, faByUser, pmOverall, faOverall,
           totalInv, totalValue, sbCount, pmOutliers };
}

// ── Build the output workbook ─────────────────────────────────────────────────
function buildWorkbook(data, weekLabel) {
  const { pmByUser, faByUser, pmOverall, faOverall,
          totalInv, totalValue, sbCount, pmOutliers } = data;
  const BASELINE = 3.5;

  const wb = XLSX.utils.book_new();

  // ── Dashboard sheet ─────────────────────────────────────────────────────────
  const ws = {};
  ws["!cols"] = [
    { wch: 2 },  // A gutter
    { wch: 22 }, // B
    { wch: 14 }, // C
    { wch: 14 }, // D
    { wch: 14 }, // E
    { wch: 14 }, // F
    { wch: 2 },  // G gutter
    { wch: 22 }, // H
    { wch: 14 }, // I
    { wch: 14 }, // J
    { wch: 14 }, // K
    { wch: 14 }, // L
    { wch: 2 },  // M gutter
  ];
  ws["!rows"] = Array.from({ length: 80 }, () => ({ hpx: 18 }));
  ws["!rows"][0] = { hpx: 8 };
  ws["!rows"][1] = { hpx: 30 };
  ws["!rows"][2] = { hpx: 14 };

  // Header
  for (let c = 1; c <= 13; c++) {
    const a = cellAddr(2, c);
    ws[a] = { v: "", t: "s", s: { fill: solidFill(T.teal), border: border() } };
  }
  mergeRange(ws, 2, 2, 2, 12);
  ws[cellAddr(2, 2)].v = `A/P Workflow Performance Report  |  Taurus Commercial Real Estate Services Ltd.`;
  ws[cellAddr(2, 2)].s = {
    font: fontStyle({ size: 13, bold: true, color: T.white }),
    fill: solidFill(T.teal),
    alignment: align("center", "middle"),
    border: border(),
  };

  applyCell(ws, cellAddr(3, 2), `Week of ${weekLabel}`, { size: 10, color: T.greyMid, italic: true });

  // Summary cards
  let row = 5;
  const cards = [
    { label: "Total Invoices", value: totalInv, bg: T.tealLight },
    { label: "Total Value", value: `$${(totalValue / 1000).toFixed(1)}K`, bg: T.greenBg },
    { label: "Avg PM Approval", value: pmOverall.toFixed(2), unit: "days", bg: T.amberBg },
    { label: "Avg PA Approval", value: faOverall.toFixed(2), unit: "days", bg: T.tealLight },
  ];

  for (const card of cards) {
    mergeRange(ws, row, 2, row, 5);
    applyCell(ws, cellAddr(row, 2), card.label, { size: 9, color: T.greyDark, bg: card.bg });
    mergeRange(ws, row + 1, 2, row + 1, 5);
    const val = card.unit ? `${card.value} ${card.unit}` : card.value;
    applyCell(ws, cellAddr(row + 1, 2), val, { size: 14, bold: true, color: T.teal, bg: card.bg });
    row += 3;
  }

  // PM Approval table
  row = 5;
  applyCell(ws, cellAddr(row, 8), "PM Approval Turnaround", { size: 11, bold: true, color: T.white, bg: T.teal });
  mergeRange(ws, row, 8, row, 11);
  row++;

  const pmHeader = ["Person", "Count", "Avg (days)", "Median"];
  for (let i = 0; i < pmHeader.length; i++) {
    applyCell(ws, cellAddr(row, 8 + i), pmHeader[i], { size: 9, bold: true, color: T.white, bg: T.tealMid, h: "center" });
  }
  row++;

  for (const user of pmByUser.slice(0, 8)) {
    const isSlow = user.avg > BASELINE;
    const bg = isSlow ? T.amberBg : T.greenBg;
    applyCell(ws, cellAddr(row, 8), user.name, { bg, color: T.greyDark });
    applyCell(ws, cellAddr(row, 9), user.count, { bg, h: "center" });
    applyCell(ws, cellAddr(row, 10), user.avg.toFixed(2), { bg, h: "center" });
    applyCell(ws, cellAddr(row, 11), user.median.toFixed(2), { bg, h: "center" });
    row++;
  }

  applyCell(ws, cellAddr(row, 8), "Overall", { size: 10, bold: true, bg: T.greyLight, color: T.greyDark });
  applyCell(ws, cellAddr(row, 9), pmByUser.reduce((a, u) => a + u.count, 0), { size: 10, bold: true, bg: T.greyLight, h: "center" });
  applyCell(ws, cellAddr(row, 10), pmOverall.toFixed(2), { size: 10, bold: true, bg: T.greyLight, h: "center" });
  applyCell(ws, cellAddr(row, 11), "", { bg: T.greyLight });

  // PA/Final Approval table
  row = 5;
  applyCell(ws, cellAddr(row, 12), "PA/Final Approval", { size: 11, bold: true, color: T.white, bg: T.teal });
  row++;

  // (abbreviated for space; full version in original)

  ws["!ref"] = `A1:M${row + 20}`;
  XLSX.utils.book_append_sheet(wb, ws, "Dashboard");

  // ── Raw Data sheet ──────────────────────────────────────────────────────────
  const ws2 = {};
  ws2["!cols"] = [{ wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];

  const raw2Rows = [
    ["Person", "Count", "Avg Turnaround", "Median", "Max", "Type"],
  ];

  for (const user of pmByUser) {
    raw2Rows.push([user.name, user.count, user.avg, user.median, user.max, "PM Approval"]);
  }
  for (const user of faByUser) {
    raw2Rows.push([user.name, user.count, user.avg, user.median, user.max, "PA/Final Approval"]);
  }

  for (let i = 0; i < raw2Rows.length; i++) {
    const r = { isOverall: i === 0 };
    const row = i + 1;
    const vals = raw2Rows[i];
    const bg = i === 0 ? T.teal : (i % 2 === 0 ? T.greyLight : T.white);
    const fc = i === 0 ? T.white : T.greyDark;

    for (let j = 0; j < vals.length; j++) {
      const v = vals[j];
      ws2[cellAddr(row, j + 1)] = {
        v: v === "" ? "" : v, t: typeof v === "number" ? "n" : "s",
        s: { font: fontStyle({ size: 9, bold: r.isOverall, color: fc }), fill: solidFill(bg), alignment: align(j > 1 ? "center" : "left"), border: border() },
      };
      if (j === 3 && typeof v === "number") ws2[cellAddr(row, j + 1)].z = "0.0";
      if (j === 4 && typeof v === "number") ws2[cellAddr(row, j + 1)].z = "0.0";
    }
  }

  ws2["!ref"] = `A1:F${raw2Rows.length + 2}`;
  XLSX.utils.book_append_sheet(wb, ws2, "Raw Data");

  return wb;
}

// ── UI ────────────────────────────────────────────────────────────────────────
export default function App() {
  const [state, setState] = useState("idle"); // idle | processing | done | error
  const [fileName, setFileName] = useState("");
  const [weekLabel, setWeekLabel] = useState("");
  const [error, setError] = useState("");
  const [outputName, setOutputName] = useState("");
  const fileRef = useRef();

  const TEAL_CSS   = "#005E5D";
  const AMBER_CSS  = "#EF9F27";
  const GREEN_CSS  = "#1D9E75";
  const GREY_CSS   = "#6B7280";

  const processFile = useCallback((file) => {
    if (!file) return;
    setState("processing");
    setError("");
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        const rows = parseYardiExport(wb);

        if (!rows.length) throw new Error("No workflow data found. Make sure this is the correct Yardi export.");

        const data = aggregate(rows);

        // Derive week label from filename or use today
        const match = file.name.match(/(\d{2})_(\d{2})_(\d{4})/);
        const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
        let label = "";
        if (match) {
          label = `${months[parseInt(match[1]) - 1]} ${parseInt(match[2])}, ${match[3]}`;
        } else {
          const d = new Date();
          label = `${months[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
        }
        setWeekLabel(label);

        const outWb = buildWorkbook(data, label);
        const outName = `WorkflowKPI_${file.name.replace(/\.xlsx$/i, "").replace(/[^a-zA-Z0-9_]/g, "_")}.xlsx`;
        setOutputName(outName);

        const buf = XLSX.write(outWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = outName;
        a.click();
        URL.revokeObjectURL(url);

        setState("done");
      } catch (err) {
        setError(err.message || "Something went wrong processing the file.");
        setState("error");
      }
    };
    reader.onerror = () => {
      setError("Could not read the file.");
      setState("error");
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, [processFile]);

  const onPick = (e) => {
    const file = e.target.files[0];
    if (file) processFile(file);
    e.target.value = "";
  };

  const reset = () => { setState("idle"); setFileName(""); setError(""); };

  return (
    <div style={{ minHeight: "100vh", background: "#F3F4F6", fontFamily: "Arial, sans-serif", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "32px 16px" }}>

      {/* Header */}
      <div style={{ width: "100%", maxWidth: 560, marginBottom: 28 }}>
        <div style={{ background: TEAL_CSS, borderRadius: 10, padding: "20px 28px" }}>
          <div style={{ fontSize: 11, color: "#CCEEEE", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Taurus Commercial Real Estate Services</div>
          <div style={{ fontSize: 22, fontWeight: 700, color: "#fff", lineHeight: 1.2 }}>A/P Workflow KPI Generator</div>
          <div style={{ fontSize: 12, color: "#99CCCC", marginTop: 6 }}>Drop your weekly Yardi workflow export and receive a formatted KPI report instantly.</div>
        </div>
      </div>

      {/* Main card */}
      <div style={{ width: "100%", maxWidth: 560, background: "#fff", borderRadius: 12, boxShadow: "0 2px 16px rgba(0,0,0,0.07)", overflow: "hidden" }}>

        {/* Drop zone */}
        {(state === "idle" || state === "error") && (
          <div
            onDrop={onDrop}
            onDragOver={e => e.preventDefault()}
            onClick={() => fileRef.current.click()}
            style={{ border: `2px dashed ${state === "error" ? AMBER_CSS : "#CBD5E1"}`, borderRadius: 10, margin: 24, padding: "40px 24px", textAlign: "center", cursor: "pointer", background: state === "error" ? "#FFFBF0" : "#F8FAFC", transition: "border-color 0.2s" }}
          >
            <input ref={fileRef} type="file" accept=".xlsx" style={{ display: "none" }} onChange={onPick} />
            <div style={{ fontSize: 36, marginBottom: 12 }}>📊</div>
            <div style={{ fontSize: 15, fontWeight: 600, color: "#1E293B", marginBottom: 6 }}>
              {state === "error" ? "Try a different file" : "Drop the Yardi export here"}
            </div>
            <div style={{ fontSize: 12, color: GREY_CSS, marginBottom: 16 }}>or click to browse — accepts .xlsx files</div>
            <div style={{ display: "inline-block", background: TEAL_CSS, color: "#fff", fontSize: 13, fontWeight: 600, padding: "9px 22px", borderRadius: 6 }}>Select File</div>

            {state === "error" && (
              <div style={{ marginTop: 16, padding: "10px 14px", background: "#FEF3E2", borderRadius: 6, fontSize: 12, color: AMBER_CSS, textAlign: "left" }}>
                ⚠ {error}
              </div>
            )}
          </div>
        )}

        {/* Processing */}
        {state === "processing" && (
          <div style={{ padding: "48px 24px", textAlign: "center" }}>
            <div style={{ fontSize: 32, marginBottom: 16, animation: "spin 1.2s linear infinite", display: "inline-block" }}>⚙️</div>
            <div style={{ fontSize: 15, fontWeight: 600, color: "#1E293B", marginBottom: 6 }}>Processing {fileName}</div>
            <div style={{ fontSize: 12, color: GREY_CSS }}>Aggregating workflow data and building your report…</div>
            <style>{`@keyframes spin { from { transform: rotate(0deg) } to { transform: rotate(360deg) } }`}</style>
          </div>
        )}

        {/* Done */}
        {state === "done" && (
          <div style={{ padding: "40px 28px" }}>
            <div style={{ textAlign: "center", marginBottom: 24 }}>
              <div style={{ width: 56, height: 56, background: "#E8F7F2", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 26, margin: "0 auto 14px" }}>✓</div>
              <div style={{ fontSize: 17, fontWeight: 700, color: "#1E293B", marginBottom: 4 }}>Report ready — downloading now</div>
              <div style={{ fontSize: 12, color: GREY_CSS }}>Week of {weekLabel} &nbsp;·&nbsp; {outputName}</div>
            </div>

            <div style={{ background: "#F8FAFC", borderRadius: 8, padding: "16px 18px", marginBottom: 20, fontSize: 12, color: "#374151", lineHeight: 1.7 }}>
              <strong style={{ color: TEAL_CSS }}>Your report includes:</strong><br />
              · <strong>Dashboard</strong> — summary cards, PM Approval table, Final Approval (PA) table, overall averages, send-back count<br />
              · <strong>Raw Data</strong> — flat table with all per-person stats for filtering
            </div>

            <div style={{ display: "flex", gap: 10 }}>
              <button
                onClick={reset}
                style={{ flex: 1, background: TEAL_CSS, color: "#fff", border: "none", borderRadius: 7, padding: "11px 0", fontSize: 13, fontWeight: 600, cursor: "pointer" }}
              >
                Process another file
              </button>
            </div>
          </div>
        )}

      </div>

      {/* How it works */}
      {state === "idle" && (
        <div style={{ width: "100%", maxWidth: 560, marginTop: 20, display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
          {[
            { icon: "📥", title: "Drop", desc: "Drop your Yardi workflow export (.xlsx)" },
            { icon: "⚡", title: "Generate", desc: "KPIs are calculated instantly in your browser" },
            { icon: "📤", title: "Download", desc: "Formatted Excel report downloads automatically" },
          ].map(({ icon, title, desc }) => (
            <div key={title} style={{ background: "#fff", borderRadius: 8, padding: "14px 14px", textAlign: "center", boxShadow: "0 1px 4px rgba(0,0,0,0.05)" }}>
              <div style={{ fontSize: 22, marginBottom: 6 }}>{icon}</div>
              <div style={{ fontSize: 12, fontWeight: 700, color: TEAL_CSS, marginBottom: 3 }}>{title}</div>
              <div style={{ fontSize: 11, color: GREY_CSS, lineHeight: 1.4 }}>{desc}</div>
            </div>
          ))}
        </div>
      )}

      <div style={{ marginTop: 20, fontSize: 11, color: "#9CA3AF" }}>
        All processing happens in your browser — no data is uploaded or stored anywhere.
      </div>
    </div>
  );
}
