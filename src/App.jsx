import { useState, useEffect, useRef, useCallback } from "react";
import * as d3 from "d3";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";

// ── Sample dataset with seeded errors ──────────────────────────────────────
const SAMPLE_DATA = [
    { id: 1, name: "Alice Chen",    dept: "Engineering", salary: 95000,  hours: 40, performance: 4.2 },
    { id: 2, name: "Bob Martinez",  dept: "Marketing",   salary: 72000,  hours: 38, performance: 3.8 },
    { id: 3, name: "Carol White",   dept: "Engineering", salary: 980000, hours: 40, performance: 4.5 }, // outlier salary
    { id: 4, name: "David Kim",     dept: "HR",          salary: 68000,  hours: 55, performance: 3.1 }, // outlier hours
    { id: 5, name: "Eva Torres",    dept: "Engineering", salary: 91000,  hours: 40, performance: 4.0 },
    { id: 6, name: "Frank Lee",     dept: "Marketing",   salary: null,   hours: 35, performance: 3.5 }, // missing
    { id: 7, name: "Grace Park",    dept: "Engineering", salary: 88000,  hours: 41, performance: 9.9 }, // outlier perf
    { id: 8, name: "Hiro Tanaka",   dept: "HR",          salary: 71000,  hours: 39, performance: 3.9 },
    { id: 9, name: "Iris Patel",    dept: "Engineering", salary: 94000,  hours: -5, performance: 4.1 }, // negative hours
    { id: 10, name: "Jake Wilson",  dept: "Marketing",   salary: 75000,  hours: 37, performance: 3.6 },
    { id: 11, name: "Kara Brown",   dept: "Engineering", salary: 89000,  hours: 42, performance: 4.3 },
    { id: 12, name: "Leo Nguyen",   dept: "HR",          salary: 66000,  hours: 40, performance: null }, // missing
    { id: 13, name: "",             dept: "Marketing",   salary: 78000,  hours: 36, performance: 3.7 }, // missing name
    { id: 14, name: "Mia Davis",    dept: "Engineering", salary: 92000,  hours: 40, performance: 4.4 },
    { id: 15, name: "Noah Clark",   dept: "HR",          salary: 69000,  hours: 38, performance: 3.3 },
];

const DEFAULT_COLUMNS = [
    { key: "id",          label: "ID",          type: "number" },
    { key: "name",        label: "Name",        type: "string" },
    { key: "dept",        label: "Department",  type: "string" },
    { key: "salary",      label: "Salary ($)",  type: "number" },
    { key: "hours",       label: "Hours/Wk",   type: "number" },
    { key: "performance", label: "Perf Score",  type: "number" },
];

const MAX_ROWS = 5000;

function toKey(label, idx, used) {
    const base = String(label ?? `column_${idx + 1}`)
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "_")
        .replace(/^_+|_+$/g, "") || `column_${idx + 1}`;
    let key = base;
    let n = 2;
    while (used.has(key)) {
        key = `${base}_${n}`;
        n += 1;
    }
    used.add(key);
    return key;
}

function parseNumeric(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    if (typeof value !== "string") return null;
    const cleaned = value.replace(/[$,%\s,]/g, "");
    if (cleaned === "") return null;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : null;
}

function cloneRows(rows) {
    return rows.map((row) => ({ ...row }));
}

function inferColumnsAndData(rows) {
    const headers = Array.from(
        rows.reduce((set, row) => {
            Object.keys(row || {}).forEach(k => set.add(String(k)));
            return set;
        }, new Set()),
    );

    const used = new Set();
    const mapped = headers.map((header, idx) => ({ header, key: toKey(header, idx, used) }));

    const numericByHeader = Object.fromEntries(
        headers.map((header) => {
            let numericCount = 0;
            let totalCount = 0;
            rows.forEach((row) => {
                const raw = row?.[header];
                if (raw === null || raw === undefined || String(raw).trim() === "") return;
                totalCount += 1;
                if (parseNumeric(raw) !== null) numericCount += 1;
            });
            return [header, totalCount > 0 && numericCount / totalCount >= 0.7];
        }),
    );

    const columns = mapped.map(({ header, key }) => ({
        key,
        label: header,
        type: numericByHeader[header] ? "number" : "string",
    }));

    const data = rows.slice(0, MAX_ROWS).map((row) => {
        const normalized = {};
        mapped.forEach(({ header, key }) => {
            const raw = row?.[header];
            if (raw === null || raw === undefined || String(raw).trim() === "") {
                normalized[key] = null;
                return;
            }
            if (numericByHeader[header]) {
                normalized[key] = parseNumeric(raw);
                return;
            }
            normalized[key] = String(raw).trim();
        });
        return normalized;
    });

    return { columns, data };
}

function toSourcePreview(rows, columns) {
    const headers = columns.map(c => c.label);
    const sourceRows = rows.map((row) => {
        const sourceRow = {};
        columns.forEach((col) => {
            sourceRow[col.label] = row[col.key] ?? null;
        });
        return sourceRow;
    });
    return { headers, sourceRows };
}

/** Build sheet from current grid state (labels as headers, same column order as the UI). */
function downloadEditedFile(data, columns, bookType = "xlsx") {
    const headerRow = columns.map((c) => c.label);
    const body = data.map((row) =>
        columns.map((c) => {
            const v = row[c.key];
            if (v === null || v === undefined || v === "") return "";
            return v;
        }),
    );
    const ws = XLSX.utils.aoa_to_sheet([headerRow, ...body]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    const stamp = new Date().toISOString().slice(0, 10);
    const ext = bookType === "csv" ? "csv" : "xlsx";
    const filename = `sheetscan-edited-${stamp}.${ext}`;
    XLSX.writeFile(wb, filename, bookType === "csv" ? { bookType: "csv" } : {});
}

function formatCellValue(value) {
    if (value === null || value === undefined || value === "") return "Empty";
    if (typeof value === "number") return Number.isFinite(value) ? value.toLocaleString() : String(value);
    return String(value);
}

function suggestionForAnomaly(anomaly) {
    if (anomaly.severity === "missing") return "Fill in a valid value or explicitly mark this field as intentionally blank.";
    if (anomaly.severity === "invalid") return "Correct to a logically valid value (for example, remove negative quantities where not allowed).";
    if (anomaly.severity === "critical") return "Validate against source records and business rules, then correct or document the exception.";
    if (anomaly.severity === "warning") return "Review against expected range for this field and confirm if the value is legitimate.";
    return "Review and confirm this cell with the data owner.";
}

function getRelatedAnomalies(anomaly, anomalies) {
    const sameColumn = anomalies.filter((a) => a.col === anomaly.col && a.row !== anomaly.row);
    const sameRow = anomalies.filter((a) => a.row === anomaly.row && a.col !== anomaly.col);
    return { sameColumn, sameRow };
}

function buildAuditReport({ anomalies, data, columns, counts }) {
    const now = new Date();
    const header = [
        "SheetScan Audit Report",
        `Generated: ${now.toLocaleString()}`,
        `Rows analyzed: ${data.length}`,
        `Columns analyzed: ${columns.length}`,
        `Total issues: ${anomalies.length}`,
        `Severity breakdown: critical ${counts.critical}, warning ${counts.warning}, missing ${counts.missing}, invalid ${counts.invalid}`,
        "",
    ];

    const body = anomalies.length
        ? anomalies.map((anomaly, idx) => {
            const { sameColumn, sameRow } = getRelatedAnomalies(anomaly, anomalies);
            const traceColumn = sameColumn.slice(0, 3).map((a) => `Row ${a.row + 1}`).join(", ") || "None";
            const traceRow = sameRow.slice(0, 3).map((a) => a.col).join(", ") || "None";
            return [
                `Issue ${idx + 1}`,
                `Severity: ${anomaly.severity.toUpperCase()}`,
                `Location: Row ${anomaly.row + 1}, Column "${anomaly.col}"`,
                `Observed value: ${formatCellValue(anomaly.value)}`,
                `Detected problem: ${anomaly.reason}`,
                `Traceback context: ${sameColumn.length} related issue(s) in same column (${traceColumn}); ${sameRow.length} related issue(s) in same row (${traceRow}).`,
                `Suggested action: ${suggestionForAnomaly(anomaly)}`,
                "",
            ].join("\n");
        }).join("\n")
        : "No anomalies were detected in this dataset.";

    return `${header.join("\n")}${body}`;
}

function buildAuditDocHtml(reportText) {
    const escaped = reportText
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<title>SheetScan Audit Report</title>
</head>
<body style="font-family: Arial, sans-serif; font-size: 12px; line-height: 1.4; color: #111;">
<h1 style="margin: 0 0 8px;">SheetScan Audit Report</h1>
<p style="margin: 0 0 14px; color: #444;">Shareable stakeholder summary of detected sheet anomalies.</p>
<pre style="white-space: pre-wrap; word-break: break-word; border: 1px solid #ddd; padding: 12px; border-radius: 6px;">${escaped}</pre>
</body>
</html>`;
}

function downloadAuditReport({ format, reportText }) {
    const stamp = new Date().toISOString().slice(0, 10);
    if (format === "pdf") {
        const pdf = new jsPDF({ unit: "pt", format: "a4" });
        const margin = 40;
        const pageHeight = pdf.internal.pageSize.getHeight();
        const maxY = pageHeight - margin;
        let y = margin;
        const lines = pdf.splitTextToSize(reportText, 515);
        pdf.setFont("courier", "normal");
        pdf.setFontSize(10);
        lines.forEach((line) => {
            if (y > maxY) {
                pdf.addPage();
                y = margin;
            }
            pdf.text(line, margin, y);
            y += 14;
        });
        pdf.save(`sheetscan-audit-report-${stamp}.pdf`);
        return;
    }

    const blob = new Blob([buildAuditDocHtml(reportText)], { type: "application/msword;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `sheetscan-audit-report-${stamp}.doc`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
}

// ── Anomaly detection engine ────────────────────────────────────────────────
function detectAnomalies(data, columns) {
    const anomalies = [];
    const numericCols = columns.filter(c => c.type === "number").map(c => c.key);

    numericCols.forEach(col => {
        const vals = data.map(r => r[col]).filter(v => v !== null && v !== undefined && !isNaN(v));
        const mean = vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
        const std = vals.length > 1 ? Math.sqrt(vals.reduce((a, b) => a + (b - mean) ** 2, 0) / vals.length) : 0;

        data.forEach((row, i) => {
            const val = row[col];
            if (val === null || val === undefined || val === "") {
                anomalies.push({ row: i, col, severity: "missing", reason: `Missing value in ${col}`, value: val, mean, std });
            } else if (val < 0 && col !== "id") {
                anomalies.push({ row: i, col, severity: "invalid", reason: `Negative value (${val}) — logically impossible`, value: val, mean, std });
            } else {
                const z = std > 0 ? Math.abs((val - mean) / std) : 0;
                if (Number.isFinite(z) && z > 2.8) {
                    anomalies.push({ row: i, col, severity: z > 4 ? "critical" : "warning", reason: `Statistical outlier: ${val.toLocaleString()} is ${z.toFixed(1)}σ from mean (${mean.toFixed(1)})`, value: val, mean, std, zScore: z });
                }
            }
        });
    });

    // String checks
    const stringCols = columns.filter(c => c.type === "string").map(c => c.key);
    data.forEach((row, i) => {
        stringCols.forEach((col) => {
            const value = row[col];
            if (value === null || value === undefined || String(value).trim() === "") {
                anomalies.push({ row: i, col, severity: "missing", reason: `Missing value in ${col}`, value });
            }
        });
    });

    return anomalies;
}

// ── D3 Dependency Graph ─────────────────────────────────────────────────────
function DependencyGraph({ selected, allAnomalies }) {
    const svgRef = useRef();

    useEffect(() => {
        if (!selected || !svgRef.current) return;
        const svg = d3.select(svgRef.current);
        svg.selectAll("*").remove();

        const W = 320, H = 280;
        const colAnomalies = allAnomalies.filter(a => a.col === selected.col && !(a.row === selected.row && a.col === selected.col));
        const rowAnomalies = allAnomalies.filter(a => a.row === selected.row && a.col !== selected.col);

        const nodes = [
            { id: "root", label: `Row ${selected.row + 1}\n${selected.col}`, type: "root", x: W / 2, y: H / 2 },
            ...colAnomalies.slice(0, 3).map((a, i) => ({ id: `col-${i}`, label: `Row ${a.row + 1}`, sublabel: a.severity, type: "col", x: 60 + i * 100, y: 60 })),
            ...rowAnomalies.slice(0, 3).map((a, i) => ({ id: `row-${i}`, label: a.col, sublabel: a.severity, type: "row", x: 60 + i * 100, y: H - 60 })),
        ];

        const links = [
            ...colAnomalies.slice(0, 3).map((_, i) => ({ source: "root", target: `col-${i}`, label: "same column" })),
            ...rowAnomalies.slice(0, 3).map((_, i) => ({ source: "root", target: `row-${i}`, label: "same row" })),
        ];

        const nodeMap = Object.fromEntries(nodes.map(n => [n.id, n]));

        const sim = d3.forceSimulation(nodes)
            .force("link", d3.forceLink(links).id(d => d.id).distance(90))
            .force("charge", d3.forceManyBody().strength(-180))
            .force("center", d3.forceCenter(W / 2, H / 2))
            .force("collision", d3.forceCollide(32));

        const defs = svg.append("defs");
        defs.append("marker").attr("id", "arrow").attr("viewBox", "0 -5 10 10")
            .attr("refX", 28).attr("markerWidth", 6).attr("markerHeight", 6).attr("orient", "auto")
            .append("path").attr("d", "M0,-5L10,0L0,5").attr("fill", "#4a5568");

        const linkG = svg.append("g");
        const linkLines = linkG.selectAll("line").data(links).join("line")
            .attr("stroke", "#2d3748").attr("stroke-width", 1.5)
            .attr("stroke-dasharray", "4,3").attr("marker-end", "url(#arrow)");

        const colorMap = { root: "#e53e3e", col: "#d69e2e", row: "#3182ce", missing: "#718096", invalid: "#e53e3e", warning: "#d69e2e", critical: "#e53e3e" };

        const nodeG = svg.append("g");
        const nodeCircles = nodeG.selectAll("circle").data(nodes).join("circle")
            .attr("r", d => d.type === "root" ? 22 : 16)
            .attr("fill", d => d.type === "root" ? colorMap[selected.severity] : colorMap[d.sublabel] || "#4a5568")
            .attr("stroke", "#1a202c").attr("stroke-width", 2)
            .style("filter", d => d.type === "root" ? "drop-shadow(0 0 8px rgba(229,62,62,0.6))" : "none");

        const labels = nodeG.selectAll("text.main").data(nodes).join("text")
            .attr("class", "main").attr("text-anchor", "middle").attr("dy", "0.35em")
            .attr("font-size", d => d.type === "root" ? "10px" : "9px")
            .attr("font-family", "'JetBrains Mono', monospace").attr("fill", "#f7fafc")
            .attr("font-weight", "bold").text(d => d.label.split("\n")[0]);

        const sublabels = nodeG.selectAll("text.sub").data(nodes).join("text")
            .attr("class", "sub").attr("text-anchor", "middle").attr("dy", "1.5em")
            .attr("font-size", "8px").attr("font-family", "'JetBrains Mono', monospace")
            .attr("fill", "#a0aec0").text(d => d.label.split("\n")[1] || "");

        sim.on("tick", () => {
            linkLines
                .attr("x1", d => nodeMap[d.source]?.x ?? d.source.x)
                .attr("y1", d => nodeMap[d.source]?.y ?? d.source.y)
                .attr("x2", d => nodeMap[d.target]?.x ?? d.target.x)
                .attr("y2", d => nodeMap[d.target]?.y ?? d.target.y);
            nodeCircles.attr("cx", d => d.x).attr("cy", d => d.y);
            labels.attr("x", d => d.x).attr("y", d => d.y);
            sublabels.attr("x", d => d.x).attr("y", d => d.y);
        });

        return () => sim.stop();
    }, [selected, allAnomalies]);

    if (!selected) return (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "280px", color: "#4a5568", fontFamily: "'JetBrains Mono', monospace", fontSize: "12px", gap: "8px" }}>
            <div style={{ fontSize: "32px" }}>⬡</div>
            <div>Click a flagged cell</div>
            <div style={{ color: "#2d3748" }}>to trace dependencies</div>
        </div>
    );

    return <svg ref={svgRef} width="320" height="280" style={{ display: "block" }} />;
}

function SourcePreview({ headers, rows, selectedRow }) {
    if (!headers.length || !rows.length) {
        return <div style={{ fontSize: "11px", color: "#4a5568", padding: "12px 0" }}>No source rows loaded yet.</div>;
    }

    const start = selectedRow === null ? 0 : Math.max(0, selectedRow - 2);
    const visibleRows = rows.slice(start, start + 8);

    return (
        <div style={{ background: "#0d1117", borderRadius: "8px", border: "1px solid #21262d", overflow: "auto", maxHeight: "230px" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "10px" }}>
                <thead>
                <tr style={{ background: "#11161d" }}>
                    <th style={{ padding: "6px 8px", color: "#6e7681", fontWeight: "600", borderBottom: "1px solid #21262d", textAlign: "left" }}>#</th>
                    {headers.map((header) => (
                        <th key={header} style={{ padding: "6px 8px", color: "#6e7681", fontWeight: "600", borderBottom: "1px solid #21262d", textAlign: "left", whiteSpace: "nowrap" }}>
                            {header}
                        </th>
                    ))}
                </tr>
                </thead>
                <tbody>
                {visibleRows.map((row, idx) => {
                    const originalRowIdx = start + idx;
                    const isSelected = selectedRow === originalRowIdx;
                    return (
                        <tr key={originalRowIdx} style={{ background: isSelected ? "rgba(229,62,62,0.12)" : "transparent" }}>
                            <td style={{ padding: "6px 8px", color: "#8b949e", borderBottom: "1px solid #21262d", whiteSpace: "nowrap" }}>{originalRowIdx + 1}</td>
                            {headers.map((header) => (
                                <td key={`${originalRowIdx}-${header}`} style={{ padding: "6px 8px", color: isSelected ? "#f0f6fc" : "#c9d1d9", borderBottom: "1px solid #21262d", whiteSpace: "nowrap" }}>
                                    {row[header] === null || row[header] === undefined || row[header] === "" ? "—" : String(row[header])}
                                </td>
                            ))}
                        </tr>
                    );
                })}
                </tbody>
            </table>
        </div>
    );
}

// ── Main App ────────────────────────────────────────────────────────────────
export default function SpreadsheetExplainer() {
    const [data, setData] = useState(SAMPLE_DATA);
    const [columns, setColumns] = useState(DEFAULT_COLUMNS);
    const [sourceHeaders, setSourceHeaders] = useState(DEFAULT_COLUMNS.map(c => c.label));
    const [sourceRows, setSourceRows] = useState(() => toSourcePreview(SAMPLE_DATA, DEFAULT_COLUMNS).sourceRows);
    const [anomalies, setAnomalies] = useState([]);
    const [selected, setSelected] = useState(null);
    const [editingCell, setEditingCell] = useState(null);
    const [editValue, setEditValue] = useState("");
    const [filter, setFilter] = useState("all");
    const [animatedRows, setAnimatedRows] = useState(new Set());
    const [uploadMessage, setUploadMessage] = useState("");
    const [historyOpen, setHistoryOpen] = useState(false);
    const [versionHistory, setVersionHistory] = useState(() => [
        {
            id: "sample-initial",
            label: "Initial sample dataset",
            timestamp: new Date().toISOString(),
            data: cloneRows(SAMPLE_DATA),
            columns: DEFAULT_COLUMNS.map((col) => ({ ...col })),
            change: null,
        },
    ]);
    const fileInputRef = useRef(null);
    const exportFormatRef = useRef(null);
    const reportFormatRef = useRef(null);

    useEffect(() => {
        setAnomalies(detectAnomalies(data, columns));
    }, [data, columns]);

    const getAnomaly = useCallback((rowIdx, colKey) =>
        anomalies.find(a => a.row === rowIdx && a.col === colKey), [anomalies]);

    const severityColor = { critical: "#e53e3e", warning: "#d69e2e", missing: "#718096", invalid: "#e53e3e" };
    const severityBg = { critical: "rgba(229,62,62,0.12)", warning: "rgba(214,158,46,0.12)", missing: "rgba(113,128,150,0.12)", invalid: "rgba(229,62,62,0.12)" };

    const handleCellClick = (anomaly) => {
        if (anomaly) setSelected(anomaly);
    };

    const startEdit = (rowIdx, colKey, val) => {
        setEditingCell(`${rowIdx}-${colKey}`);
        setEditValue(val === null || val === undefined ? "" : String(val));
    };

    const pushHistoryEntry = useCallback((entry) => {
        setVersionHistory((prev) => [entry, ...prev].slice(0, 30));
    }, []);

    const restoreVersion = (entry) => {
        const restoredData = cloneRows(entry.data);
        const restoredColumns = entry.columns.map((col) => ({ ...col }));
        const source = toSourcePreview(restoredData, restoredColumns);
        setColumns(restoredColumns);
        setData(restoredData);
        setSourceHeaders(source.headers);
        setSourceRows(source.sourceRows);
        setSelected(null);
        setFilter("all");
        setEditingCell(null);
        setUploadMessage(`Restored version from ${new Date(entry.timestamp).toLocaleString()}.`);
        setHistoryOpen(false);
    };

    const commitEdit = (rowIdx, colKey) => {
        const col = columns.find(c => c.key === colKey);
        const previousValue = data[rowIdx]?.[colKey] ?? null;
        const parsed = col.type === "number" ? (editValue === "" ? null : parseFloat(editValue)) : editValue;
        const newData = data.map((row, i) => i === rowIdx ? { ...row, [colKey]: parsed } : row);
        setData(newData);
        pushHistoryEntry({
            id: `edit-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
            label: `Edited row ${rowIdx + 1}, ${col.label}`,
            timestamp: new Date().toISOString(),
            data: cloneRows(newData),
            columns: columns.map((c) => ({ ...c })),
            change: {
                row: rowIdx + 1,
                column: col.label,
                before: previousValue,
                after: parsed,
            },
        });
        setEditingCell(null);
        setAnimatedRows(prev => new Set([...prev, rowIdx]));
        setTimeout(() => setAnimatedRows(prev => { const s = new Set(prev); s.delete(rowIdx); return s; }), 600);
    };

    const handleFileUpload = async (event) => {
        const file = event.target.files?.[0];
        if (!file) return;
        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
            if (!rows.length) {
                setUploadMessage("No data rows found. Make sure the first row has column headers.");
                return;
            }
            const normalized = inferColumnsAndData(rows);
            if (!normalized.columns.length) {
                setUploadMessage("Could not infer columns from this file.");
                return;
            }
            const rawHeaders = Array.from(
                rows.reduce((set, row) => {
                    Object.keys(row || {}).forEach(k => set.add(String(k)));
                    return set;
                }, new Set()),
            );
            const rawRows = rows.slice(0, MAX_ROWS).map((row) => {
                const nextRow = {};
                rawHeaders.forEach((header) => {
                    nextRow[header] = row?.[header] ?? null;
                });
                return nextRow;
            });
            setColumns(normalized.columns);
            setData(normalized.data);
            setSourceHeaders(rawHeaders);
            setSourceRows(rawRows);
            setSelected(null);
            setFilter("all");
            setEditingCell(null);
            setUploadMessage(`Loaded ${file.name} (${normalized.data.length} rows${rows.length > MAX_ROWS ? `, capped at ${MAX_ROWS}` : ""}).`);
            setVersionHistory([
                {
                    id: `upload-${Date.now()}`,
                    label: `Uploaded ${file.name}`,
                    timestamp: new Date().toISOString(),
                    data: cloneRows(normalized.data),
                    columns: normalized.columns.map((col) => ({ ...col })),
                    change: null,
                },
            ]);
        } catch {
            setUploadMessage("Could not parse file. Upload a valid CSV or Excel file.");
        } finally {
            event.target.value = "";
        }
    };

    const resetSample = () => {
        const source = toSourcePreview(SAMPLE_DATA, DEFAULT_COLUMNS);
        setColumns(DEFAULT_COLUMNS);
        setData(SAMPLE_DATA);
        setSourceHeaders(source.headers);
        setSourceRows(source.sourceRows);
        setSelected(null);
        setFilter("all");
        setEditingCell(null);
        setUploadMessage("Loaded sample dataset.");
        setVersionHistory([
            {
                id: `sample-${Date.now()}`,
                label: "Loaded sample dataset",
                timestamp: new Date().toISOString(),
                data: cloneRows(SAMPLE_DATA),
                columns: DEFAULT_COLUMNS.map((col) => ({ ...col })),
                change: null,
            },
        ]);
    };

    const counts = { critical: anomalies.filter(a => a.severity === "critical").length, warning: anomalies.filter(a => a.severity === "warning").length, missing: anomalies.filter(a => a.severity === "missing").length, invalid: anomalies.filter(a => a.severity === "invalid").length };
    const handleDownloadReport = () => {
        const format = reportFormatRef.current?.value === "doc" ? "doc" : "pdf";
        const reportText = buildAuditReport({ anomalies, data, columns, counts });
        downloadAuditReport({ format, reportText });
    };

    return (
        <div style={{ background: "#0d1117", minHeight: "100vh", fontFamily: "'JetBrains Mono', monospace", color: "#c9d1d9", display: "flex", flexDirection: "column" }}>
            {/* Header */}
            <div style={{ background: "#161b22", borderBottom: "1px solid #21262d", padding: "12px 24px", display: "flex", alignItems: "center", gap: "16px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                    <div style={{ width: "28px", height: "28px", background: "linear-gradient(135deg, #e53e3e, #d69e2e)", borderRadius: "6px", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "14px" }}>⚡</div>
                    <span style={{ fontSize: "15px", fontWeight: "700", color: "#f0f6fc", letterSpacing: "-0.3px" }}>SheetScan</span>
                    <span style={{ fontSize: "11px", color: "#4a5568", background: "#21262d", padding: "2px 8px", borderRadius: "10px" }}>Explainable Error Detection</span>
                </div>
                <input
                    ref={fileInputRef}
                    type="file"
                    accept=".csv,.xlsx,.xls"
                    onChange={handleFileUpload}
                    style={{ display: "none" }}
                />
                <button
                    onClick={() => fileInputRef.current?.click()}
                    style={{ background: "#238636", border: "1px solid #2ea043", color: "#f0f6fc", borderRadius: "6px", padding: "6px 10px", fontSize: "11px", cursor: "pointer" }}
                >
                    Upload CSV/Excel
                </button>
                <button
                    onClick={resetSample}
                    style={{ background: "#21262d", border: "1px solid #30363d", color: "#c9d1d9", borderRadius: "6px", padding: "6px 10px", fontSize: "11px", cursor: "pointer" }}
                >
                    Load Sample
                </button>
                <select
                    ref={exportFormatRef}
                    defaultValue="xlsx"
                    aria-label="Export file format"
                    style={{ background: "#21262d", border: "1px solid #30363d", color: "#c9d1d9", borderRadius: "6px", padding: "5px 8px", fontSize: "11px", cursor: "pointer", fontFamily: "inherit" }}
                >
                    <option value="xlsx">Excel (.xlsx)</option>
                    <option value="csv">CSV (.csv)</option>
                </select>
                <button
                    type="button"
                    title="Exports the current grid with your cell edits"
                    onClick={() => {
                        const fmt = exportFormatRef.current?.value === "csv" ? "csv" : "xlsx";
                        downloadEditedFile(data, columns, fmt);
                    }}
                    style={{ background: "#1f6feb", border: "1px solid #388bfd", color: "#f0f6fc", borderRadius: "6px", padding: "6px 10px", fontSize: "11px", cursor: "pointer" }}
                >
                    Download newly edited file
                </button>
                <button
                    type="button"
                    title="View and restore previous sheet versions"
                    onClick={() => setHistoryOpen(true)}
                    style={{ background: "#21262d", border: "1px solid #30363d", color: "#c9d1d9", borderRadius: "6px", padding: "6px 10px", fontSize: "11px", cursor: "pointer" }}
                >
                    Version History ({versionHistory.length})
                </button>
                <select
                    ref={reportFormatRef}
                    defaultValue="pdf"
                    aria-label="Audit report export format"
                    style={{ background: "#21262d", border: "1px solid #30363d", color: "#c9d1d9", borderRadius: "6px", padding: "5px 8px", fontSize: "11px", cursor: "pointer", fontFamily: "inherit" }}
                >
                    <option value="pdf">Report PDF</option>
                    <option value="doc">Report DOC</option>
                </select>
                <button
                    type="button"
                    title="Export a shareable audit report with issue tracebacks and suggested fixes"
                    onClick={handleDownloadReport}
                    style={{ background: "#8957e5", border: "1px solid #a371f7", color: "#f0f6fc", borderRadius: "6px", padding: "6px 10px", fontSize: "11px", cursor: "pointer" }}
                >
                    Download audit report
                </button>
                <div style={{ marginLeft: "auto", display: "flex", gap: "8px" }}>
                    {[["critical", "#e53e3e"], ["warning", "#d69e2e"], ["missing", "#718096"], ["invalid", "#fc8181"]].map(([sev, col]) => (
                        counts[sev] > 0 && (
                            <div key={sev} onClick={() => setFilter(filter === sev ? "all" : sev)}
                                 style={{ background: filter === sev ? col + "22" : "#21262d", border: `1px solid ${filter === sev ? col : "#30363d"}`, borderRadius: "6px", padding: "3px 10px", fontSize: "11px", cursor: "pointer", color: col, display: "flex", alignItems: "center", gap: "5px", transition: "all 0.15s" }}>
                                <span style={{ fontWeight: "700" }}>{counts[sev]}</span> {sev}
                            </div>
                        )
                    ))}
                </div>
            </div>

            <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
                {/* Main Grid */}
                <div style={{ flex: 1, overflow: "auto", padding: "16px" }}>
                    <div style={{ background: "#161b22", borderRadius: "10px", border: "1px solid #21262d", overflow: "hidden" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                            <thead>
                            <tr style={{ background: "#0d1117" }}>
                                {columns.map(col => (
                                    <th key={col.key} style={{ padding: "10px 14px", textAlign: "left", color: "#8b949e", fontWeight: "600", fontSize: "11px", letterSpacing: "0.5px", textTransform: "uppercase", borderBottom: "1px solid #21262d", whiteSpace: "nowrap" }}>
                                        {col.label}
                                    </th>
                                ))}
                            </tr>
                            </thead>
                            <tbody>
                            {data.map((row, rowIdx) => {
                                const rowHasAnomaly = anomalies.some(a => a.row === rowIdx);
                                const isAnimating = animatedRows.has(rowIdx);
                                return (
                                    <tr key={rowIdx} style={{ borderBottom: "1px solid #21262d", background: isAnimating ? "rgba(56,161,105,0.08)" : rowHasAnomaly ? "rgba(229,62,62,0.03)" : "transparent", transition: "background 0.4s" }}>
                                        {columns.map(col => {
                                            const anomaly = getAnomaly(rowIdx, col.key);
                                            const isSelected = selected?.row === rowIdx && selected?.col === col.key;
                                            const isEditing = editingCell === `${rowIdx}-${col.key}`;
                                            const val = row[col.key];

                                            return (
                                                <td key={col.key} onClick={() => handleCellClick(anomaly)}
                                                    onDoubleClick={() => startEdit(rowIdx, col.key, val)}
                                                    style={{
                                                        padding: "0", position: "relative", cursor: anomaly ? "pointer" : "default",
                                                        background: isSelected ? severityBg[anomaly?.severity] + " !important" : anomaly ? severityBg[anomaly.severity] : "transparent",
                                                        outline: isSelected ? `2px solid ${severityColor[anomaly?.severity]}` : anomaly ? `1px solid ${severityColor[anomaly.severity]}44` : "1px solid transparent",
                                                        outlineOffset: "-1px", transition: "all 0.15s"
                                                    }}>
                                                    {isEditing ? (
                                                        <input autoFocus value={editValue} onChange={e => setEditValue(e.target.value)}
                                                               onBlur={() => commitEdit(rowIdx, col.key)}
                                                               onKeyDown={e => { if (e.key === "Enter") commitEdit(rowIdx, col.key); if (e.key === "Escape") setEditingCell(null); }}
                                                               style={{ width: "100%", padding: "8px 14px", background: "#0d1117", color: "#f0f6fc", border: "none", outline: "none", fontFamily: "inherit", fontSize: "12px" }} />
                                                    ) : (
                                                        <div style={{ padding: "8px 14px", display: "flex", alignItems: "center", gap: "6px", minWidth: "80px" }}>
                                                            {anomaly && (
                                                                <span style={{ color: severityColor[anomaly.severity], fontSize: "10px", flexShrink: 0, animation: anomaly.severity === "critical" ? "pulse 1.5s infinite" : "none" }}>
                                    {anomaly.severity === "critical" ? "●" : anomaly.severity === "missing" ? "○" : "◐"}
                                  </span>
                                                            )}
                                                            <span style={{ color: anomaly ? severityColor[anomaly.severity] : val === null || val === undefined || val === "" ? "#4a5568" : "#e6edf3" }}>
                                  {val === null || val === undefined || val === "" ? "—" : typeof val === "number" ? val.toLocaleString() : val}
                                </span>
                                                        </div>
                                                    )}
                                                </td>
                                            );
                                        })}
                                    </tr>
                                );
                            })}
                            </tbody>
                        </table>
                    </div>
                    <div style={{ marginTop: "8px", color: "#4a5568", fontSize: "11px" }}>
                        💡 Click flagged cells to inspect · Double-click to edit · {anomalies.length} issues detected across {data.length} rows
                    </div>
                    {uploadMessage && (
                        <div style={{ marginTop: "6px", color: "#8b949e", fontSize: "11px" }}>
                            {uploadMessage}
                        </div>
                    )}
                </div>

                {/* Explanation Panel */}
                <div style={{ width: "360px", background: "#161b22", borderLeft: "1px solid #21262d", display: "flex", flexDirection: "column", flexShrink: 0 }}>
                    <div style={{ padding: "14px 16px", borderBottom: "1px solid #21262d" }}>
                        <div style={{ fontSize: "11px", color: "#8b949e", letterSpacing: "0.5px", textTransform: "uppercase", fontWeight: "600" }}>Explanation Inspector</div>
                    </div>

                    {selected ? (
                        <>
                            {/* Error badge */}
                            <div style={{ margin: "14px 16px 0", background: severityBg[selected.severity], border: `1px solid ${severityColor[selected.severity]}44`, borderRadius: "8px", padding: "12px" }}>
                                <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "6px" }}>
                                    <span style={{ background: severityColor[selected.severity], color: "#fff", fontSize: "10px", padding: "2px 8px", borderRadius: "4px", fontWeight: "700", textTransform: "uppercase" }}>{selected.severity}</span>
                                    <span style={{ color: "#8b949e", fontSize: "11px" }}>Row {selected.row + 1} · {selected.col}</span>
                                </div>
                                <div style={{ color: "#e6edf3", fontSize: "12px", lineHeight: "1.5" }}>{selected.reason}</div>
                                {selected.zScore && (
                                    <div style={{ marginTop: "8px", background: "#0d1117", borderRadius: "6px", padding: "8px" }}>
                                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: "11px", color: "#8b949e", marginBottom: "4px" }}>
                                            <span>Mean: <span style={{ color: "#c9d1d9" }}>{selected.mean?.toFixed(1)}</span></span>
                                            <span>Z-Score: <span style={{ color: severityColor[selected.severity] }}>{selected.zScore?.toFixed(2)}σ</span></span>
                                        </div>
                                        {/* Z-score bar */}
                                        <div style={{ height: "4px", background: "#21262d", borderRadius: "2px", overflow: "hidden" }}>
                                            <div style={{ height: "100%", width: `${Math.min(100, (selected.zScore / 6) * 100)}%`, background: `linear-gradient(90deg, #38a169, ${severityColor[selected.severity]})`, borderRadius: "2px", transition: "width 0.5s" }} />
                                        </div>
                                    </div>
                                )}
                            </div>

                            {/* Stats for column */}
                            <div style={{ margin: "10px 16px 0", display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px" }}>
                                {[["Affected Column", selected.col], ["Cell Value", selected.value === null ? "NULL" : String(selected.value)], ["Same Col Errors", anomalies.filter(a => a.col === selected.col).length], ["Same Row Errors", anomalies.filter(a => a.row === selected.row).length]].map(([label, val]) => (
                                    <div key={label} style={{ background: "#0d1117", borderRadius: "6px", padding: "8px 10px" }}>
                                        <div style={{ fontSize: "10px", color: "#4a5568", marginBottom: "2px" }}>{label}</div>
                                        <div style={{ fontSize: "12px", color: "#c9d1d9", fontWeight: "600" }}>{val}</div>
                                    </div>
                                ))}
                            </div>

                            {/* Graph */}
                            <div style={{ margin: "10px 16px 0", background: "#0d1117", borderRadius: "8px", overflow: "hidden" }}>
                                <div style={{ padding: "8px 12px", fontSize: "10px", color: "#4a5568", borderBottom: "1px solid #21262d", textTransform: "uppercase", letterSpacing: "0.5px" }}>Dependency Graph</div>
                                <DependencyGraph selected={selected} allAnomalies={anomalies} />
                            </div>

                            {/* Fix suggestion */}
                            <div style={{ margin: "10px 16px", background: "rgba(56,161,105,0.08)", border: "1px solid rgba(56,161,105,0.2)", borderRadius: "8px", padding: "10px 12px", fontSize: "11px", color: "#68d391" }}>
                                <span style={{ fontWeight: "700" }}>💡 Suggestion: </span>
                                {selected.severity === "missing" ? "Fill in the missing value or mark as intentionally empty." : selected.severity === "invalid" ? "Replace with a valid non-negative number." : `Check if ${selected.value?.toLocaleString()} is a data entry error; expected ~${selected.mean?.toFixed(0)}.`}
                                <div style={{ marginTop: "6px", color: "#4a5568", fontSize: "10px" }}>Double-click the cell to edit inline ↗</div>
                            </div>
                        </>
                    ) : (
                        <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
                            <div style={{ flex: 1 }}>
                                <DependencyGraph selected={null} allAnomalies={anomalies} />
                            </div>
                            <div style={{ padding: "16px", borderTop: "1px solid #21262d" }}>
                                <div style={{ fontSize: "11px", color: "#4a5568", marginBottom: "10px", textTransform: "uppercase", letterSpacing: "0.5px" }}>All Issues</div>
                                {anomalies.slice(0, 6).map((a, i) => (
                                    <div key={i} onClick={() => setSelected(a)} style={{ display: "flex", alignItems: "center", gap: "8px", padding: "7px 8px", borderRadius: "6px", cursor: "pointer", marginBottom: "4px", background: "#0d1117", border: "1px solid #21262d", transition: "border-color 0.15s" }}
                                         onMouseEnter={e => e.currentTarget.style.borderColor = severityColor[a.severity] + "66"}
                                         onMouseLeave={e => e.currentTarget.style.borderColor = "#21262d"}>
                                        <span style={{ color: severityColor[a.severity], fontSize: "8px" }}>●</span>
                                        <span style={{ color: "#8b949e", fontSize: "11px" }}>Row {a.row + 1}</span>
                                        <span style={{ color: "#4a5568", fontSize: "11px" }}>·</span>
                                        <span style={{ color: "#c9d1d9", fontSize: "11px", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.reason}</span>
                                    </div>
                                ))}
                                {anomalies.length > 6 && <div style={{ fontSize: "11px", color: "#4a5568", textAlign: "center", paddingTop: "6px" }}>+{anomalies.length - 6} more — click any flagged cell</div>}
                            </div>
                        </div>
                    )}
                    <div style={{ marginTop: "auto", borderTop: "1px solid #21262d", padding: "12px 16px 14px" }}>
                        <div style={{ fontSize: "10px", color: "#8b949e", marginBottom: "8px", textTransform: "uppercase", letterSpacing: "0.5px" }}>
                            Source File Preview
                        </div>
                        <SourcePreview headers={sourceHeaders} rows={sourceRows} selectedRow={selected?.row ?? null} />
                    </div>
                </div>
            </div>

            {historyOpen && (
                <div
                    onClick={() => setHistoryOpen(false)}
                    style={{ position: "fixed", inset: 0, background: "rgba(1,4,9,0.72)", display: "flex", justifyContent: "center", alignItems: "center", zIndex: 20 }}
                >
                    <div
                        onClick={(e) => e.stopPropagation()}
                        style={{ width: "min(720px, 92vw)", maxHeight: "80vh", overflow: "hidden", background: "#161b22", border: "1px solid #30363d", borderRadius: "10px", display: "flex", flexDirection: "column" }}
                    >
                        <div style={{ padding: "12px 14px", borderBottom: "1px solid #21262d", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                            <div>
                                <div style={{ fontSize: "12px", color: "#f0f6fc", fontWeight: 700 }}>Version History</div>
                                <div style={{ marginTop: "3px", fontSize: "11px", color: "#8b949e" }}>Snapshots are captured after each cell edit.</div>
                            </div>
                            <button
                                type="button"
                                onClick={() => setHistoryOpen(false)}
                                style={{ background: "#21262d", border: "1px solid #30363d", color: "#c9d1d9", borderRadius: "6px", padding: "4px 8px", fontSize: "11px", cursor: "pointer" }}
                            >
                                Close
                            </button>
                        </div>
                        <div style={{ overflow: "auto", padding: "10px 12px 12px" }}>
                            {versionHistory.map((entry, idx) => (
                                <div
                                    key={entry.id}
                                    style={{ background: "#0d1117", border: "1px solid #21262d", borderRadius: "8px", padding: "10px 12px", marginBottom: "8px", display: "flex", alignItems: "center", gap: "10px" }}
                                >
                                    <div style={{ minWidth: "76px", color: "#4a5568", fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.4px" }}>
                                        v{versionHistory.length - idx}
                                    </div>
                                    <div style={{ flex: 1, minWidth: 0 }}>
                                        <div style={{ color: "#c9d1d9", fontSize: "12px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{entry.label}</div>
                                        <div style={{ marginTop: "2px", color: "#8b949e", fontSize: "11px" }}>
                                            {new Date(entry.timestamp).toLocaleString()}
                                            {entry.change && ` • Row ${entry.change.row} • ${entry.change.column} • ${entry.change.before ?? "—"} -> ${entry.change.after ?? "—"}`}
                                        </div>
                                    </div>
                                    <button
                                        type="button"
                                        onClick={() => restoreVersion(entry)}
                                        style={{ background: "#1f6feb", border: "1px solid #388bfd", color: "#f0f6fc", borderRadius: "6px", padding: "5px 9px", fontSize: "11px", cursor: "pointer", flexShrink: 0 }}
                                    >
                                        Restore
                                    </button>
                                </div>
                            ))}
                        </div>
                    </div>
                </div>
            )}

            <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&display=swap');
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
        ::-webkit-scrollbar{width:6px;height:6px} ::-webkit-scrollbar-track{background:#0d1117} ::-webkit-scrollbar-thumb{background:#21262d;border-radius:3px}
      `}</style>
        </div>
    );
}