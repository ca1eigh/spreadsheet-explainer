import { useState, useEffect, useRef, useCallback } from "react";
import * as d3 from "d3";

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

const COLUMNS = [
    { key: "id",          label: "ID",          type: "number" },
    { key: "name",        label: "Name",        type: "string" },
    { key: "dept",        label: "Department",  type: "string" },
    { key: "salary",      label: "Salary ($)",  type: "number" },
    { key: "hours",       label: "Hours/Wk",   type: "number" },
    { key: "performance", label: "Perf Score",  type: "number" },
];

// ── Anomaly detection engine ────────────────────────────────────────────────
function detectAnomalies(data) {
    const anomalies = [];
    const numericCols = COLUMNS.filter(c => c.type === "number").map(c => c.key);

    numericCols.forEach(col => {
        const vals = data.map(r => r[col]).filter(v => v !== null && v !== undefined && !isNaN(v));
        const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
        const std = Math.sqrt(vals.reduce((a, b) => a + (b - mean) ** 2, 0) / vals.length);

        data.forEach((row, i) => {
            const val = row[col];
            if (val === null || val === undefined || val === "") {
                anomalies.push({ row: i, col, severity: "missing", reason: `Missing value in ${col}`, value: val, mean, std });
            } else if (val < 0 && col !== "id") {
                anomalies.push({ row: i, col, severity: "invalid", reason: `Negative value (${val}) — logically impossible`, value: val, mean, std });
            } else {
                const z = Math.abs((val - mean) / std);
                if (z > 2.8) {
                    anomalies.push({ row: i, col, severity: z > 4 ? "critical" : "warning", reason: `Statistical outlier: ${val.toLocaleString()} is ${z.toFixed(1)}σ from mean (${mean.toFixed(1)})`, value: val, mean, std, zScore: z });
                }
            }
        });
    });

    // String checks
    data.forEach((row, i) => {
        if (!row.name || row.name.trim() === "") {
            anomalies.push({ row: i, col: "name", severity: "missing", reason: "Missing employee name", value: row.name });
        }
    });

    return anomalies;
}

// ── D3 Dependency Graph ─────────────────────────────────────────────────────
function DependencyGraph({ selected, allAnomalies, data }) {
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

// ── Main App ────────────────────────────────────────────────────────────────
export default function SpreadsheetExplainer() {
    const [data, setData] = useState(SAMPLE_DATA);
    const [anomalies, setAnomalies] = useState([]);
    const [selected, setSelected] = useState(null);
    const [editingCell, setEditingCell] = useState(null);
    const [editValue, setEditValue] = useState("");
    const [filter, setFilter] = useState("all");
    const [animatedRows, setAnimatedRows] = useState(new Set());

    useEffect(() => {
        setAnomalies(detectAnomalies(data));
    }, [data]);

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

    const commitEdit = (rowIdx, colKey) => {
        const col = COLUMNS.find(c => c.key === colKey);
        const parsed = col.type === "number" ? (editValue === "" ? null : parseFloat(editValue)) : editValue;
        const newData = data.map((row, i) => i === rowIdx ? { ...row, [colKey]: parsed } : row);
        setData(newData);
        setEditingCell(null);
        setAnimatedRows(prev => new Set([...prev, rowIdx]));
        setTimeout(() => setAnimatedRows(prev => { const s = new Set(prev); s.delete(rowIdx); return s; }), 600);
    };

    const filteredAnomalies = filter === "all" ? anomalies : anomalies.filter(a => a.severity === filter);
    const counts = { critical: anomalies.filter(a => a.severity === "critical").length, warning: anomalies.filter(a => a.severity === "warning").length, missing: anomalies.filter(a => a.severity === "missing").length, invalid: anomalies.filter(a => a.severity === "invalid").length };

    return (
        <div style={{ background: "#0d1117", minHeight: "100vh", fontFamily: "'JetBrains Mono', monospace", color: "#c9d1d9", display: "flex", flexDirection: "column" }}>
            {/* Header */}
            <div style={{ background: "#161b22", borderBottom: "1px solid #21262d", padding: "12px 24px", display: "flex", alignItems: "center", gap: "16px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                    <div style={{ width: "28px", height: "28px", background: "linear-gradient(135deg, #e53e3e, #d69e2e)", borderRadius: "6px", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "14px" }}>⚡</div>
                    <span style={{ fontSize: "15px", fontWeight: "700", color: "#f0f6fc", letterSpacing: "-0.3px" }}>SheetScan</span>
                    <span style={{ fontSize: "11px", color: "#4a5568", background: "#21262d", padding: "2px 8px", borderRadius: "10px" }}>Explainable Error Detection</span>
                </div>
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
                                {COLUMNS.map(col => (
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
                                        {COLUMNS.map(col => {
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
                                <DependencyGraph selected={selected} allAnomalies={anomalies} data={data} />
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
                                <DependencyGraph selected={null} allAnomalies={anomalies} data={data} />
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
                </div>
            </div>

            <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&display=swap');
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
        ::-webkit-scrollbar{width:6px;height:6px} ::-webkit-scrollbar-track{background:#0d1117} ::-webkit-scrollbar-thumb{background:#21262d;border-radius:3px}
      `}</style>
        </div>
    );
}