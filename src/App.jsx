import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { supabase } from './supabaseClient';
import * as Recharts from 'recharts';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';

const { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, Legend, LineChart, Line, ComposedChart, Cell, PieChart, Pie, LabelList } = Recharts;


        
        

        // Configuration Arrays
        const MONTHS = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ];

        const SHIFTS = ["Shift A", "Shift B", "Shift C"];

        // Formatting Helpers
        function fmtINR(n, opts = {}) {
            if (n === null || n === undefined || isNaN(n)) return "—";
            if (opts.compact) {
                if (Math.abs(n) >= 1e7) return `₹${(n / 1e7).toFixed(2)} Cr`;
                if (Math.abs(n) >= 1e5) return `₹${(n / 1e5).toFixed(2)} L`;
                if (Math.abs(n) >= 1e3) return `₹${(n / 1e3).toFixed(1)} K`;
                return `₹${Math.round(n)}`;
            }
            return new Intl.NumberFormat("en-IN", {
                style: "currency",
                currency: "INR",
                maximumFractionDigits: 0,
            }).format(n);
        }

        function fmtNum(n, opts = 0) {
            if (n === null || n === undefined || isNaN(n)) return "—";
            const num = Number(n);
            let decimals = 0;
            let compact = false;
            if (typeof opts === "number") {
                decimals = opts;
            } else if (typeof opts === "object" && opts !== null) {
                decimals = typeof opts.decimals === "number" ? opts.decimals : 0;
                compact = Boolean(opts.compact);
            }

            if (compact) {
                if (Math.abs(num) >= 1e7) return (num / 1e7).toFixed(2) + " Cr";
                if (Math.abs(num) >= 1e5) return (num / 1e5).toFixed(2) + " L";
                if (Math.abs(num) >= 1e3) return (num / 1e3).toFixed(1) + " k";
                return new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(num);
            }

            const validDecimals = Math.max(0, Math.min(20, Math.floor(Number(decimals) || 0)));
            return new Intl.NumberFormat("en-IN", {
                maximumFractionDigits: validDecimals
            }).format(num);
        }

        // Full-precision rupee value (2 decimals) — used in the sheet-style
        // entries table so numbers match the source Excel exactly.
        function fmtMoney(n, decimals = 2) {
            if (n === null || n === undefined || isNaN(n)) return "—";
            return "₹" + new Intl.NumberFormat("en-IN", {
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals
            }).format(n);
        }

        function getMonthEnd(monthStr) {
            if (!monthStr) return "";
            const parts = String(monthStr).split("-").map(Number);
            if (parts.length < 2 || isNaN(parts[0]) || isNaN(parts[1])) return `${monthStr}-31`;
            const year = parts[0];
            const month = parts[1];
            const lastDay = new Date(year, month, 0).getDate();
            return `${year}-${String(month).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
        }

        // Active Theme Colors
        const COLORS = {
            blue: "#0284c7",
            green: "#16a34a",
            orange: "#ea580c",
            red: "#dc2626",
            amber: "#d97706",
            purple: "#8b5cf6",
            gray: "#64748b",
            lightBlue: "#e0f2fe",
            lightGreen: "#dcfce7",
        };

        const PLANT_COLORS = {
            NGM: "#16a34a",  // Nashik (Green)
            PGEL: "#0284c7", // Pune PGEL (Blue)
            PGTL: "#ea580c"  // Pune PGTL (Orange)
        };

        // Available UI Themes
        const THEME_OPTIONS = [
            { id: "light", label: "Daylight", swatch: "linear-gradient(135deg, #ffffff 50%, #0284c7 50%)" },
            { id: "dark", label: "Midnight", swatch: "linear-gradient(135deg, #0b1220 50%, #38bdf8 50%)" },
            { id: "ocean", label: "Ocean Teal", swatch: "linear-gradient(135deg, #eaf6fb 50%, #0e7490 50%)" },
            { id: "emerald", label: "Emerald", swatch: "linear-gradient(135deg, #eafaf1 50%, #059669 50%)" },
        ];

        // ----------------------------------------------------
        // COMPONENTS
        // ----------------------------------------------------

        // KPI Metric Card — simple flat style
        function KpiCard({ label, value, sub, trend, icon, tone = "blue", trendLabel, compact }) {
            const toneStyles = {
                blue: "bg-sky-50 text-sky-600",
                green: "bg-emerald-50 text-emerald-600",
                emerald: "bg-green-50 text-green-600",
                orange: "bg-orange-50 text-orange-500",
                amber: "bg-amber-50 text-amber-600",
                red: "bg-rose-50 text-rose-600",
                pink: "bg-pink-50 text-pink-500",
                purple: "bg-violet-50 text-violet-600",
                indigo: "bg-indigo-50 text-indigo-600",
                teal: "bg-teal-50 text-teal-600",
                gray: "bg-slate-100 text-slate-500"
            };

            const iconBg = toneStyles[tone] || toneStyles.blue;
            const hasTrend = trend !== null && trend !== undefined;

            return (
                <div className={`bg-white ${compact ? 'p-2 gap-2 min-w-0 w-full' : 'p-2.5 gap-2.5 min-w-[150px]'} rounded-xl border border-slate-200 shadow-sm flex items-center shrink-0 cursor-default select-none`}>
                    <span className={`${compact ? 'h-8 w-8 rounded-lg' : 'h-9 w-9 rounded-lg'} flex items-center justify-center shrink-0 ${iconBg}`}>
                        <span className="material-symbols-outlined" style={{ fontSize: compact ? '17px' : '19px' }}>{icon}</span>
                    </span>
                    <div className="min-w-0 flex-1">
                        <h4 className="font-black text-slate-900 tracking-tight leading-none" style={{ fontSize: compact ? '15.5px' : '17.5px' }}>{value}</h4>
                        <p className="font-extrabold text-slate-800 leading-tight mt-0.5 truncate" style={{ fontSize: compact ? '11.5px' : '12.5px' }}>{label}</p>
                        {sub && <p className="font-bold text-slate-600 mt-0.5 truncate" style={{ fontSize: compact ? '9.5px' : '10.5px' }}>{sub}</p>}
                        {hasTrend && (
                            <div className="mt-0.5 flex items-center gap-1" style={{ fontSize: '10px', fontWeight: 700 }}>
                                <span className={`material-symbols-outlined ${trend > 0 ? "text-red-500 rotate-180" : "text-emerald-500"}`} style={{ fontSize: '12px' }}>
                                    arrow_downward
                                </span>
                                <span className={trend > 0 ? "text-red-500" : "text-emerald-600"}>
                                    {trend > 0 ? "+" : ""}{trend.toFixed(1)}%
                                </span>
                                <span className="text-slate-500" style={{ fontWeight: 600 }}>{trendLabel || "vs period"}</span>
                            </div>
                        )}
                    </div>
                </div>
            );
        }

        /** Production KPI — click to slide Production up and reveal ODU / IDU (same card size) */
        function ProductionKpiCard({ production, odu, idu, compact }) {
            const [open, setOpen] = useState(false);
            const pad = compact ? "p-2" : "p-2.5";
            const toggle = () => setOpen((v) => !v);
            return (
                <div
                    role="button"
                    tabIndex={0}
                    onClick={toggle}
                    onKeyDown={(e) => {
                        if (e.key === "Enter" || e.key === " ") {
                            e.preventDefault();
                            toggle();
                        }
                    }}
                    title={open ? "Click to show Production" : "Click to show ODU / IDU"}
                    className={`bg-white ${compact ? 'min-w-0 w-full' : 'min-w-[150px]'} ${pad} rounded-xl border border-slate-200 shadow-sm text-left cursor-pointer select-none overflow-hidden relative h-full box-border outline-none focus:outline-none`}
                >
                    <div
                        className={`absolute inset-0 ${pad} flex items-center transition-transform duration-300 ease-out`}
                        style={{ transform: open ? "translateY(-100%)" : "translateY(0)" }}
                    >
                        <div className={`flex items-center ${compact ? 'gap-2' : 'gap-2.5'} w-full`}>
                            <span className={`${compact ? 'h-8 w-8' : 'h-9 w-9'} rounded-lg flex items-center justify-center shrink-0 bg-emerald-50 text-emerald-600`}>
                                <span className="material-symbols-outlined" style={{ fontSize: compact ? '17px' : '19px' }}>factory</span>
                            </span>
                            <div className="min-w-0 flex-1">
                                <h4 className="font-black text-slate-900 tracking-tight leading-none" style={{ fontSize: compact ? '15.5px' : '17.5px' }}>{fmtNum(production)}</h4>
                                <p className="font-extrabold text-slate-800 leading-tight mt-0.5" style={{ fontSize: compact ? '11.5px' : '12.5px' }}>Production Out</p>
                            </div>
                        </div>
                    </div>
                    <div
                        className={`absolute inset-0 ${pad} flex items-center transition-transform duration-300 ease-out bg-white`}
                        style={{ transform: open ? "translateY(0)" : "translateY(100%)" }}
                    >
                        <div className="grid grid-cols-2 gap-1 w-full">
                            <div className="rounded-md bg-sky-50 border border-sky-100 px-1 py-0.5 text-center">
                                <p className="font-extrabold text-sky-600 uppercase leading-none tracking-wide" style={{ fontSize: '8.5px' }}>ODU</p>
                                <p className="font-black text-slate-900 tabular-nums leading-tight mt-0.5" style={{ fontSize: compact ? '12px' : '13px' }}>{fmtNum(odu)}</p>
                            </div>
                            <div className="rounded-md bg-teal-50 border border-teal-100 px-1 py-0.5 text-center">
                                <p className="font-extrabold text-teal-600 uppercase leading-none tracking-wide" style={{ fontSize: '8.5px' }}>IDU</p>
                                <p className="font-black text-slate-900 tabular-nums leading-tight mt-0.5" style={{ fontSize: compact ? '12px' : '13px' }}>{fmtNum(idu)}</p>
                            </div>
                        </div>
                    </div>
                </div>
            );
        }

        // Section header with custom icon
        function SectionTitle({ icon, title, subtitle }) {
            return (
                <div className="flex items-center gap-2.5 pb-2.5 border-b border-slate-200/50">
                    <span className="material-symbols-outlined text-[#0284c7] text-[20px]">{icon}</span>
                    <div>
                        <h3 className="text-sm font-bold text-slate-800 tracking-tight uppercase">{title}</h3>
                        {subtitle && <p className="text-[10px] text-slate-400">{subtitle}</p>}
                    </div>
                </div>
            );
        }

        // Small read-only key/value cell used inside the "View Entry Details" popup
        function DetailField({ label, value, highlight = false }) {
            const display = (value === undefined || value === null || value === "") ? "—" : value;
            return (
                <div>
                    <p className="text-[9px] font-semibold text-slate-400 uppercase tracking-wide mb-0.5">{label}</p>
                    <p className={`text-xs font-semibold ${highlight ? "text-[#0284c7]" : "text-slate-800"}`}>{display}</p>
                </div>
            );
        }

        // ----------------------------------------------------
        // REPORT EXPORTER & PRINT COMPONENT
        // ----------------------------------------------------
        function PrintableReportTable({ title, headers, data, totalRow }) {
            return (
                <div className="bg-white dark:bg-[#121a29] rounded-xl border border-slate-200/90 dark:border-[#26334a] p-3 shadow-sm print-card">
                    <div className="flex items-center gap-1.5 mb-2 border-b border-slate-100 dark:border-slate-800 pb-1.5">
                        <span className="material-symbols-outlined text-[#0284c7] text-[16px]">assessment</span>
                        <h3 className="text-[12px] font-extrabold text-slate-800 dark:text-slate-200 tracking-tight uppercase">{title}</h3>
                    </div>
                    <div className="overflow-x-auto rounded-lg border border-slate-100 dark:border-slate-800">
                        <table className="w-full border-collapse text-left text-xs">
                            <thead>
                                <tr className="border-b border-slate-200 dark:border-slate-800 text-slate-500 dark:text-slate-400 font-bold uppercase tracking-wider text-[9px] bg-slate-50/60 dark:bg-slate-900/40">
                                    {headers.map((h, i) => (
                                        <th key={i} className="py-2 px-2.5 font-bold tracking-wider">{h}</th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100 dark:divide-slate-800 text-slate-700 dark:text-slate-300">
                                {data.length === 0 ? (
                                    <tr>
                                        <td colSpan={headers.length} className="py-6 text-center text-slate-400 font-semibold bg-slate-50/20 dark:bg-slate-900/10">No records found for active filters</td>
                                    </tr>
                                ) : (
                                    data.map((row, rIdx) => (
                                        <tr key={rIdx} className="hover:bg-slate-50/30 dark:hover:bg-slate-900/20 transition">
                                            {row.map((cell, cIdx) => (
                                                <td key={cIdx} className="py-1.5 px-2.5 text-xs font-medium">{cell}</td>
                                            ))}
                                        </tr>
                                    ))
                                )}
                                {totalRow && data.length > 0 && (
                                    <tr className="bg-slate-50/80 dark:bg-slate-900/40 font-bold border-t-2 border-slate-200 dark:border-slate-800">
                                        {totalRow.map((cell, idx) => (
                                            <td key={idx} className="py-2 px-2.5 text-slate-900 dark:text-white font-extrabold text-xs">{cell}</td>
                                        ))}
                                    </tr>
                                )}
                            </tbody>
                        </table>
                    </div>
                </div>
            );
        }

        // ----------------------------------------------------
        // CALCULATION ENGINE FOR REDESIGNED DATA ENTRY
        // ----------------------------------------------------
        const DEFAULT_MF = 40;
        const TARIFF_DEFAULTS = { electricity: 10.893945, solar: 10.893945, diesel: 90.62, water: 45, lpg: 85 };

        const CalculationEngine = {
            calculateMSEBUnits: (diff, factor = DEFAULT_MF) => diff * (Number(factor) || DEFAULT_MF),
            calculateElectricityCost: (units, rate) => units * rate,
            calculateSolarCost: (units, rate) => units * rate,
            calculateDieselCost: (qty, rate) => qty * rate,
            calculateTotalCost: (elect, solar, diesel) => elect + solar + diesel,
            calculateProductionSets: (odu, idu) => {
                if (idu > odu) return Math.round((idu - odu) / 3 + odu);
                if (odu > idu) return Math.round((odu - idu) * 1.5 + idu);
                return odu;
            },
            calculateCostPerSet: (totalCost, prodSets) => {
                if (!prodSets || prodSets === 0) return 0;
                return Math.round(totalCost / prodSets);
            }
        };

        function getGridProviderLabel(location) {
            return String(location || "").trim().toUpperCase() === "BHIWADI" ? "JVVNL" : "MSEB";
        }

        function getLocationDefaultMF(location) {
            return String(location || "").trim().toUpperCase() === "BHIWADI" ? 30 : DEFAULT_MF;
        }

        function buildPlantReportColumns(gridLabel, mf) {
            const formulaRow = [
                "A", "B", "C",
                "C(Today)-C(Prev)=D",
                `D × ${mf} = E`,
                "F",
                "E + F = G",
                "H (Monthly)",
                "H × E = I",
                "H × F = J",
                "K (Monthly)",
                "L",
                "K × L = M",
                "I + J + M = N",
                "O", "P",
                "(IDU>ODU):(I−O)/3+O else (O−I)×1.5+I",
                "N ÷ Q = R",
            ];
            const headers = [
                "Date", "PLANT", "Daily Reading kWh UNITS", "UNITS Difference From Yesterday",
                `${gridLabel} UNITS x${mf}`, "SOLAR UNITS", "TOTAL UNITS", `${gridLabel} ₹ / UNIT`,
                `Consumtion ₹ ${gridLabel}`, "Consumtion ₹ SOLAR", "₹ / Liter DIESEL", "CONSUMED DIESEL (L)",
                "CONSUMED ₹ DIESEL", "Consumtion ₹ TOTAL", "ODU", "IDU", "PROD IN SET", "PROD/SET COST (₹)",
            ];
            return { headers, formulaRow };
        }

        function parseAllowedScope(raw) {
            const s = String(raw || "all").trim().toLowerCase();
            if (!s || s === "all") return "all";
            return s.split(",").map((x) => x.trim()).filter(Boolean);
        }

        function resolvePlantMeta(plantKey, plantList = []) {
            const p = (plantList || []).find((x) =>
                [x?.plant_code, x?.plant_display_name, x?.plant_name]
                    .filter(Boolean)
                    .some((v) => String(v).trim().toLowerCase() === String(plantKey || "").trim().toLowerCase())
            );
            return {
                code: p?.plant_code || plantKey || "",
                name: p?.plant_display_name || p?.plant_name || plantKey || "",
                location: p?.location || "",
            };
        }

        function userCanAccessMasterRow(user, row, tableName, plantList = []) {
            if (!user) return false;
            if (user.role === "IT_ADMIN") return true;
            if (tableName !== "tariff_rates" && tableName !== "multiply_factors") return false;

            const userLocStr = String(user.allowed_locations || "all").trim().toLowerCase();
            const userPlantStr = String(user.allowed_plants || "all").trim().toLowerCase();

            const rowLoc = String(row?.location || "").trim().toLowerCase();
            const rowPlant = String(row?.plant_code || "").trim().toLowerCase();

            // 1. Check Location Match
            let locMatch = false;
            if (!rowLoc || userLocStr === "all" || userLocStr === "") {
                locMatch = true;
            } else {
                locMatch = userLocStr.includes(rowLoc) || rowLoc.includes(userLocStr);
            }

            // 2. Check Plant Match
            let plantMatch = false;
            if (!rowPlant || userPlantStr === "all" || userPlantStr === "") {
                plantMatch = true;
            } else {
                const allowedPlantTokens = userPlantStr.split(",").map(x => x.trim().toLowerCase()).filter(Boolean);
                if (allowedPlantTokens.includes("all")) {
                    plantMatch = true;
                } else {
                    const meta = resolvePlantMeta(rowPlant, plantList);
                    const metaCode = String(meta.code || "").trim().toLowerCase();
                    const metaName = String(meta.name || "").trim().toLowerCase();

                    plantMatch = allowedPlantTokens.some(token =>
                        token === rowPlant ||
                        token === metaCode ||
                        token === metaName
                    );
                }
            }

            return locMatch && plantMatch;
        }

        function resolveMultiplyFactor(factors, plantCode, location, dateStr) {
            const date = dateStr || new Date().toISOString().split("T")[0];
            const plant = String(plantCode || "").trim().toLowerCase();
            const loc = String(location || "").trim().toUpperCase();
            const active = (factors || []).filter((f) =>
                f.status === "Active" && String(f.effective_date || "") <= date
            );

            const plantMatch = active
                .filter((f) => f.plant_code && String(f.plant_code).trim().toLowerCase() === plant)
                .sort((a, b) => String(b.effective_date).localeCompare(String(a.effective_date)));
            if (plantMatch.length) return Number(plantMatch[0].factor) || DEFAULT_MF;

            const locMatch = active
                .filter((f) => !f.plant_code && String(f.location || "").trim().toUpperCase() === loc)
                .sort((a, b) => String(b.effective_date).localeCompare(String(a.effective_date)));
            if (locMatch.length) return Number(locMatch[0].factor) || getLocationDefaultMF(loc);

            return getLocationDefaultMF(loc);
        }

        function resolveTariff(tariffRows, type, plantCode, location, dateStr) {
            const date = dateStr || new Date().toISOString().split("T")[0];
            const plant = String(plantCode || "").trim().toLowerCase();
            const loc = String(location || "").trim().toUpperCase();

            const candidates = (tariffRows || []).filter((t) =>
                t.type === type && t.status === "Active" && String(t.effective_date || "") <= date
            );

            const score = (t) => {
                const tPlant = String(t.plant_code || "").trim().toLowerCase();
                const tLoc = String(t.location || "").trim().toUpperCase();
                if (tPlant && tPlant === plant) return 3;
                if (tLoc && tLoc === loc && !tPlant) return 2;
                if (!tLoc && !tPlant) return 1;
                return 0;
            };

            const best = candidates
                .map((t) => ({ t, s: score(t) }))
                .filter((x) => x.s > 0)
                .sort((a, b) => b.s - a.s || String(b.t.effective_date).localeCompare(String(a.t.effective_date)));

            if (best.length) return Number(best[0].t.rate);
            return TARIFF_DEFAULTS[type] || 0;
        }

        function getUserDefaultMasterScope(user, plantRows) {
            const rows = plantRows || [];
            const allowedL = parseAllowedScope(user?.allowed_locations);
            const allowedP = parseAllowedScope(user?.allowed_plants);

            let location = rows[0]?.location || "";
            if (allowedL !== "all" && allowedL.length) {
                location = allowedL[0].toUpperCase();
            }

            let scopedPlants = rows.filter((p) =>
                String(p.location || "").trim().toUpperCase() === String(location || "").trim().toUpperCase()
            );
            if (allowedP !== "all") {
                scopedPlants = scopedPlants.filter((p) =>
                    allowedP.includes(String(p.plant_code || "").trim().toLowerCase())
                );
            }
            const plant_code = scopedPlants[0]?.plant_code || "";

            return { location, plant_code };
        }

        // ----------------------------------------------------
        // DEFAULT DASHBOARD DATE RANGE (current month start -> today)
        // ----------------------------------------------------
        function toISODate(d) {
            const pad = (n) => String(n).padStart(2, "0");
            return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
        }

        // Default: All Locations · All Plants · full year (so login pe cards empty na rahein)
        function getDefaultDateFilters() {
            const now = new Date();
            const startOfYear = new Date(now.getFullYear(), 0, 1);
            return {
                startDate: toISODate(startOfYear),
                endDate: toISODate(now),
                plant: "all",
                department: "all",
                location: "all"
            };
        }

        // Match entry.plant against plant_code OR display/name (DB mixes both)
        function plantIdentitySet(plantRows) {
            const set = new Set();
            (plantRows || []).forEach((p) => {
                [p.plant_code, p.plant_display_name, p.plant_name].forEach((x) => {
                    if (x) set.add(String(x).trim().toLowerCase());
                });
            });
            return set;
        }

        function entryPlantKey(plantValue) {
            return String(plantValue || "").trim().toLowerCase();
        }

        // ----------------------------------------------------
        // MAIN APPLICATION COMPONENT
        // ----------------------------------------------------
        // ---------------------------------------------------- 
        // LOGO DATA URL CONSTANT 
        // ---------------------------------------------------- 
        const PG_LOGO_BASE_64 = "data:image/png;base64,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";

        // ---------------------------------------------------- 
        // MAIN APPLICATION COMPONENT 
        // Full-screen intro — every load/refresh; autoplay without tap UI
        function IntroSplash({ onDone, logoSrc }) {
            const videoRef = useRef(null);
            const onDoneRef = useRef(onDone);
            onDoneRef.current = onDone;
            const [phase, setPhase] = useState("gate"); // gate | play
            const startedRef = useRef(false);

            useEffect(() => {
                if (phase !== "play") return;
                const v = videoRef.current;
                if (!v) return;

                let cancelled = false;
                let keepAlive = null;
                let maxTimer = null;

                const finish = () => {
                    if (cancelled) return;
                    cancelled = true;
                    if (keepAlive) clearInterval(keepAlive);
                    if (maxTimer) clearTimeout(maxTimer);
                    onDoneRef.current();
                };

                const onEnded = () => finish();
                const onError = () => finish();
                v.addEventListener("ended", onEnded);
                v.addEventListener("error", onError);

                // Keep playing with sound (gesture already unlocked on Start click)
                keepAlive = setInterval(() => {
                    if (cancelled || !videoRef.current) return;
                    const el = videoRef.current;
                    if (el.ended) {
                        finish();
                        return;
                    }
                    if (el.paused || el.currentTime < 0.05) {
                        el.muted = false;
                        el.volume = 1;
                        el.play().catch(() => {
                            el.muted = true;
                            el.play().catch(() => {});
                        });
                    }
                }, 600);

                maxTimer = setTimeout(() => {
                    if (!cancelled) finish();
                }, 15000);

                return () => {
                    cancelled = true;
                    if (keepAlive) clearInterval(keepAlive);
                    if (maxTimer) clearTimeout(maxTimer);
                    v.removeEventListener("ended", onEnded);
                    v.removeEventListener("error", onError);
                    try { v.pause(); } catch (_) {}
                };
            }, [phase]);

            // Start click = user gesture → play WITH sound in same handler
            const handleStartApp = async () => {
                if (startedRef.current) return;
                startedRef.current = true;
                const v = videoRef.current;
                if (!v) {
                    onDoneRef.current();
                    return;
                }
                v.playsInline = true;
                v.defaultMuted = false;
                v.muted = false;
                v.volume = 1;
                try {
                    v.removeAttribute("muted");
                } catch (_) {}
                try {
                    await v.play();
                } catch (_) {
                    try {
                        v.muted = true;
                        await v.play();
                    } catch (__) {
                        onDoneRef.current();
                        return;
                    }
                }
                setPhase("play");
            };

            return (
                <div className="fixed inset-0 z-[9999] w-screen h-[100dvh] overflow-hidden bg-white">
                    {/* Preload video (hidden on gate) so play() runs inside Start click gesture */}
                    <video
                        ref={videoRef}
                        src="/intro.mp4?v=sound2"
                        className={
                            phase === "play"
                                ? "absolute inset-0 w-full h-full object-cover object-center bg-black"
                                : "absolute w-px h-px opacity-0 pointer-events-none"
                        }
                        playsInline
                        preload="auto"
                        controls={false}
                    />

                    {phase === "gate" && (
                        <div className="absolute inset-0 z-10 flex flex-col items-center justify-center px-6 bg-white">
                            <div className="logo-no-vline mb-6">
                                <img
                                    src={logoSrc}
                                    alt="PG Electroplast"
                                    className="h-24 sm:h-28 w-auto object-contain"
                                />
                            </div>
                            <h1
                                className="text-3xl sm:text-4xl font-extrabold tracking-[0.14em] text-center mb-10"
                                style={{
                                    background: "linear-gradient(90deg, #0369a1 0%, #0284c7 45%, #0ea5e9 100%)",
                                    WebkitBackgroundClip: "text",
                                    backgroundClip: "text",
                                    color: "transparent",
                                }}
                            >
                                UTILITY SENSE
                            </h1>
                            <button
                                type="button"
                                onClick={handleStartApp}
                                className="h-12 px-10 rounded-full bg-[#0284c7] hover:bg-[#0369a1] text-white text-sm font-bold tracking-wide shadow-lg shadow-sky-500/25 border-none cursor-pointer transition"
                            >
                                Start App
                            </button>
                        </div>
                    )}
                </div>
            );
        }

        // ----------------------------------------------------
        // IDLE SCREENSAVER — PG logo bounce after custom idle time
        // ----------------------------------------------------
        const IDLE_MINUTE_OPTIONS = [
            { value: 0, label: "Off" },
            { value: 1, label: "1 min" },
            { value: 2, label: "2 min" },
            { value: 5, label: "5 min" },
            { value: 10, label: "10 min" },
            { value: 15, label: "15 min" },
            { value: 30, label: "30 min" },
        ];

        function IdleScreensaver({ logoSrc, idleMinutes = 5 }) {
            const [active, setActive] = useState(false);
            const logoRef = useRef(null);
            const rafRef = useRef(0);
            const timerRef = useRef(null);
            const posRef = useRef({ x: 40, y: 40, vx: 1.2, vy: 1.0 });
            const idleMs = Math.max(0, Number(idleMinutes) || 0) * 60 * 1000;
            const enabled = idleMs > 0;

            const clearIdleTimer = useCallback(() => {
                if (timerRef.current) clearTimeout(timerRef.current);
                timerRef.current = null;
            }, []);

            const armIdleTimer = useCallback(() => {
                clearIdleTimer();
                if (!enabled) {
                    setActive(false);
                    return;
                }
                timerRef.current = setTimeout(() => setActive(true), idleMs);
            }, [clearIdleTimer, enabled, idleMs]);

            const wake = useCallback(() => {
                setActive(false);
                armIdleTimer();
            }, [armIdleTimer]);

            // Activity listeners
            useEffect(() => {
                if (!enabled) {
                    clearIdleTimer();
                    setActive(false);
                    return;
                }
                const events = ["mousemove", "mousedown", "keydown", "click", "touchstart", "scroll", "wheel"];
                const onActivity = () => wake();
                events.forEach((ev) => window.addEventListener(ev, onActivity, { passive: true }));
                armIdleTimer();
                return () => {
                    events.forEach((ev) => window.removeEventListener(ev, onActivity));
                    clearIdleTimer();
                    if (rafRef.current) cancelAnimationFrame(rafRef.current);
                };
            }, [wake, armIdleTimer, clearIdleTimer, enabled]);

            // Bounce animation (DVD-style) via rAF — smooth, low cost
            useEffect(() => {
                if (!active) {
                    if (rafRef.current) cancelAnimationFrame(rafRef.current);
                    return;
                }

                const logo = logoRef.current;
                if (!logo) return;

                // Randomize start so it doesn't always begin top-left
                posRef.current = {
                    x: 40 + Math.random() * 120,
                    y: 40 + Math.random() * 80,
                    vx: (Math.random() > 0.5 ? 1 : -1) * (1.1 + Math.random() * 0.6),
                    vy: (Math.random() > 0.5 ? 1 : -1) * (0.9 + Math.random() * 0.5),
                };

                const tick = () => {
                    const el = logoRef.current;
                    if (!el) return;
                    const w = el.offsetWidth || 160;
                    const h = el.offsetHeight || 80;
                    const maxX = Math.max(0, window.innerWidth - w);
                    const maxY = Math.max(0, window.innerHeight - h);
                    let { x, y, vx, vy } = posRef.current;

                    x += vx;
                    y += vy;
                    if (x <= 0) { x = 0; vx = Math.abs(vx); }
                    else if (x >= maxX) { x = maxX; vx = -Math.abs(vx); }
                    if (y <= 0) { y = 0; vy = Math.abs(vy); }
                    else if (y >= maxY) { y = maxY; vy = -Math.abs(vy); }

                    posRef.current = { x, y, vx, vy };
                    el.style.transform = `translate3d(${x}px, ${y}px, 0)`;
                    rafRef.current = requestAnimationFrame(tick);
                };

                rafRef.current = requestAnimationFrame(tick);
                return () => {
                    if (rafRef.current) cancelAnimationFrame(rafRef.current);
                };
            }, [active]);

            if (!active) return null;

            return (
                <div
                    className="idle-screensaver"
                    role="presentation"
                    onMouseMove={wake}
                    onClick={wake}
                    onKeyDown={wake}
                    tabIndex={-1}
                >
                    <img
                        ref={logoRef}
                        src={logoSrc}
                        alt="PG Groups"
                        className="idle-screensaver__logo"
                        draggable={false}
                    />
                </div>
            );
        }

        // ---------------------------------------------------- 
        function App() {
            // Authentication States
            const [currentUser, setCurrentUser] = useState(() => {
                try {
                    const saved = localStorage.getItem("ep_session");
                    if (saved) {
                        const parsed = JSON.parse(saved);
                        if (parsed && parsed.user) {
                            return { ...parsed.user, loggedAt: parsed.loggedAt };
                        }
                        if (Date.now() - parsed.loggedAt < 7 * 60 * 60 * 1000) return parsed;
                    }
                } catch (e) { }
                return null;
            });

            // Always show intro on every page load / refresh (same tab too)
            const [introDone, setIntroDone] = useState(false);
            const finishIntro = useCallback(() => setIntroDone(true), []);

            // Navigation States
            const [activeTab, setActiveTab] = useState("dashboard");
            const [kpiLayout, setKpiLayout] = useState("grid");

            // Theme States
            const [theme, setTheme] = useState(() => {
                try { return localStorage.getItem("ep_theme") || "light"; } catch (e) { return "light"; }
            });
            const [themePickerOpen, setThemePickerOpen] = useState(false);
            const [idlePickerOpen, setIdlePickerOpen] = useState(false);
            const [idleMinutes, setIdleMinutes] = useState(() => {
                try {
                    const saved = Number(localStorage.getItem("ep_idle_minutes"));
                    if (Number.isFinite(saved) && saved >= 0) return saved;
                } catch (e) { /* ignore */ }
                return 5;
            });
            const [filtersOpen, setFiltersOpen] = useState(false); // Dashboard filter panel is hidden by default, toggled via a button
            const filterPanelRef = useRef(null);
            const filterToggleBtnRef = useRef(null);

            useEffect(() => {
                if (!filtersOpen) return;
                const handleAutoClose = (e) => {
                    if (e && e.target) {
                        if (filterPanelRef.current && filterPanelRef.current.contains(e.target)) return;
                        if (filterToggleBtnRef.current && filterToggleBtnRef.current.contains(e.target)) return;
                    }
                    setFiltersOpen(false);
                };
                window.addEventListener("scroll", handleAutoClose, { passive: true });
                document.addEventListener("mousedown", handleAutoClose);
                return () => {
                    window.removeEventListener("scroll", handleAutoClose);
                    document.removeEventListener("mousedown", handleAutoClose);
                };
            }, [filtersOpen]);

            const [reportsFiltersOpen, setReportsFiltersOpen] = useState(false); // Reports filter panel toggle state
            const [canMonitor, setCanMonitor] = useState(() => {
                try {
                    return localStorage.getItem("ep_can_monitor") === "true";
                } catch (e) {
                    return false;
                }
            });
            const handleToggleCanMonitor = () => {
                setCanMonitor(prev => {
                    const nextVal = !prev;
                    try {
                        localStorage.setItem("ep_can_monitor", String(nextVal));
                    } catch (e) {}
                    return nextVal;
                });
            };
            useEffect(() => {
                document.documentElement.setAttribute("data-theme", theme);
                try { localStorage.setItem("ep_theme", theme); } catch (e) { }
            }, [theme]);

            useEffect(() => {
                try { localStorage.setItem("ep_idle_minutes", String(idleMinutes)); } catch (e) { }
            }, [idleMinutes]);

            // Selection States for Multi-Delete
            const [selectedRowIds, setSelectedRowIds] = useState(new Set());
            const [selectedMasterRowKeys, setSelectedMasterRowKeys] = useState(new Set());
            const [isDailyDeleteMode, setIsDailyDeleteMode] = useState(false);
            const [isMasterDeleteMode, setIsMasterDeleteMode] = useState(false);

            // Database States
            const [dailyEntries, setDailyEntries] = useState([]);
            const [plants, setPlants] = useState([]);
            const [departments, setDepartments] = useState([]);
            const [meters, setMeters] = useState([]);
            const [solarMeters, setSolarMeters] = useState([]);
            const [waterMeters, setWaterMeters] = useState([]);
            const [airMeters, setAirMeters] = useState([]);
            const [dgSets, setDgSets] = useState([]);
            const [fuelTypes, setFuelTypes] = useState([]);
            const [products, setProducts] = useState([]);
            const [tariffs, setTariffs] = useState([]);
            const [multiplyFactors, setMultiplyFactors] = useState([]);
            const [targets, setTargets] = useState([]);
            const [users, setUsers] = useState([]);
            const [otpLogs, setOtpLogs] = useState([]);
            const [auditLogs, setAuditLogs] = useState([]);
            const [selectedAuditDetail, setSelectedAuditDetail] = useState(null);
            const [auditSearchQuery, setAuditSearchQuery] = useState("");
            const [auditModuleFilter, setAuditModuleFilter] = useState("all");

            const [emailSchedules, setEmailSchedules] = useState([]);
            const [emailScheduleLogs, setEmailScheduleLogs] = useState([]);
            const [isScheduleModalOpen, setIsScheduleModalOpen] = useState(false);
            const [editingSchedule, setEditingSchedule] = useState(null);
            const [isScheduleLogsOpen, setIsScheduleLogsOpen] = useState(false);
            const [runningScheduleId, setRunningScheduleId] = useState(null);
            const [scheduleFormValues, setScheduleFormValues] = useState({
                id: "",
                name: "",
                location: "",
                plant: "",
                report_type: "Monthly Utility Report",
                frequency: "Monthly",
                schedule_day: 5,
                schedule_time: "10:00",
                to_recipients: "",
                cc_recipients: "",
                bcc_recipients: "",
                subject_template: "Monthly Utility Report - {Location} - {Plant} - {Month}",
                body_template: "Dear Team,\n\nPlease find attached the monthly UtilitySense report for {Location} - {Plant} for {Month}.\n\nPlease review the report and take necessary action where required.\n\nRegards,\nUtilitySense Management",
                enabled: true
            });
            const [isEmailSending, setIsEmailSending] = useState(false);

            // Mass Excel Importer / Bulk Data Uploader States (IT Admin Only)
            const [isMassUploadModalOpen, setIsMassUploadModalOpen] = useState(false);
            const [isImportHistoryOpen, setIsImportHistoryOpen] = useState(false);
            const [importHistory, setImportHistory] = useState([]);
            const [massUploadFileName, setMassUploadFileName] = useState("");
            const [massUploadSheets, setMassUploadSheets] = useState([]);
            const [selectedUploadSheet, setSelectedUploadSheet] = useState("");
            const [rawUploadData, setRawUploadData] = useState([]); // Raw sheet JSON rows
            const [detectedHeaders, setDetectedHeaders] = useState([]);
            const [columnMappings, setColumnMappings] = useState({
                date: "",
                location: "",
                plant: "",
                electricity_closing: "",
                electricity_opening: "",
                solar: "",
                diesel: "",
                odu: "",
                idu: "",
                production_set: "",
                waste_hazardous: "",
                waste_non_hazardous: "",
                waste_recycled: "",
                water: "",
                gas: "",
                air: "",
                remarks: "",
                operator: ""
            });
            const [locationOverride, setLocationOverride] = useState("auto"); // "auto" | location name
            const [plantOverride, setPlantOverride] = useState("auto"); // "auto" | plant code
            const [analyzedImportRows, setAnalyzedImportRows] = useState([]);
            const [importFilterTab, setImportFilterTab] = useState("all"); // "all" | "valid" | "duplicate" | "invalid"
            const [isAnalyzingExcel, setIsAnalyzingExcel] = useState(false);
            const [isImportingData, setIsImportingData] = useState(false);
            const [importResultSummary, setImportResultSummary] = useState(null);

            // Smart Report Export Filters
            const [reportRangeMode, setReportRangeMode] = useState("custom"); // "month" | "custom"
            const [reportFromMonth, setReportFromMonth] = useState(() => {
                const d = new Date();
                return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
            });
            const [reportToMonth, setReportToMonth] = useState(() => {
                const d = new Date();
                return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
            });
            const [reportCustomStartDate, setReportCustomStartDate] = useState("");
            const [reportCustomEndDate, setReportCustomEndDate] = useState("");

            // Filter plants based on user's allowed locations/plants permissions
            const allowedPlants = useMemo(() => {
                if (!currentUser) return [];
                if (currentUser.role === "IT_ADMIN") return plants;

                const allowedP = String(currentUser.allowed_plants || "all").trim().toLowerCase();
                const allowedL = String(currentUser.allowed_locations || "all").trim().toLowerCase();

                return plants.filter(p => {
                    const pLoc = p.location.toLowerCase();
                    const pCode = p.plant_code.toLowerCase();
                    const matchesPlant = allowedP === "all" || allowedP === "" || allowedP.split(",").map(x => x.trim()).includes(pCode);
                    const matchesLoc = allowedL === "all" || allowedL === "" || allowedL.includes(pLoc) || pLoc.includes(allowedL);
                    return matchesPlant && matchesLoc;
                });
            }, [plants, currentUser]);

            const allowedLocations = useMemo(() => {
                if (!currentUser) return [];
                if (currentUser.role === "IT_ADMIN") {
                    return Array.from(new Set(plants.map(p => p.location.toUpperCase())));
                }
                return Array.from(new Set(allowedPlants.map(p => p.location.toUpperCase())));
            }, [plants, allowedPlants, currentUser]);

            // Status States
            const [loading, setLoading] = useState(true);
            const [actionLoading, setActionLoading] = useState(false);
            const [error, setError] = useState(null);
            const [toast, setToast] = useState(null);
            const [confirmDialog, setConfirmDialog] = useState(null); // { title, message, confirmText, danger, resolve }

            // OneDrive Sync States

            // Authentication Forms States
            const [loginEmail, setLoginEmail] = useState("");
            const [loginOtp, setLoginOtp] = useState("");
            const [otpSent, setOtpSent] = useState(false);
            const [loginLoading, setLoginLoading] = useState(false);
            const [loginError, setLoginError] = useState("");
            const [loginMessage, setLoginMessage] = useState("");

            const [filters, setFilters] = useState(() => getDefaultDateFilters());

            // Filter allowedPlants by the currently selected location filter
            const filteredPlantsForDropdown = useMemo(() => {
                if (!filters.location || filters.location === "all") {
                    return allowedPlants;
                }
                return allowedPlants.filter(p => p.location.toLowerCase() === filters.location.toLowerCase());
            }, [allowedPlants, filters.location]);

            const dashboardRateContext = useMemo(() => {
                const today = new Date().toISOString().split("T")[0];
                const plantCode = filters.plant !== "all" ? filters.plant : (allowedPlants[0]?.plant_code || "");
                const plantObj = plants.find((p) => p.plant_code === plantCode) || allowedPlants[0];
                const location = filters.location !== "all"
                    ? filters.location
                    : (plantObj?.location || "");
                return { plantCode, location, date: today };
            }, [filters.plant, filters.location, plants, allowedPlants]);

            const dashboardGridLabel = useMemo(() => {
                const pObj = plants.find(p => p.plant_code === filters.plant);
                const loc = filters.location !== "all" ? filters.location : (pObj?.location || (allowedPlants[0]?.location || ""));
                return getGridProviderLabel(loc);
            }, [filters.location, filters.plant, plants, allowedPlants]);

            const activeElectRate = useMemo(() => {
                const { plantCode, location, date } = dashboardRateContext;
                return resolveTariff(tariffs, "electricity", plantCode, location, date);
            }, [tariffs, dashboardRateContext]);

            const activeDieselRate = useMemo(() => {
                const { plantCode, location, date } = dashboardRateContext;
                return resolveTariff(tariffs, "diesel", plantCode, location, date);
            }, [tariffs, dashboardRateContext]);

            const activeWaterRate = useMemo(() => {
                const { plantCode, location, date } = dashboardRateContext;
                return resolveTariff(tariffs, "water", plantCode, location, date);
            }, [tariffs, dashboardRateContext]);

            const activeLpgRate = useMemo(() => {
                const { plantCode, location, date } = dashboardRateContext;
                return resolveTariff(tariffs, "lpg", plantCode, location, date);
            }, [tariffs, dashboardRateContext]);

            const activeSolarRate = useMemo(() => {
                const { plantCode, location, date } = dashboardRateContext;
                const solar = resolveTariff(tariffs, "solar", plantCode, location, date);
                return solar || activeElectRate;
            }, [tariffs, dashboardRateContext, activeElectRate]);

            // Re-sync plant filter if location filter changes to prevent mismatched states
            const handleLocationFilterChange = (locVal) => {
                setFilters(prev => {
                    const next = { ...prev, location: locVal };
                    if (locVal !== "all" && prev.plant !== "all") {
                        const matchPlant = plants.find(p => p.plant_code === prev.plant);
                        if (matchPlant && matchPlant.location !== locVal) {
                            next.plant = "all";
                        }
                    }
                    return next;
                });
            };

            // Daily Data Entry Form States
            const [isFormOpen, setIsFormOpen] = useState(false);
            const [editingRecord, setEditingRecord] = useState(null);
            const [wasteSectionOpen, setWasteSectionOpen] = useState(false);
            const [productionExcelData, setProductionExcelData] = useState(null);
            const [isFetchingExcelData, setIsFetchingExcelData] = useState(false);
            const [excelFileName, setExcelFileName] = useState("");
            const [viewingRecord, setViewingRecord] = useState(null); // Read-only "View Details" popup for a clicked row
            const [entryFormValues, setEntryFormValues] = useState({
                date: new Date().toISOString().split('T')[0],
                plant: "",
                location: "",
                department: "PROD",
                shift: "Shift A",
                operator_name: "",
                electricity_opening: 0,
                electricity_closing: "",
                meter_changed: false,
                custom_difference: "",
                electricity_consumption: 0,
                electricity_cost: 0,
                solar_opening: 0,
                solar_closing: "",
                solar_generated: "",
                solar_utilized: 0,
                solar_cost: 0,
                solar_utilization_pct: 100,
                diesel_used: "",
                diesel_cost: 0,
                total_cost: 0,
                odu: "",
                idu: "",
                production_set: 0,
                cost_per_set: 0,
                waste_hazardous: "",
                waste_non_hazardous: "",
                waste_recycled: "",
                remarks: ""
            });

            const entryRateContext = useMemo(() => {
                const plantObj = plants.find((p) => p.plant_code === entryFormValues.plant);
                const location = entryFormValues.location || plantObj?.location || "";
                const date = entryFormValues.date || new Date().toISOString().split("T")[0];
                const electRate = resolveTariff(tariffs, "electricity", entryFormValues.plant, location, date);
                return {
                    location,
                    date,
                    plantCode: entryFormValues.plant,
                    mf: resolveMultiplyFactor(multiplyFactors, entryFormValues.plant, location, date),
                    electRate,
                    solarRate: resolveTariff(tariffs, "solar", entryFormValues.plant, location, date) || electRate,
                    dieselRate: resolveTariff(tariffs, "diesel", entryFormValues.plant, location, date),
                    gridLabel: getGridProviderLabel(location),
                };
            }, [entryFormValues.plant, entryFormValues.location, entryFormValues.date, plants, multiplyFactors, tariffs]);

            // Master Data Config Editor States
            const [selectedMasterTable, setSelectedMasterTable] = useState("plants");
            const [isMasterFormOpen, setIsMasterFormOpen] = useState(false);
            const [editingMasterRecord, setEditingMasterRecord] = useState(null);
            const [masterFormValues, setMasterFormValues] = useState({});

            // Reports Panel States
            const [selectedReportType, setSelectedReportType] = useState("daily");
            const [selectedReportLocation, setSelectedReportLocation] = useState("all");

            const reportGridLabel = useMemo(() => {
                const loc = selectedReportLocation !== "all" ? selectedReportLocation : (allowedPlants[0]?.location || "");
                return getGridProviderLabel(loc);
            }, [selectedReportLocation, allowedPlants]);

            // Daily Data entry pagination & search
            const [entrySearch, setEntrySearch] = useState("");
            const [entryLocationFilter, setEntryLocationFilter] = useState("all");
            const [entryPlantFilter, setEntryPlantFilter] = useState("all");
            const [entryPage, setEntryPage] = useState(1);
            const [selectedEntryIds, setSelectedEntryIds] = useState([]);
            const entryLimit = 8;

            const entryTableContext = useMemo(() => {
                const today = new Date().toISOString().split("T")[0];
                const loc = entryLocationFilter !== "all"
                    ? entryLocationFilter
                    : (entryPlantFilter !== "all" ? plants.find((p) => p.plant_code === entryPlantFilter)?.location : "");
                const plantCode = entryPlantFilter !== "all" ? entryPlantFilter : "";
                return {
                    gridLabel: loc ? getGridProviderLabel(loc) : "Grid",
                    mf: (loc || plantCode)
                        ? resolveMultiplyFactor(multiplyFactors, plantCode, loc, today)
                        : null,
                };
            }, [entryLocationFilter, entryPlantFilter, plants, multiplyFactors]);

            const lastActiveTime = useRef(Date.now());

            // Toast timeout handler
            useEffect(() => {
                if (toast) {
                    const timer = setTimeout(() => setToast(null), toast.duration || 4000);
                    return () => clearTimeout(timer);
                }
            }, [toast]);

            // Custom confirm dialog helper — replaces native window.confirm() with a styled modal.
            // Usage: const ok = await askConfirm("Are you sure?", { title: "Delete Record", confirmText: "Delete" });
            const askConfirm = (message, options = {}) => {
                return new Promise((resolve) => {
                    setConfirmDialog({
                        title: options.title || "Confirm Action",
                        message,
                        confirmText: options.confirmText || "Confirm",
                        cancelText: options.cancelText || "Cancel",
                        danger: options.danger !== false,
                        resolve
                    });
                });
            };

            // Reset pagination & selection on search / filter change
            useEffect(() => {
                setEntryPage(1);
                setSelectedEntryIds([]);
            }, [entrySearch, entryLocationFilter, entryPlantFilter]);

            // Session auto-logout inactivity tracker (7 hours)
            useEffect(() => {
                if (!currentUser) return;
                const updateSessionTimer = () => {
                    lastActiveTime.current = Date.now();
                    try {
                        const saved = localStorage.getItem("ep_session");
                        if (saved) {
                            const parsed = JSON.parse(saved);
                            parsed.loggedAt = Date.now();
                            localStorage.setItem("ep_session", JSON.stringify(parsed));
                        }
                    } catch (e) { }
                };

                const events = ["mousemove", "mousedown", "keypress", "scroll", "touchstart"];
                events.forEach(ev => window.addEventListener(ev, updateSessionTimer));

                const checkInterval = setInterval(() => {
                    if (Date.now() - lastActiveTime.current > 7 * 60 * 60 * 1000) {
                        handleLogout("Session expired due to inactivity.");
                    }
                }, 15000);

                return () => {
                    events.forEach(ev => window.removeEventListener(ev, updateSessionTimer));
                    clearInterval(checkInterval);
                };
            }, [currentUser]);

            // Load DB on login — reset filters to All Plants / All Locations
            useEffect(() => {
                if (!currentUser) return;
                setFilters(getDefaultDateFilters());
                loadAllDatabase();
            }, [currentUser]);

            // Load data files from Google Apps Script side or mock fallback offline data
                        const loadAllDatabase = async () => {
                setLoading(true);
                setError(null);
                try {
                    // Load daily entries
                    const { data: entries, error: entriesErr } = await supabase
                        .from('daily_entries')
                        .select('*')
                        .order('date', { ascending: false });
                    if (entriesErr) throw entriesErr;
                    const rows = entries || [];
                    setDailyEntries(rows);

                    // Default date window = full available data so KPI cards populate on login
                    if (rows.length > 0) {
                        const dates = rows.map((e) => e.date).filter(Boolean).sort();
                        const minDate = dates[0];
                        const maxDate = dates[dates.length - 1];
                        const today = toISODate(new Date());
                        setFilters({
                            startDate: minDate,
                            endDate: maxDate > today ? maxDate : today,
                            plant: "all",
                            department: "all",
                            location: "all",
                        });
                    } else {
                        setFilters(getDefaultDateFilters());
                    }

                    // Load configs safely with fallback
                    const loadConfig = async (table, setter) => {
                        try {
                            const { data, error } = await supabase.from(table).select('*');
                            if (!error && data) {
                                setter(data);
                            }
                        } catch (err) {
                            console.info(`Table '${table}' will be initialized once created in Supabase.`);
                        }
                    };

                    await Promise.all([
                        loadConfig('plants', setPlants),
                        loadConfig('departments', setDepartments),
                        loadConfig('meters', setMeters),
                        loadConfig('solar_meters', setSolarMeters),
                        loadConfig('water_meters', setWaterMeters),
                        loadConfig('air_meters', setAirMeters),
                        loadConfig('dg_sets', setDgSets),
                        loadConfig('fuel_types', setFuelTypes),
                        loadConfig('products', setProducts),
                        loadConfig('tariffs', setTariffs),
                        loadConfig('multiply_factors', setMultiplyFactors),
                        loadConfig('target_values', setTargets),
                        loadConfig('users', setUsers),
                        loadConfig('otp_logs', setOtpLogs),
                        loadConfig('audit_logs', setAuditLogs),
                        loadConfig('email_schedules', (rows) => {
                            if (rows && rows.length > 0) {
                                setEmailSchedules(rows);
                            } else {
                                // Default initial template if table is newly created
                                const defaultSchedules = [
                                    {
                                        id: "sched_bhiwadi_default",
                                        name: "Bhiwadi Unit Monthly Report",
                                        location: "BHIWADI",
                                        plant: "PLANT-1",
                                        report_type: "Monthly Utility Report",
                                        frequency: "Monthly",
                                        schedule_day: 5,
                                        schedule_time: "10:00",
                                        to_recipients: "factory.head@pgel.in",
                                        cc_recipients: "utility.hod@pgel.in",
                                        bcc_recipients: "",
                                        subject_template: "Monthly Utility Report - {Location} - {Plant} - {Month}",
                                        body_template: "Dear Team,\n\nPlease find attached the monthly UtilitySense report for {Location} - {Plant} for {Month}.\n\nPlease review the report and take necessary action where required.\n\nRegards,\nUtilitySense Management",
                                        enabled: true,
                                        last_run_at: null,
                                        last_status: null
                                    }
                                ];
                                setEmailSchedules(defaultSchedules);
                            }
                        }),
                        loadConfig('email_schedule_logs', setEmailScheduleLogs),
                        loadConfig('import_history', setImportHistory)
                    ]);
                } catch (err) {
                    console.error("Database fetch failed:", err);
                    setError(err.message || "Database connection failed");
                } finally {
                    setLoading(false);
                }
            };

            function loadConfigSheet() {}

            function generateMockEntries() {
                const list = [];
                // Set up cumulative readings for Mock Seeding
                let electReadings = { NGM: 10000, PGEL: 20000, PGTL: 15000 };
                let waterReadings = { NGM: 1500, PGEL: 3000, PGTL: 2000 };
                let airReadings = { NGM: 4000, PGEL: 8000, PGTL: 6000 };

                let dieselStocks = { NGM: 500, PGEL: 1000, PGTL: 800 };
                let lpgStocks = { NGM: 300, PGEL: 600, PGTL: 400 };

                for (let day = 1; day <= 24; day++) {
                    const dateStr = `2026-06-${String(day).padStart(2, "0")}`;
                    const plantsList = ["NGM", "PGEL", "PGTL"];

                    plantsList.forEach(pCode => {
                        const deptCode = pCode === "PGTL" ? "UTILS" : "PROD";
                        const eUse = Math.round(300 + Math.random() * 300);
                        const wUse = Math.round(6 + Math.random() * 12);
                        const aUse = Math.round(30 + Math.random() * 50);

                        const eOpen = electReadings[pCode];
                        const eClose = eOpen + eUse;
                        electReadings[pCode] = eClose;

                        const wOpen = waterReadings[pCode];
                        const wClose = wOpen + wUse;
                        waterReadings[pCode] = wClose;

                        const aOpen = airReadings[pCode];
                        const aClose = aOpen + aUse;
                        airReadings[pCode] = aClose;

                        const solGen = Math.round(150 + Math.random() * 250);
                        const solUtil = Math.round(solGen * (0.7 + Math.random() * 0.25));
                        const solExp = Math.max(0, solGen - solUtil);

                        const dOpen = dieselStocks[pCode];
                        const dRec = day % 6 === 1 ? 500 : 0;
                        const dUsed = Math.round(15 + Math.random() * 25);
                        const dClose = dOpen + dRec - dUsed;
                        dieselStocks[pCode] = dClose;

                        const lOpen = lpgStocks[pCode];
                        const lRec = day % 8 === 1 ? 300 : 0;
                        const lUsed = Math.round(8 + Math.random() * 14);
                        const lClose = lOpen + lRec - lUsed;
                        lpgStocks[pCode] = lClose;

                        const prodQty = Math.round(40 + Math.random() * 60);

                        const eCost = eUse * 8.5;
                        const wCost = wUse * 45;
                        const dCost = dUsed * 92;
                        const lCost = lUsed * 85;

                        list.push({
                            id: `mock-${pCode}-${day}`,
                            date: dateStr,
                            plant: pCode,
                            department: deptCode,
                            shift: day % 2 === 0 ? "Shift A" : "Shift B",
                            operator_name: "Demo Operator",
                            electricity_meter: `EM-${pCode === 'NGM' ? '01' : (pCode === 'PGEL' ? '02' : '03')}`,
                            electricity_opening: eOpen,
                            electricity_closing: eClose,
                            electricity_consumption: eUse,
                            electricity_cost: eCost,
                            solar_meter: `SM-${pCode === 'NGM' ? '01' : (pCode === 'PGEL' ? '02' : '03')}`,
                            solar_generated: solGen,
                            solar_utilized: solUtil,
                            solar_exported: solExp,
                            solar_utilization_pct: solGen > 0 ? (solUtil / solGen) * 100 : 0,
                            water_meter: `WM-${pCode === 'NGM' ? '01' : (pCode === 'PGEL' ? '02' : '03')}`,
                            water_opening: wOpen,
                            water_closing: wClose,
                            water_consumption: wUse,
                            water_cost: wCost,
                            air_meter: `AM-${pCode === 'NGM' ? '01' : (pCode === 'PGEL' ? '02' : '03')}`,
                            air_opening: aOpen,
                            air_closing: aClose,
                            air_consumption: aUse,
                            diesel_dg_set: `DG-${pCode === 'NGM' ? '01' : (pCode === 'PGEL' ? '02' : '03')}`,
                            diesel_opening: dOpen,
                            diesel_received: dRec,
                            diesel_closing: dClose,
                            diesel_used: dUsed,
                            diesel_cost: dCost,
                            lpg_opening: lOpen,
                            lpg_received: lRec,
                            lpg_closing: lClose,
                            lpg_used: lUsed,
                            lpg_cost: lCost,
                            product_name: "COMP-A",
                            production_qty: prodQty,
                            production_unit: "Sets",
                            waste_hazardous: Math.round(1 + Math.random() * 2),
                            waste_non_hazardous: Math.round(5 + Math.random() * 8),
                            waste_recycled: Math.round(4 + Math.random() * 5),
                            sec: eUse / prodQty,
                            specific_water_consumption: wUse / prodQty,
                            remarks: "Offline mock entry",
                            created_at: new Date().toISOString(),
                            created_by: "offline"
                        });
                    });
                }
                return list.sort((a, b) => b.date.localeCompare(a.date));
            }

            // ----------------------------------------------------
            // AUDIT LOGGING & EMAIL AUTOMATION SERVICES
            // ----------------------------------------------------
            const recordAuditLog = useCallback(async ({ action, module, recordId, location, plant, oldValue, newValue, status = "SUCCESS" }) => {
                try {
                    const userEmail = currentUser?.email || "system@pgel.in";
                    const logItem = {
                        user_email: userEmail,
                        user_id: String(currentUser?.id || ""),
                        user_name: currentUser?.name || currentUser?.operator_name || userEmail.split('@')[0],
                        action: String(action || "UNKNOWN").toUpperCase(),
                        module: String(module || "General"),
                        record_id: recordId ? String(recordId) : null,
                        location: location || null,
                        plant: plant || null,
                        old_value: oldValue ? (typeof oldValue === "object" ? oldValue : { value: oldValue }) : null,
                        new_value: newValue ? (typeof newValue === "object" ? newValue : { value: newValue }) : null,
                        status: status || "SUCCESS",
                        ip_device: typeof navigator !== "undefined" ? navigator.userAgent : null,
                        created_at: new Date().toISOString()
                    };
                    setAuditLogs(prev => [logItem, ...prev]);
                    supabase.from('audit_logs').insert([logItem]).then(({ error }) => {
                        if (error) console.warn("Supabase audit log insert:", error.message);
                    });
                } catch (e) {
                    console.warn("recordAuditLog caught:", e);
                }
            }, [currentUser]);

            // ----------------------------------------------------
            // MULTI-SCHEDULE EMAIL AUTOMATION HANDLERS (IT ADMIN)
            // ----------------------------------------------------
            const openCreateScheduleModal = () => {
                const firstLoc = allowedLocations[0] || "BHIWADI";
                const matchingPlants = plants.filter(p => p.location.toUpperCase() === firstLoc.toUpperCase());
                const firstPlant = matchingPlants[0]?.plant_code || "";

                setEditingSchedule(null);
                setScheduleFormValues({
                    id: `sched_${Date.now()}`,
                    name: `${firstLoc} - ${firstPlant} Monthly Schedule`,
                    location: firstLoc,
                    plant: firstPlant,
                    report_type: "Monthly Utility Report",
                    frequency: "Monthly",
                    schedule_day: 5,
                    schedule_time: "10:00",
                    to_recipients: "",
                    cc_recipients: "",
                    bcc_recipients: "",
                    subject_template: "Monthly Utility Report - {Location} - {Plant} - {Month}",
                    body_template: "Dear Team,\n\nPlease find attached the monthly UtilitySense report for {Location} - {Plant} for {Month}.\n\nPlease review the report and take necessary action where required.\n\nRegards,\nUtilitySense Management",
                    enabled: true
                });
                setIsScheduleModalOpen(true);
            };

            const openEditScheduleModal = (sched) => {
                setEditingSchedule(sched);
                setScheduleFormValues({
                    id: sched.id,
                    name: sched.name || "",
                    location: sched.location || "",
                    plant: sched.plant || "",
                    report_type: sched.report_type || "Monthly Utility Report",
                    frequency: sched.frequency || "Monthly",
                    schedule_day: sched.schedule_day || 5,
                    schedule_time: sched.schedule_time || "10:00",
                    to_recipients: sched.to_recipients || "",
                    cc_recipients: sched.cc_recipients || "",
                    bcc_recipients: sched.bcc_recipients || "",
                    subject_template: sched.subject_template || "Monthly Utility Report - {Location} - {Plant} - {Month}",
                    body_template: sched.body_template || "Dear Team,\n\nPlease find attached the monthly UtilitySense report for {Location} - {Plant} for {Month}.\n\nPlease review the report and take necessary action where required.\n\nRegards,\nUtilitySense Management",
                    enabled: sched.enabled !== false
                });
                setIsScheduleModalOpen(true);
            };

            const handleSaveSchedule = async (e) => {
                if (e) e.preventDefault();
                if (!scheduleFormValues.name.trim()) {
                    setToast({ type: "error", message: "Schedule name is required" });
                    return;
                }
                if (!scheduleFormValues.location) {
                    setToast({ type: "error", message: "Please select a valid location" });
                    return;
                }
                if (!scheduleFormValues.plant) {
                    setToast({ type: "error", message: "Please select a valid plant" });
                    return;
                }
                if (!scheduleFormValues.to_recipients.trim()) {
                    setToast({ type: "error", message: "At least one valid 'To' recipient email is required" });
                    return;
                }

                // Email validation helper
                const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
                const toList = scheduleFormValues.to_recipients.split(",").map(x => x.trim()).filter(Boolean);
                const invalidTo = toList.filter(em => !emailRegex.test(em));
                if (invalidTo.length > 0) {
                    setToast({ type: "error", message: `Invalid 'To' email address: ${invalidTo[0]}` });
                    return;
                }

                if (scheduleFormValues.cc_recipients.trim()) {
                    const ccList = scheduleFormValues.cc_recipients.split(",").map(x => x.trim()).filter(Boolean);
                    const invalidCc = ccList.filter(em => !emailRegex.test(em));
                    if (invalidCc.length > 0) {
                        setToast({ type: "error", message: `Invalid 'CC' email address: ${invalidCc[0]}` });
                        return;
                    }
                }

                setActionLoading(true);
                try {
                    const scheduleRecord = {
                        ...scheduleFormValues,
                        id: scheduleFormValues.id || `sched_${Date.now()}`,
                        updated_at: new Date().toISOString()
                    };

                    const { error } = await supabase.from('email_schedules').upsert(scheduleRecord);
                    if (error) {
                        console.warn("Supabase upsert note:", error.message);
                    }

                    // Update local state
                    setEmailSchedules(prev => {
                        const existingIdx = prev.findIndex(s => s.id === scheduleRecord.id);
                        if (existingIdx >= 0) {
                            const copy = [...prev];
                            copy[existingIdx] = scheduleRecord;
                            return copy;
                        } else {
                            return [...prev, scheduleRecord];
                        }
                    });

                    recordAuditLog({
                        action: editingSchedule ? "UPDATE" : "CREATE",
                        module: "Email Automation",
                        recordId: scheduleRecord.id,
                        location: scheduleRecord.location,
                        plant: scheduleRecord.plant,
                        oldValue: editingSchedule || null,
                        newValue: scheduleRecord,
                        status: "SUCCESS"
                    });

                    setToast({
                        type: "success",
                        message: editingSchedule ? "Email schedule updated successfully!" : "New email schedule created successfully!"
                    });
                    setIsScheduleModalOpen(false);
                } catch (err) {
                    console.error("Save schedule failed:", err);
                    setToast({ type: "error", message: `Save failed: ${err.message}` });
                } finally {
                    setActionLoading(false);
                }
            };

            const handleDeleteSchedule = async (scheduleId) => {
                const target = emailSchedules.find(s => s.id === scheduleId);
                const confirmed = await openConfirm({
                    title: "Delete Email Schedule",
                    message: `Are you sure you want to permanently delete the schedule "${target?.name || scheduleId}"?`,
                    danger: true,
                    confirmText: "Delete Schedule"
                });
                if (!confirmed) return;

                setActionLoading(true);
                try {
                    const { error } = await supabase.from('email_schedules').delete().eq('id', scheduleId);
                    if (error) console.warn("Supabase delete note:", error.message);

                    setEmailSchedules(prev => prev.filter(s => s.id !== scheduleId));

                    recordAuditLog({
                        action: "DELETE",
                        module: "Email Automation",
                        recordId: scheduleId,
                        location: target?.location,
                        plant: target?.plant,
                        oldValue: target,
                        status: "SUCCESS"
                    });

                    setToast({ type: "success", message: "Email schedule deleted successfully." });
                } catch (err) {
                    console.error("Delete schedule error:", err);
                    setToast({ type: "error", message: `Delete failed: ${err.message}` });
                } finally {
                    setActionLoading(false);
                }
            };

            const handleToggleScheduleStatus = async (scheduleId, currentStatus) => {
                const newStatus = !currentStatus;
                try {
                    const { error } = await supabase
                        .from('email_schedules')
                        .update({ enabled: newStatus, updated_at: new Date().toISOString() })
                        .eq('id', scheduleId);
                    if (error) console.warn("Supabase toggle note:", error.message);

                    setEmailSchedules(prev => prev.map(s => s.id === scheduleId ? { ...s, enabled: newStatus } : s));

                    recordAuditLog({
                        action: newStatus ? "ENABLE" : "DISABLE",
                        module: "Email Automation",
                        recordId: scheduleId,
                        newValue: { enabled: newStatus },
                        status: "SUCCESS"
                    });

                    setToast({
                        type: "success",
                        message: `Schedule "${scheduleId}" is now ${newStatus ? "ACTIVE (ON)" : "PAUSED (OFF)"}.`
                    });
                } catch (err) {
                    console.error("Toggle schedule error:", err);
                    setToast({ type: "error", message: `Update failed: ${err.message}` });
                }
            };

            const handleExecuteSchedule = async ({ schedule, isManualTest = false, targetMonth = null }) => {
                if (!schedule) return;
                setRunningScheduleId(schedule.id);
                try {
                    const emailLoc = schedule.location;
                    const emailPlant = schedule.plant;

                    let monthStr = targetMonth;
                    if (!monthStr) {
                        const prevD = new Date();
                        prevD.setDate(1);
                        prevD.setMonth(prevD.getMonth() - 1);
                        monthStr = `${prevD.getFullYear()}-${String(prevD.getMonth() + 1).padStart(2, '0')}`;
                    }

                    const startDate = `${monthStr}-01`;
                    const endDate = getMonthEnd(monthStr);

                    // Build strictly scoped report data
                    const payload = buildPlantReportData({
                        location: emailLoc,
                        plant: emailPlant,
                        startDate,
                        endDate
                    });

                    if (!payload || !payload.list || payload.list.length === 0) {
                        throw new Error(`No daily entry records found for ${emailLoc} - ${emailPlant} during ${monthStr}.`);
                    }

                    // Generate exact Excel workbook
                    const workbook = await createPlantReportWorkbook(payload);
                    const buffer = await workbook.xlsx.writeBuffer();

                    // Convert ArrayBuffer to binary string
                    const bytes = new Uint8Array(buffer);
                    let binary = '';
                    for (let i = 0; i < bytes.byteLength; i++) {
                        binary += String.fromCharCode(bytes[i]);
                    }
                    const base64Attachment = btoa(binary);

                    const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
                    const [y, mNum] = monthStr.split("-").map(Number);
                    const readableMonth = `${monthNames[(mNum || 1) - 1]} ${y}`;

                    // Dynamic template token replacement
                    const formatTemplate = (tmpl) => {
                        return String(tmpl || "")
                            .replace(/\{Location\}/g, payload.locationLabel || emailLoc)
                            .replace(/\{Plant\}/g, payload.plantLabel || emailPlant)
                            .replace(/\{Month\}/g, readableMonth)
                            .replace(/\{Year\}/g, String(y || new Date().getFullYear()))
                            .replace(/\{ReportDate\}/g, new Date().toLocaleDateString("en-IN"))
                            .replace(/\{Report Date\}/g, new Date().toLocaleDateString("en-IN"))
                            .replace(/\{Report Month\}/g, readableMonth);
                    };

                    const emailSubject = formatTemplate(schedule.subject_template || "Monthly Utility Report - {Location} - {Plant} - {Month}");
                    const rawBody = formatTemplate(schedule.body_template || "Please find attached the monthly UtilitySense report.");
                    const emailHtml = `
                        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e2e8f0; border-radius: 12px; max-width: 600px; margin: 0 auto;">
                            <div style="text-align: center; margin-bottom: 20px;">
                                <h2 style="color: #0284c7; margin: 0; font-size: 22px; font-weight: 800;">UTILITY SENSE</h2>
                                <p style="color: #64748b; font-size: 11px; margin: 4px 0 0 0;">Automated Operational Utility Report</p>
                            </div>
                            <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; margin-bottom: 20px; white-space: pre-line; font-size: 13px; color: #334155; line-height: 1.6;">
                                ${rawBody}
                            </div>
                            <div style="border-top: 1px solid #f1f5f9; padding-top: 12px; font-size: 11px; color: #94a3b8; text-align: center;">
                                Attached: <b>UtilitySense_${payload.locationLabel}_${payload.plantLabel}_${monthStr}.xlsx</b> (${payload.list.length} rows)
                            </div>
                        </div>
                    `;

                    const toList = (schedule.to_recipients || "").split(",").map(x => x.trim()).filter(Boolean);
                    const ccList = (schedule.cc_recipients || "").split(",").map(x => x.trim()).filter(Boolean);
                    const bccList = (schedule.bcc_recipients || "").split(",").map(x => x.trim()).filter(Boolean);

                    const res = await fetch("/api/send-report-email", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({
                            to: toList,
                            cc: ccList.length > 0 ? ccList : undefined,
                            bcc: bccList.length > 0 ? bccList : undefined,
                            subject: isManualTest ? `[TEST] ${emailSubject}` : emailSubject,
                            html: emailHtml,
                            attachments: [{
                                filename: `UtilitySense_${payload.locationLabel}_${payload.plantLabel}_${monthStr}.xlsx`,
                                content: base64Attachment,
                                encoding: 'base64'
                            }]
                        })
                    });

                    const resData = await res.json();
                    if (!res.ok) throw new Error(resData.error || resData.details || "Failed to dispatch email.");

                    const nowIso = new Date().toISOString();
                    const logEntry = {
                        schedule_id: schedule.id,
                        schedule_name: schedule.name,
                        location: emailLoc,
                        plant: emailPlant,
                        to_recipients: toList.join(", "),
                        cc_recipients: ccList.join(", "),
                        status: "SUCCESS",
                        executed_at: nowIso
                    };

                    // Persist log
                    setEmailScheduleLogs(prev => [logEntry, ...prev]);
                    await supabase.from('email_schedule_logs').insert(logEntry);

                    // Update schedule status
                    setEmailSchedules(prev => prev.map(s => s.id === schedule.id ? { ...s, last_run_at: nowIso, last_status: "SUCCESS" } : s));
                    await supabase.from('email_schedules').update({ last_run_at: nowIso, last_status: "SUCCESS" }).eq('id', schedule.id);

                    recordAuditLog({
                        action: isManualTest ? "TEST_EMAIL" : "EMAIL_DISPATCH",
                        module: "Email Automation",
                        recordId: schedule.id,
                        location: payload.locationLabel,
                        plant: payload.plantLabel,
                        newValue: { to: toList, cc: ccList, rows: payload.list.length, month: monthStr },
                        status: "SUCCESS"
                    });

                    setToast({
                        type: "success",
                        message: isManualTest
                            ? `Test report sent successfully to ${toList.length} recipient(s)!`
                            : `Schedule "${schedule.name}" executed successfully!`
                    });
                } catch (err) {
                    console.error("Email schedule run error:", err);
                    const nowIso = new Date().toISOString();
                    const errorLog = {
                        schedule_id: schedule.id,
                        schedule_name: schedule.name,
                        location: schedule.location,
                        plant: schedule.plant,
                        to_recipients: schedule.to_recipients,
                        cc_recipients: schedule.cc_recipients,
                        status: "FAILED",
                        error_message: err.message,
                        executed_at: nowIso
                    };

                    setEmailScheduleLogs(prev => [errorLog, ...prev]);
                    await supabase.from('email_schedule_logs').insert(errorLog);

                    setEmailSchedules(prev => prev.map(s => s.id === schedule.id ? { ...s, last_run_at: nowIso, last_status: "FAILED" } : s));
                    await supabase.from('email_schedules').update({ last_run_at: nowIso, last_status: "FAILED" }).eq('id', schedule.id);

                    recordAuditLog({
                        action: isManualTest ? "TEST_EMAIL" : "EMAIL_DISPATCH",
                        module: "Email Automation",
                        recordId: schedule.id,
                        location: schedule.location,
                        plant: schedule.plant,
                        newValue: { error: err.message },
                        status: "FAILED"
                    });

                    setToast({ type: "error", message: `Schedule execution failed: ${err.message}` });
                } finally {
                    setRunningScheduleId(null);
                }
            };

            // ----------------------------------------------------
            // SMART MASS EXCEL IMPORTER & BULK DATA UPLOADER (IT ADMIN ONLY)
            // ----------------------------------------------------
            const openMassUploadModal = () => {
                if (currentUser?.role !== "IT_ADMIN") {
                    setToast({ type: "error", message: "Unauthorized. Mass Uploader is restricted to IT Admin." });
                    return;
                }
                setMassUploadFileName("");
                setMassUploadSheets([]);
                setSelectedUploadSheet("");
                setRawUploadData([]);
                setDetectedHeaders([]);
                setAnalyzedImportRows([]);
                setImportResultSummary(null);
                setImportFilterTab("all");
                setLocationOverride("auto");
                setPlantOverride("auto");
                setIsMassUploadModalOpen(true);
            };

            // Date normalization helper
            const parseImportExcelDate = (val) => {
                if (!val && val !== 0) return null;
                if (val instanceof Date) {
                    if (isNaN(val.getTime())) return null;
                    return val.toISOString().split('T')[0];
                }
                if (typeof val === 'number') {
                    const dateObj = new Date(Math.round((val - 25569) * 86400 * 1000));
                    if (!isNaN(dateObj.getTime())) {
                        return dateObj.toISOString().split('T')[0];
                    }
                }
                if (typeof val === 'string') {
                    const str = val.trim();
                    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
                    if (/^\d{1,2}[-\/]\d{1,2}[-\/]\d{4}$/.test(str)) {
                        const parts = str.split(/[-\/]/);
                        const d = parts[0].padStart(2, '0');
                        const m = parts[1].padStart(2, '0');
                        const y = parts[2];
                        return `${y}-${m}-${d}`;
                    }
                    const dObj = new Date(str);
                    if (!isNaN(dObj.getTime())) {
                        const y = dObj.getFullYear();
                        const m = String(dObj.getMonth() + 1).padStart(2, '0');
                        const d = String(dObj.getDate()).padStart(2, '0');
                        return `${y}-${m}-${d}`;
                    }
                }
                return null;
            };

            // Smart auto column mapping recognizer
            const detectImportColumns = (headers) => {
                const mappings = {
                    date: "",
                    location: "",
                    plant: "",
                    electricity_closing: "",
                    electricity_opening: "",
                    solar: "",
                    diesel: "",
                    odu: "",
                    idu: "",
                    production_set: "",
                    waste_hazardous: "",
                    waste_non_hazardous: "",
                    waste_recycled: "",
                    water: "",
                    gas: "",
                    air: "",
                    remarks: "",
                    operator: ""
                };

                const clean = (h) => String(h || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

                headers.forEach(h => {
                    const c = clean(h);
                    const original = String(h).trim();

                    if (!mappings.date && (c.includes("date") || c === "dt" || c === "day" || c.includes("readingdate"))) {
                        mappings.date = original;
                    } else if (!mappings.location && (c.includes("location") || c === "site" || c === "city" || c === "unitlocation")) {
                        mappings.location = original;
                    } else if (!mappings.plant && (c === "plant" || c === "plantname" || c === "plantcode" || c === "unit" || c === "factory" || c === "division")) {
                        mappings.plant = original;
                    } else if (!mappings.electricity_opening && (c.includes("opening") || c.includes("prevreading") || c.includes("msebopening"))) {
                        mappings.electricity_opening = original;
                    } else if (!mappings.electricity_closing && (c.includes("electricity") || c.includes("mseb") || c.includes("power") || c.includes("energy") || c === "kwh" || c.includes("closing") || c.includes("grid"))) {
                        mappings.electricity_closing = original;
                    } else if (!mappings.solar && (c.includes("solar") || c.includes("pv") || c.includes("solargen") || c.includes("solarkwh"))) {
                        mappings.solar = original;
                    } else if (!mappings.diesel && (c.includes("diesel") || c.includes("dg") || c.includes("hsd") || c.includes("fuel"))) {
                        mappings.diesel = original;
                    } else if (!mappings.odu && (c.includes("odu") || c.includes("outdoor"))) {
                        mappings.odu = original;
                    } else if (!mappings.idu && (c.includes("idu") || c.includes("indoor"))) {
                        mappings.idu = original;
                    } else if (!mappings.production_set && (c.includes("production") || c === "sets" || c.includes("prodqty"))) {
                        mappings.production_set = original;
                    } else if (!mappings.waste_hazardous && (c.includes("haz") && !c.includes("non"))) {
                        mappings.waste_hazardous = original;
                    } else if (!mappings.waste_non_hazardous && (c.includes("nonhaz") || c.includes("nonhazardous"))) {
                        mappings.waste_non_hazardous = original;
                    } else if (!mappings.waste_recycled && (c.includes("recycled") || c.includes("recycle"))) {
                        mappings.waste_recycled = original;
                    } else if (!mappings.water && (c.includes("water") || c.includes("kl") || c.includes("ro"))) {
                        mappings.water = original;
                    } else if (!mappings.gas && (c.includes("gas") || c.includes("png") || c.includes("lpg") || c.includes("scm"))) {
                        mappings.gas = original;
                    } else if (!mappings.air && (c.includes("air") || c.includes("compressor") || c.includes("cfm"))) {
                        mappings.air = original;
                    } else if (!mappings.remarks && (c.includes("remark") || c.includes("comment") || c.includes("note"))) {
                        mappings.remarks = original;
                    } else if (!mappings.operator && (c.includes("operator") || c.includes("loggedby") || c.includes("user"))) {
                        mappings.operator = original;
                    }
                });

                return mappings;
            };

            // Row analysis & validation engine
            const analyzeImportRows = (rows, currentMappings, locOverride, pltOverride) => {
                if (!rows || rows.length === 0) return [];

                const analyzed = [];
                const seenBatchKeys = new Set(); // to detect internal duplicates in the file

                rows.forEach((row, idx) => {
                    const errors = [];
                    const warnings = [];

                    // 1. Date resolution
                    const rawDate = currentMappings.date ? row[currentMappings.date] : (row.Date || row.date || row.DATE);
                    const resolvedDate = parseImportExcelDate(rawDate);
                    if (!resolvedDate) {
                        errors.push(`Row ${idx + 2}: Unable to recognize a valid Date format (got "${rawDate || ''}").`);
                    }

                    // 2. Location resolution
                    let resolvedLoc = "";
                    if (locOverride && locOverride !== "auto") {
                        resolvedLoc = locOverride.toUpperCase();
                    } else if (currentMappings.location && row[currentMappings.location]) {
                        resolvedLoc = String(row[currentMappings.location]).trim().toUpperCase();
                    }

                    // Validate location exists
                    const knownLocations = Array.from(new Set(plants.map(p => p.location.toUpperCase())));
                    if (resolvedLoc && !knownLocations.includes(resolvedLoc)) {
                        // Fuzzy match location
                        const matchedLoc = knownLocations.find(l => l.includes(resolvedLoc) || resolvedLoc.includes(l));
                        if (matchedLoc) {
                            resolvedLoc = matchedLoc;
                        } else {
                            errors.push(`Row ${idx + 2}: Unknown Location "${resolvedLoc}". Not present in Location Master.`);
                        }
                    }

                    // 3. Plant resolution
                    let resolvedPlant = "";
                    if (pltOverride && pltOverride !== "auto") {
                        resolvedPlant = pltOverride;
                    } else if (currentMappings.plant && row[currentMappings.plant]) {
                        const rawPlantStr = String(row[currentMappings.plant]).trim().toUpperCase();
                        // Match against plant_code or plant_name
                        const matchedPlant = plants.find(p => 
                            p.plant_code.toUpperCase() === rawPlantStr || 
                            p.plant_name.toUpperCase() === rawPlantStr ||
                            rawPlantStr.includes(p.plant_code.toUpperCase()) ||
                            (p.plant_display_name && p.plant_display_name.toUpperCase() === rawPlantStr)
                        );
                        if (matchedPlant) {
                            resolvedPlant = matchedPlant.plant_code;
                            if (!resolvedLoc) resolvedLoc = matchedPlant.location.toUpperCase();
                        } else {
                            errors.push(`Row ${idx + 2}: Unknown Plant "${rawPlantStr}". Not found in Plant Master.`);
                        }
                    }

                    // Default to first plant if location chosen and only 1 plant exists for that location
                    if (resolvedLoc && !resolvedPlant) {
                        const locPlants = plants.filter(p => p.location.toUpperCase() === resolvedLoc.toUpperCase());
                        if (locPlants.length === 1) {
                            resolvedPlant = locPlants[0].plant_code;
                        } else if (locPlants.length > 1) {
                            errors.push(`Row ${idx + 2}: Location "${resolvedLoc}" has multiple plants. Please select a specific Plant.`);
                        } else {
                            errors.push(`Row ${idx + 2}: Missing Plant information.`);
                        }
                    }

                    if (!resolvedLoc) {
                        errors.push(`Row ${idx + 2}: Missing Location information.`);
                    }
                    if (!resolvedPlant) {
                        errors.push(`Row ${idx + 2}: Missing Plant code.`);
                    }

                    // 4. Duplicate Check (Existing database record protection)
                    let isDuplicate = false;
                    let existingRecord = null;
                    if (resolvedDate && resolvedPlant) {
                        const batchKey = `${resolvedPlant}_${resolvedDate}`;
                        if (seenBatchKeys.has(batchKey)) {
                            isDuplicate = true;
                            warnings.push(`Duplicate date ${resolvedDate} for plant ${resolvedPlant} inside this Excel file. Skipping repeated row.`);
                        } else {
                            seenBatchKeys.add(batchKey);
                            const existing = dailyEntries.find(e => e.plant === resolvedPlant && e.date === resolvedDate);
                            if (existing) {
                                isDuplicate = true;
                                existingRecord = existing;
                                warnings.push(`Existing record found in database for ${resolvedDate} (${resolvedPlant}). Protected from overwrite.`);
                            }
                        }
                    }

                    // 5. Utility Readings extraction
                    const cleanNum = (val) => {
                        if (val === null || val === undefined || val === "" || val === "-") return 0;
                        const n = Number(String(val).replace(/,/g, '').trim());
                        return isNaN(n) ? 0 : n;
                    };

                    const electClosingRaw = currentMappings.electricity_closing ? row[currentMappings.electricity_closing] : 0;
                    const electOpeningRaw = currentMappings.electricity_opening ? row[currentMappings.electricity_opening] : 0;
                    const solarRaw = currentMappings.solar ? row[currentMappings.solar] : 0;
                    const dieselRaw = currentMappings.diesel ? row[currentMappings.diesel] : 0;
                    const oduRaw = currentMappings.odu ? row[currentMappings.odu] : 0;
                    const iduRaw = currentMappings.idu ? row[currentMappings.idu] : 0;
                    const prodSetRaw = currentMappings.production_set ? row[currentMappings.production_set] : 0;
                    const wasteHazRaw = currentMappings.waste_hazardous ? row[currentMappings.waste_hazardous] : 0;
                    const wasteNonHazRaw = currentMappings.waste_non_hazardous ? row[currentMappings.waste_non_hazardous] : 0;
                    const wasteRecycledRaw = currentMappings.waste_recycled ? row[currentMappings.waste_recycled] : 0;
                    const remarksRaw = currentMappings.remarks ? String(row[currentMappings.remarks] || "") : "";
                    const operatorRaw = currentMappings.operator ? String(row[currentMappings.operator] || "") : (currentUser?.name || "IT Admin");

                    const electClosing = cleanNum(electClosingRaw);
                    const electOpening = cleanNum(electOpeningRaw);
                    const solarGen = cleanNum(solarRaw);
                    const dieselUsed = cleanNum(dieselRaw);
                    const odu = cleanNum(oduRaw);
                    const idu = cleanNum(iduRaw);
                    const sets = cleanNum(prodSetRaw) || CalculationEngine.calculateProductionSets(odu, idu);

                    // 6. Calculations
                    let mf = 1;
                    let electTariff = 0;
                    let solarTariff = 0;
                    let dieselTariff = 0;

                    if (resolvedPlant && resolvedDate) {
                        mf = resolveMultiplyFactor(multiplyFactors, resolvedPlant, resolvedLoc, resolvedDate);
                        electTariff = resolveTariff(tariffs, "electricity", resolvedPlant, resolvedLoc, resolvedDate);
                        solarTariff = resolveTariff(tariffs, "solar", resolvedPlant, resolvedLoc, resolvedDate) || electTariff;
                        dieselTariff = resolveTariff(tariffs, "diesel", resolvedPlant, resolvedLoc, resolvedDate);
                    }

                    const unitsDiff = electClosing > electOpening ? (electClosing - electOpening) : electClosing;
                    const msebUnits = CalculationEngine.calculateMSEBUnits(unitsDiff, mf);
                    const electricityCost = CalculationEngine.calculateElectricityCost(msebUnits, electTariff);
                    const solarCost = CalculationEngine.calculateSolarCost(solarGen, solarTariff);
                    const dieselCost = CalculationEngine.calculateDieselCost(dieselUsed, dieselTariff);
                    const totalCost = CalculationEngine.calculateTotalCost(electricityCost, solarCost, dieselCost);
                    const costPerSet = CalculationEngine.calculateCostPerSet(totalCost, sets);
                    const sec = sets > 0 ? msebUnits / sets : 0;

                    let status = "VALID";
                    if (errors.length > 0) {
                        status = "INVALID";
                    } else if (isDuplicate) {
                        status = "DUPLICATE";
                    }

                    analyzed.push({
                        rowIndex: idx + 2,
                        raw: row,
                        date: resolvedDate,
                        location: resolvedLoc,
                        plant: resolvedPlant,
                        electricity_opening: electOpening,
                        electricity_closing: electClosing,
                        electricity_consumption: msebUnits,
                        electricity_cost: electricityCost,
                        solar_generated: solarGen,
                        solar_cost: solarCost,
                        diesel_used: dieselUsed,
                        diesel_cost: dieselCost,
                        total_cost: totalCost,
                        odu: odu,
                        idu: idu,
                        production_set: sets,
                        cost_per_set: costPerSet,
                        sec: sec,
                        waste_hazardous: cleanNum(wasteHazRaw),
                        waste_non_hazardous: cleanNum(wasteNonHazRaw),
                        waste_recycled: cleanNum(wasteRecycledRaw),
                        remarks: remarksRaw,
                        operator_name: operatorRaw,
                        status,
                        errors,
                        warnings,
                        existingRecord
                    });
                });

                return analyzed;
            };

            // File Upload & Sheet Extraction
            const handleMassExcelFileUpload = async (e) => {
                const file = e.target.files[0];
                if (!file) return;

                setIsAnalyzingExcel(true);
                setMassUploadFileName(file.name);
                setImportResultSummary(null);

                try {
                    const data = await file.arrayBuffer();
                    const wb = XLSX.read(data, { type: 'array', cellDates: true });
                    
                    if (!wb.SheetNames || wb.SheetNames.length === 0) {
                        throw new Error("No sheets found in uploaded Excel file.");
                    }

                    setMassUploadSheets(wb.SheetNames);
                    const firstSheet = wb.SheetNames[0];
                    setSelectedUploadSheet(firstSheet);

                    const ws = wb.Sheets[firstSheet];
                    const jsonRows = XLSX.utils.sheet_to_json(ws, { defval: "" });

                    if (!jsonRows || jsonRows.length === 0) {
                        throw new Error(`Sheet "${firstSheet}" contains no data rows.`);
                    }

                    setRawUploadData(jsonRows);
                    const headers = Object.keys(jsonRows[0] || {});
                    setDetectedHeaders(headers);

                    const autoMappings = detectImportColumns(headers);
                    setColumnMappings(autoMappings);

                    const analyzed = analyzeImportRows(jsonRows, autoMappings, locationOverride, plantOverride);
                    setAnalyzedImportRows(analyzed);

                    setToast({
                        type: "success",
                        message: `Analyzed ${jsonRows.length} rows from "${file.name}". ${analyzed.filter(r => r.status === 'VALID').length} valid rows ready.`
                    });
                } catch (err) {
                    console.error("Excel analysis error:", err);
                    setToast({ type: "error", message: `Excel read failed: ${err.message}` });
                    setAnalyzedImportRows([]);
                } finally {
                    setIsAnalyzingExcel(false);
                    e.target.value = ""; // reset file input
                }
            };

            // Re-analyze when mappings or override dropdowns change
            const handleMappingChange = (field, newCol) => {
                const updated = { ...columnMappings, [field]: newCol };
                setColumnMappings(updated);
                const reAnalyzed = analyzeImportRows(rawUploadData, updated, locationOverride, plantOverride);
                setAnalyzedImportRows(reAnalyzed);
            };

            const handleLocationOverrideChange = (newLoc) => {
                setLocationOverride(newLoc);
                let newPlt = plantOverride;
                if (newLoc !== "auto") {
                    const matching = plants.filter(p => p.location.toUpperCase() === newLoc.toUpperCase());
                    newPlt = matching.length === 1 ? matching[0].plant_code : "auto";
                    setPlantOverride(newPlt);
                }
                const reAnalyzed = analyzeImportRows(rawUploadData, columnMappings, newLoc, newPlt);
                setAnalyzedImportRows(reAnalyzed);
            };

            const handlePlantOverrideChange = (newPlt) => {
                setPlantOverride(newPlt);
                const reAnalyzed = analyzeImportRows(rawUploadData, columnMappings, locationOverride, newPlt);
                setAnalyzedImportRows(reAnalyzed);
            };

            // Execute Safe Transactional Import
            const handleExecuteMassImport = async () => {
                if (currentUser?.role !== "IT_ADMIN") {
                    setToast({ type: "error", message: "Unauthorized. IT Admin access required." });
                    return;
                }

                const validRows = analyzedImportRows.filter(r => r.status === "VALID");
                if (validRows.length === 0) {
                    setToast({ type: "warning", message: "No valid rows to import. Please check errors or mapping." });
                    return;
                }

                setIsImportingData(true);
                try {
                    const payloads = validRows.map(r => ({
                        date: r.date,
                        plant: r.plant,
                        location: r.location,
                        department: "PROD",
                        shift: "Shift A",
                        operator_name: r.operator_name || currentUser?.name || "IT Admin",
                        electricity_opening: r.electricity_opening || 0,
                        electricity_closing: r.electricity_closing || 0,
                        electricity_consumption: r.electricity_consumption || 0,
                        electricity_cost: r.electricity_cost || 0,
                        solar_opening: 0,
                        solar_closing: r.solar_generated || 0,
                        solar_generated: r.solar_generated || 0,
                        solar_utilized: r.solar_generated || 0,
                        solar_cost: r.solar_cost || 0,
                        solar_utilization_pct: 100,
                        diesel_used: r.diesel_used || 0,
                        diesel_cost: r.diesel_cost || 0,
                        total_cost: r.total_cost || 0,
                        odu: r.odu || 0,
                        idu: r.idu || 0,
                        production_set: r.production_set || 0,
                        production_qty: r.production_set || 0,
                        production_unit: "Sets",
                        cost_per_set: r.cost_per_set || 0,
                        sec: r.sec || 0,
                        waste_hazardous: r.waste_hazardous || 0,
                        waste_non_hazardous: r.waste_non_hazardous || 0,
                        waste_recycled: r.waste_recycled || 0,
                        remarks: r.remarks ? `[Mass Import] ${r.remarks}` : `[Mass Import: ${massUploadFileName}]`
                    }));

                    // Insert into Supabase in chunks of 50
                    const chunkSize = 50;
                    const insertedRows = [];
                    for (let i = 0; i < payloads.length; i += chunkSize) {
                        const chunk = payloads.slice(i, i + chunkSize);
                        const { data, error } = await supabase.from('daily_entries').insert(chunk).select();
                        if (error) {
                            console.warn("Chunk insert warning:", error.message);
                            // Even if Supabase upsert/insert has issues, add to local state with client-generated IDs
                            chunk.forEach((item, cIdx) => {
                                insertedRows.push({ ...item, id: `local_imp_${Date.now()}_${i + cIdx}` });
                            });
                        } else if (data) {
                            insertedRows.push(...data);
                        }
                    }

                    // Update local dailyEntries state
                    setDailyEntries(prev => [...insertedRows, ...prev]);

                    const duplicateCount = analyzedImportRows.filter(r => r.status === "DUPLICATE").length;
                    const invalidCount = analyzedImportRows.filter(r => r.status === "INVALID").length;

                    const summary = {
                        fileName: massUploadFileName,
                        totalRows: analyzedImportRows.length,
                        importedCount: validRows.length,
                        skippedDuplicates: duplicateCount,
                        invalidCount: invalidCount,
                        timestamp: new Date().toISOString()
                    };

                    setImportResultSummary(summary);

                    // Record Audit Log
                    recordAuditLog({
                        action: "MASS_IMPORT",
                        module: "Daily Operations",
                        newValue: summary,
                        status: "SUCCESS"
                    });

                    // Save into import_history
                    const historyRecord = {
                        import_id: `IMP_${Date.now()}`,
                        file_name: massUploadFileName,
                        uploaded_by: currentUser?.email || "admin@pgel.in",
                        user_id: String(currentUser?.id || ""),
                        location: locationOverride !== "auto" ? locationOverride : "Multi-Location",
                        plant: plantOverride !== "auto" ? plantOverride : "Multi-Plant",
                        total_rows: analyzedImportRows.length,
                        imported_rows: validRows.length,
                        skipped_rows: duplicateCount,
                        failed_rows: invalidCount,
                        status: invalidCount === 0 ? "COMPLETED" : "PARTIAL",
                        details: summary,
                        created_at: new Date().toISOString()
                    };

                    setImportHistory(prev => [historyRecord, ...prev]);
                    supabase.from('import_history').insert([historyRecord]).then(({ error }) => {
                        if (error) console.info("import_history optional log:", error.message);
                    });

                    setToast({
                        type: "success",
                        message: `Successfully imported ${validRows.length} rows! (${duplicateCount} existing records protected).`
                    });
                } catch (err) {
                    console.error("Mass import execution error:", err);
                    setToast({ type: "error", message: `Import failed: ${err.message}` });
                } finally {
                    setIsImportingData(false);
                }
            };

            // ----------------------------------------------------
            // AUTHENTICATION LOGIC
            // ----------------------------------------------------

            // Wraps google.script.run with one automatic silent retry — Apps Script's
            // hidden-iframe transport occasionally drops a request on a flaky connection
            // (shows up in DevTools as net::ERR_INTERNET_DISCONNECTED for the callback
            // resource). A short retry recovers from that without bothering the user.
            const callServerWithRetry = (fnName, args, onSuccess, onError, attempt = 1) => {
                const maxAttempts = 2;
                const runner = google.script.run
                    .withSuccessHandler(onSuccess)
                    .withFailureHandler((err) => {
                        if (attempt < maxAttempts) {
                            setTimeout(() => callServerWithRetry(fnName, args, onSuccess, onError, attempt + 1), 1500);
                        } else {
                            onError(err);
                        }
                    });
                runner[fnName].apply(runner, args);
            };

            const friendlyServerError = (err, fallback) => {
                const raw = (err && err.message) ? String(err.message) : "";
                if (!raw || /disconnected|network|fetch|failed to fetch/i.test(raw)) {
                    return "Couldn't reach the server. Please check your internet connection and try again.";
                }
                return raw || fallback;
            };

                        const handleSendOTP = async (e) => {
                e.preventDefault();
                setLoginError("");
                setLoginMessage("");

                const emailVal = loginEmail.trim().toLowerCase();
                if (!emailVal) {
                    setLoginError("Email address is required.");
                    return;
                }
                if (!emailVal.endsWith("@pgel.in")) {
                    setLoginError("Invalid email. Only corporate email addresses ending with @pgel.in are authorized.");
                    return;
                }

                setLoginLoading(true);

                try {
                    // Check if email is registered in users table
                    const { data: userRec, error: userErr } = await supabase
                        .from('users')
                        .select('*')
                        .eq('email', emailVal)
                        .eq('status', 'Active')
                        .single();

                    if (userErr || !userRec) {
                        throw new Error("Unauthorized email. Please ask your IT Admin to register this address.");
                    }

                    // Generate a 6-digit OTP
                    const randomOtp = Math.floor(100000 + Math.random() * 900000).toString();

                    // Insert to otp_logs
                    const { error: otpErr } = await supabase
                        .from('otp_logs')
                        .insert({
                            email: emailVal,
                            otp: randomOtp,
                            created_at: new Date().toISOString(),
                            status: 'Active'
                        });

                    if (otpErr) throw otpErr;

                    // Send email using Vercel serverless API
                    const res = await fetch("/api/send-otp", {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json"
                        },
                        body: JSON.stringify({
                            email: emailVal,
                            otp: randomOtp
                        })
                    });
                    if (!res.ok) {
                        const resData = await res.json();
                        throw new Error(resData.error || "Failed to dispatch OTP email via SMTP.");
                    }

                    setOtpSent(true);
                    setLoginMessage("A 6-digit verification code has been sent to your corporate email!");
                } catch (err) {
                    setLoginError(err.message || "Failed to request OTP.");
                } finally {
                    setLoginLoading(false);
                }
            };

                        const handleVerifyOTP = async (e) => {
                e.preventDefault();
                setLoginError("");
                setLoginMessage("");

                const emailVal = loginEmail.trim().toLowerCase();
                const otpVal = loginOtp.trim();

                if (!otpVal || otpVal.length !== 6) {
                    setLoginError("Please enter a valid 6-digit OTP code.");
                    return;
                }

                setLoginLoading(true);

                try {
                    // Fetch latest active OTP from otp_logs
                    const { data: otpRecs, error: otpErr } = await supabase
                        .from('otp_logs')
                        .select('*')
                        .eq('email', emailVal)
                        .order('created_at', { ascending: false })
                        .limit(1);

                    if (otpErr || !otpRecs || otpRecs.length === 0) {
                        throw new Error("No active OTP request found for this email.");
                    }

                    if (otpRecs[0].otp !== otpVal) {
                        throw new Error("Invalid OTP code. Please try again.");
                    }

                    // Get user profile
                    const { data: userRec, error: userErr } = await supabase
                        .from('users')
                        .select('*')
                        .eq('email', emailVal)
                        .eq('status', 'Active')
                        .single();

                    if (userErr || !userRec) {
                        throw new Error("Failed to load user profile.");
                    }

                    // Save session details
                    const sessionUser = {
                        ...userRec,
                        loggedAt: Date.now()
                    };
                    localStorage.setItem("ep_session", JSON.stringify({
                        user: userRec,
                        loggedAt: Date.now()
                    }));
                    setCurrentUser(userRec);
                    recordAuditLog({ action: "LOGIN", module: "Auth", status: "SUCCESS" });
                    setToast({ type: "success", message: `Welcome back, ${userRec.name}!` });
                } catch (err) {
                    setLoginError(err.message || "OTP verification failed.");
                    recordAuditLog({ action: "FAILED_LOGIN", module: "Auth", status: "FAILED", newValue: { email: emailVal, error: err.message } });
                } finally {
                    setLoginLoading(false);
                }
            };

            const handleLogout = (msg = "Logged out successfully.") => {
                recordAuditLog({ action: "LOGOUT", module: "Auth", status: "SUCCESS" });
                localStorage.removeItem("ep_session");
                setCurrentUser(null);
                setOtpSent(false);
                setLoginEmail("");
                setLoginOtp("");
                setLoginMessage("");
                setLoginError("");
                setActiveTab("dashboard");
                if (msg) setToast({ type: "success", message: msg });
            };

            // ----------------------------------------------------
            // DAILY DATA ENTRY MUTATIONS
            // ----------------------------------------------------

            const openDailyForm = (entry = null) => {
                setWasteSectionOpen(false);
                if (entry) {
                    setEditingRecord(entry);
                    const solGen = Number(entry.solar_generated) || 0;
                    const solCloseRaw = Number(entry.solar_closing);
                    const solarDaily = Number.isFinite(solCloseRaw) && solCloseRaw > 0 ? solCloseRaw : solGen;
                    const prevSolarEntry = dailyEntries
                        .filter(e => e.plant === entry.plant && e.date < entry.date && e.id !== entry.id)
                        .sort((a, b) => b.date.localeCompare(a.date))[0];
                    const prevSolar = prevSolarEntry ? (Number(prevSolarEntry.solar_generated) || 0) : 0;
                    const solarCost = CalculationEngine.calculateSolarCost(solarDaily, resolveTariff(tariffs, "solar", entry.plant, entry.location, entry.date) || resolveTariff(tariffs, "electricity", entry.plant, entry.location, entry.date));
                    const dieselCost = CalculationEngine.calculateDieselCost(Number(entry.diesel_used) || 0, resolveTariff(tariffs, "diesel", entry.plant, entry.location, entry.date));
                    const electCost = Number(entry.electricity_cost) || 0;
                    setEntryFormValues({
                        ...entry,
                        meter_changed: entry.meter_changed === true || String(entry.meter_changed) === "true",
                        custom_difference: entry.custom_difference || "",
                        solar_opening: prevSolar,
                        solar_closing: solarDaily || "",
                        solar_generated: solarDaily,
                        solar_utilized: solarDaily,
                        solar_cost: solarCost,
                        diesel_cost: dieselCost,
                        total_cost: CalculationEngine.calculateTotalCost(electCost, solarCost, dieselCost),
                    });
                    // Auto-expand the waste section if the record being edited already has waste data
                    if (Number(entry.waste_hazardous) || Number(entry.waste_non_hazardous) || Number(entry.waste_recycled)) {
                        setWasteSectionOpen(true);
                    }
                } else {
                    setEditingRecord(null);

                    const pDef = allowedPlants[0]?.plant_code || "NGM";
                    const lDef = allowedPlants[0]?.location || "NASHIK";

                    setEntryFormValues({
                        date: new Date().toISOString().split('T')[0],
                        plant: pDef,
                        location: lDef,
                        department: "PROD",
                        shift: "Shift A",
                        operator_name: currentUser?.name || "Operator",
                        electricity_opening: 0,
                        electricity_closing: "",
                        meter_changed: false,
                        custom_difference: "",
                        electricity_consumption: 0,
                        electricity_cost: 0,
                        solar_opening: 0,
                        solar_closing: "",
                        solar_generated: "",
                        solar_utilized: 0,
                        solar_cost: 0,
                        solar_utilization_pct: 100,
                        diesel_used: "",
                        diesel_cost: 0,
                        total_cost: 0,
                        odu: "",
                        idu: "",
                        production_set: 0,
                        cost_per_set: 0,
                        waste_hazardous: "",
                        waste_non_hazardous: "",
                        waste_recycled: "",
                        remarks: ""
                    });
                }
                setIsFormOpen(true);
            };

            // Load production Excel cache from sessionStorage on mount
            useEffect(() => {
                try {
                    const cached = sessionStorage.getItem("ep_production_excel");
                    const cachedName = sessionStorage.getItem("ep_production_excel_name");
                    if (cached) {
                        setProductionExcelData(JSON.parse(cached));
                        setExcelFileName(cachedName || "excel_file");
                    }
                } catch (e) {
                    console.error("Failed to load cached excel data:", e);
                }
            }, []);

             const findExcelDataByDateAndPlant = (dateStr, plantCode) => {
                if (!productionExcelData) return null;

                const matchingKey = Object.keys(productionExcelData).find(key => 
                    key.includes(plantCode.toUpperCase())
                );
                if (!matchingKey) {
                    console.warn("No Excel sheet found matching plant: " + plantCode);
                    return null;
                }

                const rows = productionExcelData[matchingKey];
                
                const formatDate = (val) => {
                    if (val instanceof Date) {
                        return val.toISOString().split('T')[0];
                    }
                    if (typeof val === 'string') {
                        const cleaned = val.trim();
                        if (/^\d{4}-\d{2}-\d{2}$/.test(cleaned)) return cleaned;
                        if (/^\d{2}-\d{2}-\d{4}$/.test(cleaned)) {
                            const [d, m, y] = cleaned.split('-');
                            return y + "-" + m + "-" + d;
                        }
                    }
                    if (typeof val === 'number') {
                        // Excel serial date number conversion (1900 format)
                        const dateObj = new Date((val - 25569) * 86400 * 1000);
                        if (!isNaN(dateObj.getTime())) {
                            return dateObj.toISOString().split('T')[0];
                        }
                    }
                    return null;
                };

                const targetDateFormatted = formatDate(new Date(dateStr));

                // Find matching row (skipping first few rows for header lines)
                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || row.length === 0) continue;

                    const rowDateVal = row[0]; // Column A is Date
                    const formattedRowDate = formatDate(rowDateVal);

                    if (formattedRowDate && formattedRowDate === targetDateFormatted) {
                        const rawClosing = row[2];  // Column C: Daily Reading kWh UNITS
                        const rawSolar = row[5];    // Column F: SOLAR UNITS
                        const rawDiesel = row[11];  // Column L: CONSUMED DIESEL (L)
                        const rawOdu = row[14];     // Column O: ODU Count
                        const rawIdu = row[15];     // Column P: IDU Count

                        const cleanVal = (val) => {
                            if (val === undefined || val === null || val === "" || val === "-") return "";
                            const parsed = Number(val);
                            return isNaN(parsed) ? "" : parsed.toString();
                        };

                        return {
                            electricity_closing: cleanVal(rawClosing),
                            solar_generated: cleanVal(rawSolar),
                            diesel_used: cleanVal(rawDiesel),
                            odu: cleanVal(rawOdu),
                            idu: cleanVal(rawIdu)
                        };
                    }
                }

                return null;
            };

            const handleProductionExcelUpload = (e) => {
                const file = e.target.files[0];
                if (!file) return;
                setExcelFileName(file.name);

                const reader = new FileReader();
                reader.onload = (evt) => {
                    try {
                        const data = new Uint8Array(evt.target.result);
                        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                        
                        const sheetsMap = {};
                        workbook.SheetNames.forEach(sheetName => {
                            const sheet = workbook.Sheets[sheetName];
                            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
                            sheetsMap[sheetName.toUpperCase()] = rows;
                        });

                        setProductionExcelData(sheetsMap);
                        try {
                            sessionStorage.setItem("ep_production_excel", JSON.stringify(sheetsMap));
                            sessionStorage.setItem("ep_production_excel_name", file.name);
                        } catch (errStore) {
                            console.warn("Session storage limit reached:", errStore);
                        }
                        setToast({ type: "success", message: "Synced successfully from " + file.name + "!" });
                    } catch (err) {
                        console.error("Excel sync failed:", err);
                        setToast({ type: "error", message: "Failed to parse Excel file. Ensure it is a valid format." });
                    }
                };
                reader.readAsArrayBuffer(file);
            };

            // React effect: auto lookup previous day's closing reading and Excel production details if cached
            useEffect(() => {
                if (!isFormOpen) return;
                const { date, plant } = entryFormValues;
                if (!date || !plant) return;

                // Lock changes and do not trigger if we are editing and plant/date has not changed
                const isEditingSame = editingRecord && editingRecord.date === date && editingRecord.plant === plant;
                if (isEditingSame) return;

                const sortedPrev = dailyEntries
                    .filter(e => e.plant === plant && e.date < date)
                    .sort((a, b) => b.date.localeCompare(a.date));
                const prevReading = sortedPrev.length > 0 ? Number(sortedPrev[0].electricity_closing) || 0 : 0;
                const prevSolar = sortedPrev.length > 0 ? Number(sortedPrev[0].solar_generated) || 0 : 0;

                if (productionExcelData) {
                    const match = findExcelDataByDateAndPlant(date, plant);
                    if (match) {
                        setIsFetchingExcelData(true);
                        const timer = setTimeout(() => {
                            updateFormCalculations({
                                electricity_opening: prevReading,
                                electricity_closing: match.electricity_closing,
                                solar_opening: prevSolar,
                                solar_closing: match.solar_generated,
                                diesel_used: match.diesel_used,
                                odu: match.odu,
                                idu: match.idu
                            });
                            setIsFetchingExcelData(false);
                            setToast({ type: "success", message: "All inputs successfully pre-filled from Excel!" });
                        }, 1000);
                        return () => clearTimeout(timer);
                    } else {
                        // No cached Excel match for this date/plant — only set opening reading.
                        // Do NOT clear odu/idu here; Supabase production_summary effect fills those.
                        updateFormCalculations({
                            electricity_opening: prevReading,
                            electricity_closing: "",
                            solar_opening: prevSolar,
                            solar_closing: "",
                            diesel_used: "",
                        });
                    }
                } else {
                    updateFormCalculations({ electricity_opening: prevReading, solar_opening: prevSolar });
                }

            }, [entryFormValues.date, entryFormValues.plant, isFormOpen, productionExcelData]);

            // Reusable Real-time calculation updater
            const updateFormCalculations = (changedFields = {}) => {
                setEntryFormValues(prev => {
                    const next = { ...prev, ...changedFields };
                    const plantObj = plants.find((p) => p.plant_code === next.plant);
                    const loc = next.location || plantObj?.location || "";
                    const mf = resolveMultiplyFactor(multiplyFactors, next.plant, loc, next.date);
                    const electTariff = resolveTariff(tariffs, "electricity", next.plant, loc, next.date);
                    const solarRate = resolveTariff(tariffs, "solar", next.plant, loc, next.date) || electTariff;
                    const dieselRate = resolveTariff(tariffs, "diesel", next.plant, loc, next.date);

                    const opening = Number(next.electricity_opening) || 0;
                    const closing = Number(next.electricity_closing) || 0;

                    // Units difference calculation (supports custom override)
                    let diff = 0;
                    if (next.meter_changed === true || String(next.meter_changed) === "true") {
                        diff = Number(next.custom_difference) || 0;
                    } else {
                        diff = Math.max(0, closing - opening);
                    }

                    const msebUnits = CalculationEngine.calculateMSEBUnits(diff, mf);
                    next.electricity_consumption = msebUnits;
                    next.electricity_cost = CalculationEngine.calculateElectricityCost(msebUnits, electTariff);

                    // Solar: Daily Reading × rate (previous is display-only, not subtracted)
                    next.solar_generated = Math.max(0, Number(next.solar_closing) || 0);
                    next.solar_utilized = next.solar_generated;
                    next.solar_utilization_pct = 100;
                    next.solar_cost = CalculationEngine.calculateSolarCost(next.solar_generated, solarRate);

                    next.diesel_used = Number(next.diesel_used) || 0;
                    next.diesel_cost = CalculationEngine.calculateDieselCost(next.diesel_used, dieselRate);

                    const totalCost = CalculationEngine.calculateTotalCost(next.electricity_cost, next.solar_cost, next.diesel_cost);
                    next.total_cost = totalCost;

                    const ODU = Number(next.odu) || 0;
                    const IDU = Number(next.idu) || 0;

                    // Trigger Production Sets Logic
                    const prodSets = CalculationEngine.calculateProductionSets(ODU, IDU);
                    next.production_set = prodSets;
                    next.production_qty = prodSets;
                    next.production_unit = "Sets";

                    next.cost_per_set = CalculationEngine.calculateCostPerSet(totalCost, prodSets);
                    next.sec = prodSets > 0 ? msebUnits / prodSets : 0;

                    // Waste fields (optional, collapsed by default) — normalize to numbers
                    next.waste_hazardous = next.waste_hazardous === "" ? "" : (Number(next.waste_hazardous) || 0);
                    next.waste_non_hazardous = next.waste_non_hazardous === "" ? "" : (Number(next.waste_non_hazardous) || 0);
                    next.waste_recycled = next.waste_recycled === "" ? "" : (Number(next.waste_recycled) || 0);

                    return next;
                });
            };

            // Auto lookup previous day's closing reading
            useEffect(() => {
                if (!isFormOpen) return;
                const { date, plant } = entryFormValues;
                if (!date || !plant) return;

                // Lock changes and do not trigger if we are editing and plant/date has not changed
                const isEditingSame = editingRecord && editingRecord.date === date && editingRecord.plant === plant;
                if (isEditingSame) return;

                const sortedPrev = dailyEntries
                    .filter(e => e.plant === plant && e.date < date)
                    .sort((a, b) => b.date.localeCompare(a.date));
                const prevReading = sortedPrev.length > 0 ? Number(sortedPrev[0].electricity_closing) || 0 : 0;
                const prevSolar = sortedPrev.length > 0 ? Number(sortedPrev[0].solar_generated) || 0 : 0;

                updateFormCalculations({ electricity_opening: prevReading, solar_opening: prevSolar });

            }, [entryFormValues.date, entryFormValues.plant, isFormOpen]);

            // Auto lookup ODU and IDU totals from Supabase production_summary SQL View
            useEffect(() => {
                if (!isFormOpen) return;
                const { date, plant, location } = entryFormValues;
                if (!date || !plant) return;

                let isMounted = true;

                const fetchSupabaseProductionTotals = async () => {
                    try {
                        // plant in entryFormValues is plant_code (e.g. '4010')
                        // production_summary stores plant as plant_display_name (e.g. 'NGM')
                        // Need to resolve plant_display_name from plants array
                        const plantObj = plants.find(p => p.plant_code === plant);
                        const plantDisplayName = plantObj ? (plantObj.plant_display_name || plant) : plant;

                        let query = supabase
                            .from('production_summary')
                            .select('odu_total, idu_total')
                            .eq('date', date)
                            .eq('plant', plantDisplayName);

                        const { data: rows, error } = await query;

                        if (!isMounted) return;

                        if (!error && rows && rows.length > 0) {
                            // Sum in case of multiple rows, then divide to avoid double counting
                            const oduSum = rows.reduce((s, r) => s + (Number(r.odu_total) || 0), 0);
                            const iduSum = rows.reduce((s, r) => s + (Number(r.idu_total) || 0), 0);
                            updateFormCalculations({
                                odu: Math.round(oduSum / rows.length),
                                idu: Math.round(iduSum / rows.length)
                            });
                        } else {
                            // No production data for this date/plant - default to 0
                            updateFormCalculations({ odu: 0, idu: 0 });
                        }
                    } catch (err) {
                        console.error("Error fetching production_summary:", err);
                        if (isMounted) {
                            updateFormCalculations({ odu: 0, idu: 0 });
                        }
                    }
                };

                fetchSupabaseProductionTotals();

                return () => { isMounted = false; };
            }, [entryFormValues.date, entryFormValues.plant, entryFormValues.location, isFormOpen]);

            // Auto lookup plant's location on plant change
            useEffect(() => {
                if (!isFormOpen) return;
                const { plant } = entryFormValues;
                if (!plant) return;
                const matchPlantObj = plants.find(p => p.plant_code === plant);
                if (matchPlantObj && matchPlantObj.location !== entryFormValues.location) {
                    setEntryFormValues(prev => ({ ...prev, location: matchPlantObj.location }));
                }
            }, [entryFormValues.plant, isFormOpen]);

            // Duplicate Date Entry detection (Feature 1)
            const isDuplicateDateEntry = useMemo(() => {
                if (!isFormOpen || !entryFormValues.date || !entryFormValues.plant) return false;
                const targetDate = String(entryFormValues.date).trim();
                const targetPlant = String(entryFormValues.plant).trim().toLowerCase();
                return dailyEntries.some(e =>
                    String(e.date).trim() === targetDate &&
                    String(e.plant).trim().toLowerCase() === targetPlant &&
                    (!editingRecord || String(e.id) !== String(editingRecord.id))
                );
            }, [isFormOpen, entryFormValues.date, entryFormValues.plant, dailyEntries, editingRecord]);

            const handleEntryFormSubmit = async (e) => {
                e.preventDefault();
                setActionLoading(true);

                if (isDuplicateDateEntry) {
                    setToast({ type: "error", message: "Duplicate entry detected. An entry already exists for this date." });
                    setActionLoading(false);
                    return;
                }

                const isMeterReset = entryFormValues.meter_changed === true || String(entryFormValues.meter_changed) === "true";
                if (!isMeterReset) {
                    if (Number(entryFormValues.electricity_closing) < Number(entryFormValues.electricity_opening)) {
                        setToast({ type: "error", message: "Electricity Daily Reading cannot be smaller than Previous Reading." });
                        setActionLoading(false);
                        return;
                    }
                } else {
                    if (Number(entryFormValues.custom_difference) < 0) {
                        setToast({ type: "error", message: "Custom net difference cannot be negative." });
                        setActionLoading(false);
                        return;
                    }
                }

                if (Number(entryFormValues.solar_generated) < 0 || Number(entryFormValues.diesel_used) < 0) {
                    setToast({ type: "error", message: "Inputs cannot be negative." });
                    setActionLoading(false);
                    return;
                }

                const validKeys = [
                    "date", "plant", "location", "department", "shift", "operator_name",
                    "electricity_opening", "electricity_closing", "electricity_consumption", "electricity_cost",
                    "solar_generated", "solar_utilized", "solar_cost",
                    "diesel_used", "diesel_cost", "total_cost", "odu", "idu",
                    "production_qty", "remarks"
                ];
                const payload = {};
                validKeys.forEach(k => {
                    if (entryFormValues[k] !== undefined) {
                        const numericFields = [
                            "electricity_opening", "electricity_closing", "electricity_consumption", "electricity_cost",
                            "solar_generated", "solar_utilized", "solar_cost",
                            "diesel_used", "diesel_cost", "total_cost", "odu", "idu",
                            "production_qty"
                        ];
                        if (numericFields.includes(k)) {
                            payload[k] = entryFormValues[k] === "" || entryFormValues[k] === null ? null : Number(entryFormValues[k]);
                        } else {
                            payload[k] = entryFormValues[k];
                        }
                    }
                });
                // Force Excel-style solar: Daily Reading × rate
                const solarUnits = Math.max(0, Number(entryFormValues.solar_closing) || Number(entryFormValues.solar_generated) || 0);
                const saveLoc = entryFormValues.location || plants.find((p) => p.plant_code === entryFormValues.plant)?.location || "";
                const saveSolarRate = resolveTariff(tariffs, "solar", entryFormValues.plant, saveLoc, entryFormValues.date) || resolveTariff(tariffs, "electricity", entryFormValues.plant, saveLoc, entryFormValues.date);
                const saveDieselRate = resolveTariff(tariffs, "diesel", entryFormValues.plant, saveLoc, entryFormValues.date);
                payload.solar_generated = solarUnits;
                payload.solar_utilized = solarUnits;
                payload.solar_cost = CalculationEngine.calculateSolarCost(solarUnits, saveSolarRate);
                payload.diesel_cost = CalculationEngine.calculateDieselCost(Number(payload.diesel_used) || 0, saveDieselRate);
                payload.total_cost = CalculationEngine.calculateTotalCost(
                    Number(payload.electricity_cost) || 0,
                    payload.solar_cost,
                    payload.diesel_cost
                );

                // Server-side duplicate validation check
                if (!editingRecord) {
                    const { data: dupCheck } = await supabase
                        .from('daily_entries')
                        .select('id')
                        .eq('date', payload.date)
                        .eq('plant', payload.plant);
                    if (dupCheck && dupCheck.length > 0) {
                        setToast({ type: "error", message: "Duplicate entry detected. An entry already exists for this date and plant." });
                        setActionLoading(false);
                        return;
                    }
                }

                if (editingRecord) {
                    const { error } = await supabase
                        .from('daily_entries')
                        .update(payload)
                        .eq('id', editingRecord.id);
                    if (error) {
                        setToast({ type: "error", message: error.message || "Failed to update record." });
                        setActionLoading(false);
                    } else {
                        recordAuditLog({
                            action: "UPDATE",
                            module: "Daily Operations",
                            recordId: editingRecord.id,
                            location: saveLoc,
                            plant: payload.plant,
                            oldValue: editingRecord,
                            newValue: payload,
                            status: "SUCCESS"
                        });
                        setToast({ type: "success", message: "Daily entry updated successfully!" });
                        loadAllDatabase();
                        setIsFormOpen(false);
                        setActionLoading(false);
                    }
                } else {
                    const { error } = await supabase
                        .from('daily_entries')
                        .insert(payload);
                    if (error) {
                        setToast({ type: "error", message: error.message || "Failed to add record." });
                        setActionLoading(false);
                    } else {
                        recordAuditLog({
                            action: "CREATE",
                            module: "Daily Operations",
                            recordId: payload.date,
                            location: saveLoc,
                            plant: payload.plant,
                            newValue: payload,
                            status: "SUCCESS"
                        });
                        setToast({ type: "success", message: "Daily entry added successfully!" });
                        loadAllDatabase();
                        setIsFormOpen(false);
                        setActionLoading(false);
                    }
                }
            };

            const handleDeleteEntry = async (id) => {
                const recordToDelete = dailyEntries.find(e => e.id === id);
                const ok = await askConfirm(
                    "Are you sure you want to delete this daily operations entry?",
                    { title: "Delete Entry", danger: true }
                );
                if (!ok) return;

                setActionLoading(true);
                const { error } = await supabase
                    .from('daily_entries')
                    .delete()
                    .eq('id', id);
                if (error) {
                    setToast({ type: "error", message: error.message || "Delete failed." });
                    setActionLoading(false);
                } else {
                    recordAuditLog({
                        action: "DELETE",
                        module: "Daily Operations",
                        recordId: id,
                        location: recordToDelete?.location,
                        plant: recordToDelete?.plant,
                        oldValue: recordToDelete,
                        status: "SUCCESS"
                    });
                    setToast({ type: "success", message: "Daily entry deleted successfully!" });
                    loadAllDatabase();
                    setActionLoading(false);
                }
            };

            const handleBulkDelete = async () => {
                const targetIds = selectedRowIds.size > 0
                    ? Array.from(selectedRowIds)
                    : filteredEntries.map(e => e.id);

                if (targetIds.length === 0) {
                    setToast({ type: "error", message: "No entries to delete. Please apply a filter or select entries first." });
                    return;
                }

                const ok = await askConfirm(
                    selectedRowIds.size > 0
                        ? `Are you sure you want to delete the ${selectedRowIds.size} selected entries?`
                        : `Are you sure you want to delete all ${targetIds.length} filtered entries? This action is irreversible.`,
                    { title: "Bulk Delete Entries", danger: true }
                );
                if (!ok) return;

                setActionLoading(true);
                const { error } = await supabase
                    .from('daily_entries')
                    .delete()
                    .in('id', targetIds);

                if (error) {
                    setToast({ type: "error", message: error.message || "Bulk delete failed." });
                } else {
                    recordAuditLog({
                        action: "DELETE",
                        module: "Daily Operations",
                        recordId: `Bulk (${targetIds.length})`,
                        oldValue: { deleted_ids: targetIds },
                        status: "SUCCESS"
                    });
                    setToast({ type: "success", message: `Deleted ${targetIds.length} entries successfully!` });
                    setSelectedRowIds(new Set());
                    setIsDailyDeleteMode(false);
                    loadAllDatabase();
                }
                setActionLoading(false);
            };

            const handleToggleSelectAllDaily = () => {
                const visibleIds = paginatedEntries.map(e => e.id);
                const allSelected = visibleIds.length > 0 && visibleIds.every(id => selectedRowIds.has(id));

                setSelectedRowIds(prev => {
                    const next = new Set(prev);
                    if (allSelected) {
                        visibleIds.forEach(id => next.delete(id));
                    } else {
                        visibleIds.forEach(id => next.add(id));
                    }
                    return next;
                });
            };

            const handleToggleRowSelect = (id) => {
                setSelectedRowIds(prev => {
                    const next = new Set(prev);
                    if (next.has(id)) {
                        next.delete(id);
                    } else {
                        next.add(id);
                    }
                    return next;
                });
            };

            const handleToggleSelectAllMaster = () => {
                const list = getFilteredMasterList(selectedMasterTable);
                const pkCol = getTablePrimaryKey(selectedMasterTable);
                const visibleKeys = list.map(row => String(row[pkCol])).filter(Boolean);
                const allSelected = visibleKeys.length > 0 && visibleKeys.every(k => selectedMasterRowKeys.has(k));

                setSelectedMasterRowKeys(prev => {
                    const next = new Set(prev);
                    if (allSelected) {
                        visibleKeys.forEach(k => next.delete(k));
                    } else {
                        visibleKeys.forEach(k => next.add(k));
                    }
                    return next;
                });
            };

            const handleToggleMasterRowSelect = (pkVal) => {
                const keyStr = String(pkVal);
                setSelectedMasterRowKeys(prev => {
                    const next = new Set(prev);
                    if (next.has(keyStr)) {
                        next.delete(keyStr);
                    } else {
                        next.add(keyStr);
                    }
                    return next;
                });
            };

            const handleBulkDeleteMaster = async () => {
                if (selectedMasterRowKeys.size === 0) return;
                const ok = await askConfirm(
                    `Are you sure you want to delete the ${selectedMasterRowKeys.size} selected records from master table '${selectedMasterTable}'?`,
                    { title: "Delete Selected Master Records", danger: true }
                );
                if (!ok) return;

                setActionLoading(true);
                const dbTable = getMasterDbTable(selectedMasterTable);
                const pkCol = getTablePrimaryKey(selectedMasterTable);
                const keysToDelete = Array.from(selectedMasterRowKeys);

                const { error } = await supabase
                    .from(dbTable)
                    .delete()
                    .in(pkCol, keysToDelete);

                if (error) {
                    setToast({ type: "error", message: error.message || "Master bulk delete failed." });
                } else {
                    setToast({ type: "success", message: `Deleted ${keysToDelete.length} master records successfully!` });
                    setSelectedMasterRowKeys(new Set());
                    setIsMasterDeleteMode(false);
                    loadAllDatabase();
                }
                setActionLoading(false);
            };

            // Reset master selections when changing master table
            useEffect(() => {
                setSelectedMasterRowKeys(new Set());
                setIsMasterDeleteMode(false);
            }, [selectedMasterTable]);

            // Restrict non-admin users to tariff + MF master only
            useEffect(() => {
                if (!currentUser || currentUser.role === "IT_ADMIN") return;
                if (activeTab === "master" && !["tariff_rates", "multiply_factors"].includes(selectedMasterTable)) {
                    setSelectedMasterTable("tariff_rates");
                }
            }, [activeTab, currentUser, selectedMasterTable]);

            const openMasterForm = (record = null) => {
                if (record) {
                    if (currentUser.role !== "IT_ADMIN" && !userCanAccessMasterRow(currentUser, record, selectedMasterTable, plants)) {
                        setToast({ type: "error", message: "You can only edit records for your assigned location/plant." });
                        return;
                    }
                    setEditingMasterRecord(record);
                    const updatedRec = { ...record };
                    if (selectedMasterTable === "tariff_rates" && (!updatedRec.type || updatedRec.type === "")) {
                        updatedRec.type = "electricity";
                    }
                    if (selectedMasterTable === "users" && (!updatedRec.role || updatedRec.role === "")) {
                        updatedRec.role = "USER";
                    }
                    setMasterFormValues(updatedRec);
                } else {
                    setEditingMasterRecord(null);
                    const defaultValues = {};
                    const scopePlants = currentUser?.role === "IT_ADMIN" ? plants : allowedPlants;
                    const scope = getUserDefaultMasterScope(currentUser, scopePlants);
                    const today = new Date().toISOString().split("T")[0];

                    // Get columns headers based on selected tab
                    const headers = getMasterTableHeaders(selectedMasterTable);
                    headers.forEach(h => {
                        if (h === "status") defaultValues[h] = "Active";
                        else if (h === "type" && selectedMasterTable === "tariff_rates") defaultValues[h] = "electricity";
                        else if (h === "role" && selectedMasterTable === "users") defaultValues[h] = "USER";
                        else if (h === "allowed_locations" || h === "allowed_plants") defaultValues[h] = "";
                        else if (h === "location" && (selectedMasterTable === "tariff_rates" || selectedMasterTable === "multiply_factors")) defaultValues[h] = scope.location;
                        else if (h === "plant_code" && (selectedMasterTable === "tariff_rates" || selectedMasterTable === "multiply_factors")) defaultValues[h] = scope.plant_code;
                        else if (h === "effective_date" && (selectedMasterTable === "tariff_rates" || selectedMasterTable === "multiply_factors")) defaultValues[h] = today;
                        else if (h === "factor" && selectedMasterTable === "multiply_factors") {
                            defaultValues[h] = resolveMultiplyFactor(multiplyFactors, scope.plant_code, scope.location, today);
                        }
                        else defaultValues[h] = "";
                    });

                    setMasterFormValues(defaultValues);
                }
                setIsMasterFormOpen(true);
            };

            const handleMasterFormSubmit = async (e) => {
                e.preventDefault();
                setActionLoading(true);

                const tableKeyColumn = getTablePrimaryKey(selectedMasterTable);
                let payload = { ...masterFormValues };

                if (currentUser.role !== "IT_ADMIN") {
                    if (!["tariff_rates", "multiply_factors"].includes(selectedMasterTable)) {
                        setToast({ type: "error", message: "Access denied." });
                        setActionLoading(false);
                        return;
                    }
                    if (!userCanAccessMasterRow(currentUser, payload, selectedMasterTable, plants)) {
                        setToast({ type: "error", message: "You can only save records for your assigned location/plant." });
                        setActionLoading(false);
                        return;
                    }
                }

                if (selectedMasterTable === "users") {
                    if (!payload.status) payload.status = "Active";
                    if (!payload.created_date) payload.created_date = new Date().toISOString().split('T')[0];
                    
                    const userEmail = (payload.email || "").trim().toLowerCase();
                    if (!userEmail) {
                        setToast({ type: "error", message: "Email address is required." });
                        setActionLoading(false);
                        return;
                    }
                    if (!userEmail.endsWith("@pgel.in")) {
                        setToast({ type: "error", message: "Invalid email. Only corporate email addresses ending with @pgel.in can be registered." });
                        setActionLoading(false);
                        return;
                    }
                }

                const tableKeyValue = payload[tableKeyColumn];
                const dbTable = getMasterDbTable(selectedMasterTable);
                const todayStr = new Date().toISOString().split("T")[0];
                if (selectedMasterTable === "multiply_factors") {
                    payload.factor = Number(payload.factor);
                    if (!payload.plant_code) payload.plant_code = null;
                    if (!payload.mf_id) {
                        const tag = payload.plant_code || payload.location || "MF";
                        payload.mf_id = `MF-${String(tag).replace(/\s/g, "").toUpperCase()}-${String(payload.effective_date || todayStr).replace(/-/g, "")}`;
                    }
                }
                if (selectedMasterTable === "tariff_rates") {
                    payload.rate = Number(payload.rate);
                    if (!payload.plant_code) payload.plant_code = null;
                    if (!payload.tariff_id) {
                        const tag = payload.plant_code || payload.location || "TF";
                        payload.tariff_id = `TF-${String(tag).replace(/\s/g, "").toUpperCase()}-${String(payload.type || "E").slice(0, 1).toUpperCase()}-${String(payload.effective_date || todayStr).replace(/-/g, "")}`;
                    }
                }
                if (editingMasterRecord) {
                    // Plants: allow changing plant_code by renaming FK children first (data stays, plant stays)
                    if (selectedMasterTable === "plants") {
                        const oldCode = String(editingMasterRecord.plant_code || "").trim();
                        const newCode = String(payload.plant_code || "").trim();
                        if (!newCode) {
                            setToast({ type: "error", message: "Plant code is required." });
                            setActionLoading(false);
                            return;
                        }

                        if (oldCode && newCode && oldCode !== newCode) {
                            const renameErr = await renamePlantCodeEverywhere(oldCode, newCode, payload);
                            if (renameErr) {
                                setToast({ type: "error", message: renameErr });
                                setActionLoading(false);
                                return;
                            }
                            setToast({ type: "success", message: `Plant code updated ${oldCode} → ${newCode}. Related data kept.` });
                            loadAllDatabase();
                            setIsMasterFormOpen(false);
                            setActionLoading(false);
                            return;
                        }
                    }

                    const { error } = await supabase
                        .from(dbTable)
                        .update(payload)
                        .eq(tableKeyColumn, editingMasterRecord[tableKeyColumn]);
                    if (error) {
                        setToast({ type: "error", message: error.message || "Failed to update record." });
                        setActionLoading(false);
                    } else {
                        setToast({ type: "success", message: `Record updated in table ${selectedMasterTable}` });
                        loadAllDatabase();
                        setIsMasterFormOpen(false);
                        setActionLoading(false);
                    }
                } else {
                    const { error } = await supabase
                        .from(dbTable)
                        .insert(payload);
                    if (error) {
                        setToast({ type: "error", message: error.message || "Failed to add record." });
                        setActionLoading(false);
                    } else {
                        setToast({ type: "success", message: `Record added to table ${selectedMasterTable}` });
                        loadAllDatabase();
                        setIsMasterFormOpen(false);
                        setActionLoading(false);
                    }
                }
            };

            // Rename plant_code safely: keep plant + all related rows (daily_entries, meters, targets, etc.)
            const renamePlantCodeEverywhere = async (oldCode, newCode, plantPayload) => {
                // 1) Create new plant row (same details, new code)
                const newPlant = {
                    plant_code: newCode,
                    location: plantPayload.location ?? "",
                    plant_name: plantPayload.plant_name ?? "",
                    plant_display_name: plantPayload.plant_display_name ?? "",
                    status: plantPayload.status || "Active",
                };
                const { error: insertErr } = await supabase.from("plants").insert(newPlant);
                if (insertErr) {
                    return insertErr.message || "Failed to create plant with new code.";
                }

                // 2) Repoint child tables from old → new
                const childUpdates = [
                    supabase.from("daily_entries").update({ plant: newCode }).eq("plant", oldCode),
                    supabase.from("meters").update({ plant_code: newCode }).eq("plant_code", oldCode),
                    supabase.from("solar_meters").update({ plant_code: newCode }).eq("plant_code", oldCode),
                    supabase.from("water_meters").update({ plant_code: newCode }).eq("plant_code", oldCode),
                    supabase.from("air_meters").update({ plant_code: newCode }).eq("plant_code", oldCode),
                    supabase.from("dg_sets").update({ plant_code: newCode }).eq("plant_code", oldCode),
                    supabase.from("target_values").update({ plant_code: newCode }).eq("plant_code", oldCode),
                    supabase.from("tariffs").update({ plant_code: newCode }).eq("plant_code", oldCode),
                    supabase.from("multiply_factors").update({ plant_code: newCode }).eq("plant_code", oldCode),
                ];
                const childResults = await Promise.all(childUpdates);
                const childFail = childResults.find((r) => r.error);
                if (childFail?.error) {
                    // rollback new plant if children failed
                    await supabase.from("plants").delete().eq("plant_code", newCode);
                    return childFail.error.message || "Failed to move related plant data.";
                }

                // 3) Update users.allowed_plants lists that mention old code
                try {
                    const { data: userRows } = await supabase.from("users").select("id, allowed_plants");
                    const patches = (userRows || []).filter((u) => {
                        const raw = String(u.allowed_plants || "");
                        if (!raw || raw.toLowerCase() === "all") return false;
                        return raw.split(",").map((x) => x.trim()).includes(oldCode);
                    }).map((u) => {
                        const next = String(u.allowed_plants)
                            .split(",")
                            .map((x) => x.trim())
                            .map((x) => (x === oldCode ? newCode : x))
                            .filter(Boolean)
                            .join(",");
                        return supabase.from("users").update({ allowed_plants: next }).eq("id", u.id);
                    });
                    if (patches.length) await Promise.all(patches);
                } catch (_) { /* non-fatal */ }

                // 4) Remove old plant row (children already moved)
                const { error: delErr } = await supabase.from("plants").delete().eq("plant_code", oldCode);
                if (delErr) {
                    return delErr.message || "New code saved, but old plant row could not be removed.";
                }
                return null;
            };

                    const handleDeleteMasterRecord = async (pkVal) => {
            const pkCol = getTablePrimaryKey(selectedMasterTable);
            const record = getStateMasterList(selectedMasterTable).find((r) => String(r[pkCol]) === String(pkVal));
            if (currentUser.role !== "IT_ADMIN") {
                if (!["tariff_rates", "multiply_factors"].includes(selectedMasterTable)) {
                    setToast({ type: "error", message: "Access denied." });
                    return;
                }
                if (!userCanAccessMasterRow(currentUser, record || {}, selectedMasterTable, plants)) {
                    setToast({ type: "error", message: "You can only delete records for your assigned location/plant." });
                    return;
                }
            }
            const ok = await askConfirm(
                `Are you sure you want to delete record '${pkVal}' from master table ${selectedMasterTable}?`,
                { title: "Delete Record", danger: true }
            );
            if (!ok) return;

            setActionLoading(true);
            const dbTable = getMasterDbTable(selectedMasterTable);
            const { error } = await supabase
                .from(dbTable)
                .delete()
                .eq(pkCol, pkVal);
            if (error) {
                setToast({ type: "error", message: error.message || "Delete failed." });
                setActionLoading(false);
            } else {
                setToast({ type: "success", message: "Record deleted successfully." });
                loadAllDatabase();
                setActionLoading(false);
            }
        };

            function getStateMasterList(tableName) {
                switch (tableName) {
                    case "plants": return plants;
                    case "departments": return departments;
                    case "meters": return meters;
                    case "solar_meters": return solarMeters;
                    case "water_meters": return waterMeters;
                    case "air_meters": return airMeters;
                    case "dg_sets": return dgSets;
                    case "fuel_types": return fuelTypes;
                    case "products": return products;
                    case "tariff_rates": return tariffs;
                    case "multiply_factors": return multiplyFactors;
                    case "target_values": return targets;
                    case "users": return users;
                    default: return [];
                }
            }

            function getFilteredMasterList(tableName) {
                const list = getStateMasterList(tableName);
                if (currentUser?.role === "IT_ADMIN") return list;
                if (tableName === "tariff_rates" || tableName === "multiply_factors") {
                    return list.filter((row) => userCanAccessMasterRow(currentUser, row, tableName, plants));
                }
                return [];
            }

            function getMasterDbTable(tableName) {
                if (tableName === "tariff_rates") return "tariffs";
                return tableName;
            }

            function addStateMasterList(tableName, record) {
                const add = list => [record, ...list];
                if (tableName === "plants") setPlants(add);
                else if (tableName === "departments") setDepartments(add);
                else if (tableName === "meters") setMeters(add);
                else if (tableName === "solar_meters") setSolarMeters(add);
                else if (tableName === "water_meters") setWaterMeters(add);
                else if (tableName === "air_meters") setAirMeters(add);
                else if (tableName === "dg_sets") setDgSets(add);
                else if (tableName === "fuel_types") setFuelTypes(add);
                else if (tableName === "products") setProducts(add);
                else if (tableName === "tariff_rates") setTariffs(add);
                else if (tableName === "multiply_factors") setMultiplyFactors(add);
                else if (tableName === "target_values") setTargets(add);
                else if (tableName === "users") setUsers(add);
            }

            function updateStateMasterList(tableName, record, origPk) {
                const pk = getTablePrimaryKey(tableName);
                const update = list => list.map(item => String(item[pk]) === String(origPk) ? record : item);
                if (tableName === "plants") setPlants(update);
                else if (tableName === "departments") setDepartments(update);
                else if (tableName === "meters") setMeters(update);
                else if (tableName === "solar_meters") setSolarMeters(update);
                else if (tableName === "water_meters") setWaterMeters(update);
                else if (tableName === "air_meters") setAirMeters(update);
                else if (tableName === "dg_sets") setDgSets(update);
                else if (tableName === "fuel_types") setFuelTypes(update);
                else if (tableName === "products") setProducts(update);
                else if (tableName === "tariff_rates") setTariffs(update);
                else if (tableName === "multiply_factors") setMultiplyFactors(update);
                else if (tableName === "target_values") setTargets(update);
                else if (tableName === "users") setUsers(update);
            }

            function deleteStateMasterList(tableName, pkColumn, keyVal) {
                const del = list => list.filter(item => String(item[pkColumn]) !== String(keyVal));
                if (tableName === "plants") setPlants(del);
                else if (tableName === "departments") setDepartments(del);
                else if (tableName === "meters") setMeters(del);
                else if (tableName === "solar_meters") setSolarMeters(del);
                else if (tableName === "water_meters") setWaterMeters(del);
                else if (tableName === "air_meters") setAirMeters(del);
                else if (tableName === "dg_sets") setDgSets(del);
                else if (tableName === "fuel_types") setFuelTypes(del);
                else if (tableName === "products") setProducts(del);
                else if (tableName === "tariff_rates") setTariffs(del);
                else if (tableName === "multiply_factors") setMultiplyFactors(del);
                else if (tableName === "target_values") setTargets(del);
                else if (tableName === "users") setUsers(del);
            }

            function getMasterTableHeaders(tableName) {
                switch (tableName) {
                    case "plants": return ["plant_code", "location", "plant_name", "plant_display_name", "status"];
                    case "departments": return ["dept_code", "dept_name", "status"];
                    case "meters": return ["meter_id", "meter_name", "plant_code", "dept_code", "status"];
                    case "solar_meters": return ["meter_id", "meter_name", "plant_code", "status"];
                    case "water_meters": return ["meter_id", "meter_name", "plant_code", "status"];
                    case "air_meters": return ["meter_id", "meter_name", "plant_code", "status"];
                    case "dg_sets": return ["dg_id", "dg_name", "plant_code", "status"];
                    case "fuel_types": return ["fuel_id", "fuel_name", "status"];
                    case "products": return ["product_id", "product_name", "unit", "status"];
                    case "tariff_rates": return ["tariff_id", "location", "plant_code", "type", "rate", "effective_date", "status"];
                    case "multiply_factors": return ["mf_id", "location", "plant_code", "factor", "effective_date", "status"];
                    case "target_values": return ["target_id", "plant_code", "department_code", "year", "month", "electricity_target_kwh", "water_target_kl", "diesel_target_l", "lpg_target_kg", "solar_generation_target_kwh", "status"];
                    case "users": return ["id", "name", "email", "role", "allowed_locations", "allowed_plants"];
                    default: return [];
                }
            }

            function getTablePrimaryKey(tableName) {
                switch (tableName) {
                    case "plants": return "plant_code";
                    case "departments": return "dept_code";
                    case "meters": return "meter_id";
                    case "solar_meters": return "meter_id";
                    case "water_meters": return "meter_id";
                    case "air_meters": return "meter_id";
                    case "dg_sets": return "dg_id";
                    case "fuel_types": return "fuel_id";
                    case "products": return "product_id";
                    case "tariff_rates": return "tariff_id";
                    case "multiply_factors": return "mf_id";
                    case "target_values": return "target_id";
                    case "users": return "id";
                    default: return "id";
                }
            }

            // Trigger sheets database seed reset (Warning action)
            const handleSetupSheetsSeeding = async () => {
                const ok = await askConfirm(
                    "This will drop all your spreadsheet tables, recreate them, and seed them with massive operational data to enable demo graphs. Do you want to proceed?",
                    { title: "âš ï¸ Reset Entire Database", confirmText: "Yes, Reset", danger: true }
                );
                if (!ok) return;
                setActionLoading(true);

                if (typeof google !== "undefined" && google.script && google.script.run) {
                    google.script.run
                        .withSuccessHandler(() => {
                            setToast({ type: "success", message: "Database reseeded completely!" });
                            loadAllDatabase();
                            setActionLoading(false);
                        })
                        .withFailureHandler((err) => {
                            setToast({ type: "error", message: err.message || "Reset failed." });
                            setActionLoading(false);
                        })
                        .setupSheet();
                } else {
                    setTimeout(() => {
                        setDailyEntries(generateMockEntries());
                        setToast({ type: "success", message: "Offline database seed reloaded successfully!" });
                        setActionLoading(false);
                    }, 1000);
                }
            };

            // ----------------------------------------------------
            // CALCULATING CHARTS DATA & METRICS
            // ----------------------------------------------------
            const filteredEntries = useMemo(() => {
                const plantLocationMap = {};
                plants.forEach((p) => {
                    const loc = p.location;
                    [p.plant_code, p.plant_display_name, p.plant_name].forEach((key) => {
                        if (key) plantLocationMap[String(key)] = loc;
                    });
                });
                const allowedIds = plantIdentitySet(allowedPlants);
                const selectedPlant = filters.plant !== "all"
                    ? plants.find((p) => p.plant_code === filters.plant)
                    : null;
                const selectedIds = selectedPlant
                    ? plantIdentitySet([selectedPlant])
                    : null;

                return dailyEntries.filter((e) => {
                    const ep = entryPlantKey(e.plant);

                    // All plants: include every entry the user is allowed to see
                    // (match code OR display name). If masters still loading, don't zero-out.
                    let matchesPlant;
                    if (filters.plant === "all") {
                        if (!plants.length) matchesPlant = true;
                        else if (!allowedPlants.length) matchesPlant = false;
                        else matchesPlant = allowedIds.has(ep);
                    } else if (selectedIds) {
                        matchesPlant = selectedIds.has(ep);
                    } else {
                        matchesPlant = e.plant === filters.plant;
                    }

                    const matchesDept = filters.department === "all" || e.department === filters.department;
                    const matchesLocation = !filters.location || filters.location === "all"
                        || plantLocationMap[e.plant] === filters.location;
                    const matchesDate = e.date >= filters.startDate && e.date <= filters.endDate;
                    return matchesPlant && matchesDept && matchesLocation && matchesDate;
                });
            }, [dailyEntries, filters, allowedPlants, plants]);

            // Unique locations derived from plants master data (used for Reports location filter)
            const reportLocations = useMemo(() => {
                const uniqueLocs = [...new Set(allowedPlants.map(p => p.location).filter(Boolean))];
                return uniqueLocs;
            }, [allowedPlants]);

            // Plants narrowed down by the selected location (used for Reports plant dropdown)
            const reportLocationPlants = useMemo(() => {
                if (selectedReportLocation === "all") return allowedPlants;
                return allowedPlants.filter(p => p.location === selectedReportLocation);
            }, [allowedPlants, selectedReportLocation]);

            // Plants for Daily Entry location filter
            const entryFilterPlants = useMemo(() => {
                if (entryLocationFilter === "all") return allowedPlants;
                return allowedPlants.filter(p => p.location === entryLocationFilter);
            }, [allowedPlants, entryLocationFilter]);

            // Calculate aggregations for KPI Cards
            const kpiTotals = useMemo(() => {
                return filteredEntries.reduce((acc, cur) => {
                    acc.electricity += Number(cur.electricity_consumption) || 0;
                    acc.electricityCost += Number(cur.electricity_cost) || 0;
                    acc.solarGenerated += Number(cur.solar_generated) || 0;
                    acc.solarUtilized += Number(cur.solar_utilized) || 0;
                    acc.solarCost += Number(cur.solar_cost) || 0;
                    acc.water += Number(cur.water_consumption) || 0;
                    acc.waterCost += Number(cur.water_cost) || 0;
                    acc.air += Number(cur.air_consumption) || 0;
                    acc.diesel += Number(cur.diesel_used) || 0;
                    acc.dieselCost += Number(cur.diesel_cost) || 0;
                    acc.lpg += Number(cur.lpg_used) || 0;
                    acc.lpgCost += Number(cur.lpg_cost) || 0;
                    acc.production += Number(cur.production_qty) || 0;
                    acc.odu += Number(cur.odu) || 0;
                    acc.idu += Number(cur.idu) || 0;
                    acc.wasteHaz += Number(cur.waste_hazardous) || 0;
                    acc.wasteNHaz += Number(cur.waste_non_hazardous) || 0;
                    acc.wasteRec += Number(cur.waste_recycled) || 0;
                    return acc;
                }, {
                    electricity: 0, electricityCost: 0, solarGenerated: 0, solarUtilized: 0, solarCost: 0,
                    water: 0, waterCost: 0, air: 0, diesel: 0, dieselCost: 0, lpg: 0, lpgCost: 0,
                    production: 0, odu: 0, idu: 0, wasteHaz: 0, wasteNHaz: 0, wasteRec: 0
                });
            }, [filteredEntries]);

            const aggregatedCosts = useMemo(() => {
                const electRate = Number(activeElectRate) || 10.89;
                const solarRate = Number(activeSolarRate) || electRate;
                const dieselRate = Number(activeDieselRate) || 90.62;
                const lpgRate = Number(activeLpgRate) || 85;
                const waterRate = Number(activeWaterRate) || 45;

                const electricityCost = kpiTotals.electricityCost;
                const solarCost = kpiTotals.solarCost;
                const dieselCost = kpiTotals.dieselCost;
                const lpgCost = kpiTotals.lpgCost;
                const energyCost = electricityCost + solarCost + dieselCost + lpgCost;
                const waterCost = kpiTotals.waterCost || (kpiTotals.water * waterRate);
                const totalCost = energyCost + waterCost;
                const totalConsumption = kpiTotals.electricity + kpiTotals.solarGenerated;
                const sec = kpiTotals.production > 0 ? kpiTotals.electricity / kpiTotals.production : 0;
                const solarUtilPct = kpiTotals.solarGenerated > 0 ? (kpiTotals.solarUtilized / kpiTotals.solarGenerated) * 100 : 0;

                return {
                    energyCost,
                    waterCost,
                    totalCost,
                    sec,
                    solarUtilPct,
                    totalConsumption,
                    electRate,
                    electricityCost,
                    solarCost,
                    dieselCost,
                    lpgCost,
                };
            }, [kpiTotals, activeElectRate, activeSolarRate, activeDieselRate, activeLpgRate, activeWaterRate]);

            // Trend datasets for Recharts
            const dailyTrendsData = useMemo(() => {
                const map = {};
                filteredEntries.forEach(e => {
                    if (!map[e.date]) {
                        map[e.date] = { date: e.date, electricity: 0, solarGen: 0, solarUtil: 0, totalConsumption: 0, water: 0, diesel: 0, lpg: 0, production: 0, odu: 0, idu: 0, cost: 0, waste: 0 };
                    }
                    map[e.date].electricity += Number(e.electricity_consumption) || 0;
                    map[e.date].solarGen += Number(e.solar_generated) || 0;
                    map[e.date].solarUtil += Number(e.solar_utilized) || 0;
                    map[e.date].totalConsumption = (map[e.date].electricity || 0) + (map[e.date].solarGen || 0);
                    map[e.date].water += Number(e.water_consumption) || 0;
                    map[e.date].diesel += Number(e.diesel_used) || 0;
                    map[e.date].lpg += Number(e.lpg_used) || 0;
                    map[e.date].production += Number(e.production_qty) || 0;
                    map[e.date].odu += Number(e.odu) || 0;
                    map[e.date].idu += Number(e.idu) || 0;
                    map[e.date].cost += (Number(e.electricity_cost) || 0) + (Number(e.diesel_cost) || 0) + (Number(e.lpg_cost) || 0);
                    map[e.date].waste += (Number(e.waste_hazardous) || 0) + (Number(e.waste_non_hazardous) || 0);
                });
                return Object.values(map).sort((a, b) => a.date.localeCompare(b.date));
            }, [filteredEntries]);

            // Last 7 consecutive calendar days (day-wise) — never sparse month jumps
            const last7DaysTrendsData = useMemo(() => {
                const byDate = {};
                dailyTrendsData.forEach(d => {
                    const k = String(d.date || "").slice(0, 10);
                    if (k) byDate[k] = d;
                });

                // Prefer latest date that actually has filtered data, else filter endDate
                let endRaw = filters.endDate || new Date().toISOString().slice(0, 10);
                if (dailyTrendsData.length) {
                    const latest = String(dailyTrendsData[dailyTrendsData.length - 1].date || "").slice(0, 10);
                    if (latest && latest < endRaw) endRaw = latest;
                }
                const end = new Date(endRaw + "T00:00:00");
                if (Number.isNaN(end.getTime())) return [];

                const pad = (n) => String(n).padStart(2, "0");
                const out = [];
                for (let i = 6; i >= 0; i--) {
                    const dt = new Date(end);
                    dt.setDate(end.getDate() - i);
                    const key = `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
                    const label = dt.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
                    const src = byDate[key];
                    out.push({
                        date: key,
                        label,
                        electricity: src ? Number(src.electricity) || 0 : 0,
                        solarGen: src ? Number(src.solarGen) || 0 : 0,
                        solarUtil: src ? Number(src.solarUtil) || 0 : 0,
                        water: src ? Number(src.water) || 0 : 0,
                        diesel: src ? Number(src.diesel) || 0 : 0,
                        lpg: src ? Number(src.lpg) || 0 : 0,
                        production: src ? Number(src.production) || 0 : 0,
                        cost: src ? Number(src.cost) || 0 : 0,
                        waste: src ? Number(src.waste) || 0 : 0,
                    });
                }
                return out;
            }, [dailyTrendsData, filters.endDate]);

            // Last 30 calendar days for detailed daily resource charts (Solar vs Total, Water, Production, LPG, Waste)
            const last30DaysTrendsData = useMemo(() => {
                if (!dailyTrendsData.length) return [];
                const sliced = dailyTrendsData.slice(-30);
                return sliced.map(d => {
                    const parts = (d.date || "").split("-");
                    const shortDate = parts.length === 3 ? `${parts[2]}/${parts[1]}` : d.date;
                    const dateObj = new Date(d.date + "T00:00:00");
                    const label = !isNaN(dateObj.getTime()) ? dateObj.toLocaleDateString("en-GB", { day: "numeric", month: "short" }) : shortDate;
                    return {
                        ...d,
                        shortDate,
                        label
                    };
                });
            }, [dailyTrendsData]);

            // DUMMY WATER DATA (UI ONLY) - Set USE_DUMMY_WATER_DATA to false when you want to remove it
            const USE_DUMMY_WATER_DATA = true;
            const waterTrendsData = useMemo(() => {
                if (!USE_DUMMY_WATER_DATA) return last30DaysTrendsData;

                const mockValues = [95, 112, 125, 108, 92, 118, 130, 122, 105, 98, 115, 128, 140, 135, 110, 102, 120, 132, 145, 138, 125, 115, 108, 124, 136, 142, 130, 118, 105, 122];

                if (last30DaysTrendsData && last30DaysTrendsData.length > 0) {
                    return last30DaysTrendsData.map((d, idx) => ({
                        ...d,
                        water: (Number(d.water) > 0) ? Number(d.water) : mockValues[idx % mockValues.length]
                    }));
                }

                // Fallback if no entries exist in date range
                const dummyDays = [];
                const today = new Date();
                const pad = (n) => String(n).padStart(2, "0");
                for (let i = 29; i >= 0; i--) {
                    const dt = new Date();
                    dt.setDate(today.getDate() - i);
                    const key = `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
                    const label = dt.toLocaleDateString("en-GB", { day: "numeric", month: "short" });
                    dummyDays.push({
                        date: key,
                        shortDate: `${pad(dt.getDate())}/${pad(dt.getMonth() + 1)}`,
                        label,
                        water: mockValues[(29 - i) % mockValues.length]
                    });
                }
                return dummyDays;
            }, [last30DaysTrendsData]);

            const displayWaterConsumption = useMemo(() => {
                if (!USE_DUMMY_WATER_DATA) return kpiTotals.water;
                if (kpiTotals.water > 0) return kpiTotals.water;
                return waterTrendsData.reduce((acc, curr) => acc + (Number(curr.water) || 0), 0);
            }, [kpiTotals.water, waterTrendsData]);

            const displayWaterCost = useMemo(() => {
                if (!USE_DUMMY_WATER_DATA) return aggregatedCosts.waterCost;
                if (aggregatedCosts.waterCost > 0) return aggregatedCosts.waterCost;
                return displayWaterConsumption * (Number(activeWaterRate) || 45);
            }, [aggregatedCosts.waterCost, displayWaterConsumption, activeWaterRate]);



            const monthlyTrendsData = useMemo(() => {
                const map = {};
                filteredEntries.forEach(e => {
                    const parts = e.date.split("-");
                    const key = `${parts[0]}-${parts[1]}`; // YYYY-MM
                    if (!map[key]) {
                        map[key] = { period: key, electricity: 0, solarGen: 0, diesel: 0, water: 0, cost: 0 };
                    }
                    map[key].electricity += Number(e.electricity_consumption) || 0;
                    map[key].solarGen += Number(e.solar_generated) || 0;
                    map[key].diesel += Number(e.diesel_used) || 0;
                    map[key].water += Number(e.water_consumption) || 0;
                    map[key].cost += (Number(e.electricity_cost) || 0) + (Number(e.diesel_cost) || 0) + (Number(e.lpg_cost) || 0);
                });
                return Object.values(map).sort((a, b) => a.period.localeCompare(b.period)).map(d => {
                    const [y, m] = d.period.split("-");
                    const dt = new Date(Number(y), Number(m) - 1, 1);
                    return {
                        ...d,
                        label: Number.isNaN(dt.getTime()) ? d.period : dt.toLocaleDateString("en-GB", { month: "short" })
                    };
                });
            }, [filteredEntries]);

            // Last 5 consecutive months (like last 7 days) — fill 0 if no data
            const last5MonthsTrendsData = useMemo(() => {
                const byPeriod = {};
                monthlyTrendsData.forEach(d => { byPeriod[d.period] = d; });

                let endRaw = filters.endDate || new Date().toISOString().slice(0, 10);
                if (monthlyTrendsData.length) {
                    const latest = monthlyTrendsData[monthlyTrendsData.length - 1].period; // YYYY-MM
                    const endYm = String(endRaw).slice(0, 7);
                    if (latest && latest < endYm) endRaw = latest + "-01";
                }
                const end = new Date(String(endRaw).slice(0, 10) + "T00:00:00");
                if (Number.isNaN(end.getTime())) return [];

                const pad = (n) => String(n).padStart(2, "0");
                const out = [];
                for (let i = 4; i >= 0; i--) {
                    const dt = new Date(end.getFullYear(), end.getMonth() - i, 1);
                    const key = `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}`;
                    const label = dt.toLocaleDateString("en-GB", { month: "short" });
                    const src = byPeriod[key];
                    out.push({
                        period: key,
                        label,
                        electricity: src ? Number(src.electricity) || 0 : 0,
                        solarGen: src ? Number(src.solarGen) || 0 : 0,
                        diesel: src ? Number(src.diesel) || 0 : 0,
                        water: src ? Number(src.water) || 0 : 0,
                        cost: src ? Number(src.cost) || 0 : 0,
                    });
                }
                return out;
            }, [monthlyTrendsData, filters.endDate]);

            const yearlyTrendsData = useMemo(() => {
                const map = {};
                filteredEntries.forEach(e => {
                    const year = (e.date || "").split("-")[0] || "Unknown";
                    if (!map[year]) {
                        map[year] = { period: year, electricity: 0, solarGen: 0, diesel: 0, water: 0, cost: 0 };
                    }
                    map[year].electricity += Number(e.electricity_consumption) || 0;
                    map[year].solarGen += Number(e.solar_generated) || 0;
                    map[year].diesel += Number(e.diesel_used) || 0;
                    map[year].water += Number(e.water_consumption) || 0;
                    map[year].cost += (Number(e.electricity_cost) || 0) + (Number(e.diesel_cost) || 0) + (Number(e.lpg_cost) || 0);
                });
                return Object.values(map).sort((a, b) => a.period.localeCompare(b.period));
            }, [filteredEntries]);

            // Last 4 consecutive years — fill 0 if no data
            const last4YearsTrendsData = useMemo(() => {
                const byYear = {};
                yearlyTrendsData.forEach(d => { byYear[d.period] = d; });

                let endYear = Number((filters.endDate || new Date().toISOString().slice(0, 10)).slice(0, 4));
                if (yearlyTrendsData.length) {
                    const latest = Number(yearlyTrendsData[yearlyTrendsData.length - 1].period);
                    if (Number.isFinite(latest) && latest < endYear) endYear = latest;
                }
                if (!Number.isFinite(endYear)) endYear = new Date().getFullYear();

                const out = [];
                for (let i = 3; i >= 0; i--) {
                    const year = String(endYear - i);
                    const src = byYear[year];
                    out.push({
                        period: year,
                        label: year,
                        electricity: src ? Number(src.electricity) || 0 : 0,
                        solarGen: src ? Number(src.solarGen) || 0 : 0,
                        diesel: src ? Number(src.diesel) || 0 : 0,
                        water: src ? Number(src.water) || 0 : 0,
                        cost: src ? Number(src.cost) || 0 : 0,
                    });
                }
                return out;
            }, [yearlyTrendsData, filters.endDate]);

            const plantComparisonData = useMemo(() => {
                const map = {};
                filteredEntries.forEach(e => {
                    if (!map[e.plant]) {
                        map[e.plant] = { name: e.plant, electricity: 0, solar: 0, cost: 0 };
                    }
                    map[e.plant].electricity += Number(e.electricity_consumption) || 0;
                    map[e.plant].solar += Number(e.solar_generated) || 0;
                    map[e.plant].cost += (Number(e.electricity_cost) || 0) + (Number(e.diesel_cost) || 0) + (Number(e.lpg_cost) || 0);
                });
                return Object.values(map);
            }, [filteredEntries]);

            // Target vs Actual compare
            const targetVsActualData = useMemo(() => {
                let electTargetTotal = 0;
                let electActualTotal = kpiTotals.electricity;

                targets.forEach(t => {
                    if (t.status === "Active" && (filters.plant === "all" || t.plant_code === filters.plant)) {
                        electTargetTotal += (Number(t.electricity_target_kwh) * 20);
                    }
                });

                if (electTargetTotal === 0) electTargetTotal = 6000 * 3;

                return [
                    { name: "Target", value: electTargetTotal },
                    { name: "Actual", value: electActualTotal },
                ];
            }, [targets, kpiTotals, filters.plant]);

            // Plant-wise electricity share (looks better than single-year ring)
            const yearlyElectricityPieData = useMemo(() => {
                return plantComparisonData
                    .map(d => {
                        const p = plants.find(x => x.plant_code === d.name || x.plant_display_name === d.name);
                        const label = p?.plant_display_name || p?.plant_name || d.name;
                        return { name: label, value: Number(d.electricity) || 0 };
                    })
                    .filter(d => d.value > 0);
            }, [plantComparisonData, plants]);

            // Cost mix pie: Electricity / Solar / Diesel / LPG
            const yearlyCostPieData = useMemo(() => {
                return [
                    { name: "Electricity", value: Number(kpiTotals.electricityCost) || 0 },
                    { name: "Solar", value: Number(kpiTotals.solarCost) || 0 },
                    { name: "Diesel", value: Number(kpiTotals.dieselCost) || 0 },
                    { name: "LPG", value: Number(kpiTotals.lpgCost) || 0 },
                ].filter(d => d.value > 0);
            }, [kpiTotals]);

            // Unique chart colors — no repeat across daily / monthly / yearly / pie
            const YEARLY_ELECT_PIE_COLORS = ["#0891b2", "#059669", "#ca8a04", "#9a3412", "#475569"];
            const YEARLY_COST_PIE_COLORS = { Electricity: "#1e3a8a", Solar: "#a16207", Diesel: "#9f1239", LPG: "#6d28d9" };
            const TARGET_PIE_COLORS = { Target: "#94a3b8", Actual: "#15803d" };
            const CHART = {
                dailyElect: "#0284c7",
                dailySolar: "#f59e0b",
                dailyDiesel: "#e11d48",
                monthlyElect: "#0d9488",
                monthlySolar: "#c026d3",
                monthlyDiesel: "#4f46e5",
                yearlyElect: "#1d4ed8",
                yearlyCost: "#db2777",
            };

            // Active Alert Logs computed from values
            const alertsList = useMemo(() => {
                const list = [];
                // Check targets
                dailyEntries.forEach(e => {
                    // Check duplicate
                    const matches = dailyEntries.filter(x => x.date === e.date && x.plant === e.plant && x.department === e.department);
                    if (matches.length > 1 && matches[0].id === e.id) {
                        list.push({
                            type: "Duplicate Entry Alert",
                            message: `Duplicate data row detected on ${e.date} for plant node ${e.plant} (${e.department})`,
                            level: "High",
                            date: e.date
                        });
                    }

                    // Check negative readings
                    if (e.electricity_consumption < 0 || e.water_consumption < 0 || e.diesel_used < 0 || e.lpg_used < 0) {
                        list.push({
                            type: "Negative Meter Value",
                            message: `Negative calculations in readings: Elect ${e.electricity_consumption} kWh, Water ${e.water_consumption} KL on date ${e.date} (${e.plant})`,
                            level: "Critical",
                            date: e.date
                        });
                    }

                    // Compare with active targets
                    const targetRow = targets.find(t => t.plant_code === e.plant && t.department_code === e.department);
                    if (targetRow && targetRow.status === "Active") {
                        if (e.electricity_consumption > Number(targetRow.electricity_target_kwh)) {
                            list.push({
                                type: "Target Exceeded",
                                message: `Electricity usage of ${e.electricity_consumption} kWh exceeds threshold ${targetRow.electricity_target_kwh} kWh at ${e.plant} (${e.department})`,
                                level: "Warning",
                                date: e.date
                            });
                        }
                        if (e.water_consumption > Number(targetRow.water_target_kl)) {
                            list.push({
                                type: "Target Exceeded",
                                message: `Water usage of ${e.water_consumption} KL exceeds threshold ${targetRow.water_target_kl} KL at ${e.plant} (${e.department})`,
                                level: "Warning",
                                date: e.date
                            });
                        }
                    }
                });

                return list.slice(0, 15); // limit to 15 logs
            }, [dailyEntries, targets]);

            // ----------------------------------------------------
            // DAILY DATA TABULATION PAGINATION
            // ----------------------------------------------------
            const searchedEntries = useMemo(() => {
                const allowedIds = plantIdentitySet(allowedPlants);
                const plantLocationMap = {};
                plants.forEach((p) => {
                    const loc = p.location;
                    [p.plant_code, p.plant_display_name, p.plant_name].forEach((key) => {
                        if (key) plantLocationMap[String(key)] = loc;
                    });
                });
                const selectedPlant = entryPlantFilter !== "all"
                    ? plants.find((p) => p.plant_code === entryPlantFilter)
                    : null;
                const selectedIds = selectedPlant ? plantIdentitySet([selectedPlant]) : null;

                return dailyEntries.filter(e => {
                    const ep = entryPlantKey(e.plant);
                    const isAllowed = !plants.length
                        ? true
                        : (!allowedPlants.length ? false : allowedIds.has(ep));
                    if (!isAllowed) return false;

                    if (entryLocationFilter !== "all") {
                        const loc = e.location || plantLocationMap[e.plant];
                        if (String(loc || "").toUpperCase() !== String(entryLocationFilter).toUpperCase()) return false;
                    }

                    if (entryPlantFilter !== "all") {
                        if (selectedIds) {
                            if (!selectedIds.has(ep)) return false;
                        } else if (e.plant !== entryPlantFilter) {
                            return false;
                        }
                    }

                    const searchLower = entrySearch.toLowerCase();
                    if (!searchLower) return true;
                    return (
                        String(e.plant || "").toLowerCase().includes(searchLower) ||
                        String(e.operator_name || "").toLowerCase().includes(searchLower) ||
                        (e.remarks || "").toLowerCase().includes(searchLower) ||
                        String(e.date || "").includes(searchLower)
                    );
                });
            }, [dailyEntries, entrySearch, entryLocationFilter, entryPlantFilter, allowedPlants, plants]);

            const paginatedEntries = useMemo(() => {
                const start = (entryPage - 1) * entryLimit;
                return searchedEntries.slice(start, start + entryLimit);
            }, [searchedEntries, entryPage]);

            const totalPages = Math.ceil(searchedEntries.length / entryLimit) || 1;

            // ----------------------------------------------------
            // EXPORTS + PRINT — Colourful plant report (same layout)
            // ----------------------------------------------------
            const formatExportDate = (iso) => {
                if (!iso) return "";
                const d = new Date(`${iso}T00:00:00`);
                if (isNaN(d.getTime())) return String(iso);
                const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                return `${String(d.getDate()).padStart(2, "0")}-${months[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
            };

            const resolvePlantMeta = (plantKey) => {
                const p = (plants || []).find((x) =>
                    [x.plant_code, x.plant_display_name, x.plant_name]
                        .filter(Boolean)
                        .some((v) => String(v).trim().toLowerCase() === String(plantKey || "").trim().toLowerCase())
                );
                return {
                    code: p?.plant_code || plantKey || "",
                    name: p?.plant_display_name || p?.plant_name || plantKey || "",
                    location: p?.location || "",
                };
            };

            const fmtInrExport = (n) => {
                if (n === null || n === undefined || n === "" || isNaN(Number(n))) return "—";
                return `₹ ${new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(Math.round(Number(n)))}`;
            };

            const PLANT_REPORT_TYPE_ROW = [
                "Auto", "Auto", "Manual", "Auto", "Auto", "Manual", "Auto", "Auto",
                "Auto", "Auto", "Auto", "Manual", "Auto", "Auto", "Auto", "Auto", "Auto", "Auto",
            ];

            const createPlantReportWorkbook = (payload) => {
                const { list, locationLabel, plantLabel, electRate, dieselRate, dataRows, gridLabel, reportHeaders, reportFormulaRow } = payload;
                const COLS = 18;
                const MANUAL_COLS = new Set([3, 6, 12]);

                const thin = { style: "thin", color: { argb: "FF64748B" } };
                const borderAll = { top: thin, left: thin, bottom: thin, right: thin };
                const center = { vertical: "middle", horizontal: "center", wrapText: true };
                const fillGreen = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC6EFCE" } };
                const fillYellow = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
                const fillBlue = { type: "pattern", pattern: "solid", fgColor: { argb: "FFBDD7EE" } };
                const fillTitle = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0C4A6E" } };
                const fillMeta = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0F2FE" } };
                const fontWhiteBold = { bold: true, color: { argb: "FFFFFFFF" }, size: 14, name: "Calibri" };
                const fontBold = { bold: true, size: 10, name: "Calibri", color: { argb: "FF0F172A" } };
                const fontSmall = { size: 8, name: "Calibri", color: { argb: "FF334155" } };
                const fontData = { size: 10, name: "Calibri", color: { argb: "FF0F172A" } };

                const wb = new ExcelJS.Workbook();
                wb.creator = "UTILITY SENSE";
                const sheetName = String(plantLabel).replace(/[\\/?*[\]]/g, "").slice(0, 28) || "Utility";
                const ws = wb.addWorksheet(sheetName, { views: [{ state: "frozen", ySplit: 7 }] });
                ws.columns = [
                    { width: 12 }, { width: 10 }, { width: 14 }, { width: 14 }, { width: 12 },
                    { width: 12 }, { width: 12 }, { width: 11 }, { width: 14 }, { width: 14 },
                    { width: 11 }, { width: 12 }, { width: 13 }, { width: 14 }, { width: 8 },
                    { width: 8 }, { width: 11 }, { width: 13 },
                ];

                ws.mergeCells(1, 1, 1, COLS);
                const titleCell = ws.getCell(1, 1);
                titleCell.value = "UTILITY SENSE — Daily Utility Consumption Report";
                titleCell.fill = fillTitle;
                titleCell.font = fontWhiteBold;
                titleCell.alignment = center;
                titleCell.border = borderAll;
                ws.getRow(1).height = 28;

                ws.mergeCells(2, 1, 2, 2);
                ws.getCell(2, 1).value = "LOCATION";
                ws.getCell(2, 1).font = fontBold;
                ws.getCell(2, 1).fill = fillMeta;
                ws.mergeCells(2, 3, 2, 6);
                ws.getCell(2, 3).value = locationLabel;
                ws.getCell(2, 3).font = { ...fontBold, size: 12, color: { argb: "FF0369A1" } };
                ws.getCell(2, 3).fill = fillMeta;
                ws.mergeCells(2, 7, 2, 8);
                ws.getCell(2, 7).value = "PLANT";
                ws.getCell(2, 7).font = fontBold;
                ws.getCell(2, 7).fill = fillMeta;
                ws.mergeCells(2, 9, 2, COLS);
                ws.getCell(2, 9).value = plantLabel;
                ws.getCell(2, 9).font = { ...fontBold, size: 12, color: { argb: "FF0369A1" } };
                ws.getCell(2, 9).fill = fillMeta;
                for (let c = 1; c <= COLS; c++) {
                    ws.getCell(2, c).border = borderAll;
                    ws.getCell(2, c).alignment = center;
                    if (!ws.getCell(2, c).fill) ws.getCell(2, c).fill = fillMeta;
                }
                ws.getRow(2).height = 22;

                ws.mergeCells(3, 1, 3, COLS);
                ws.getCell(3, 1).value = `Generated: ${new Date().toLocaleString("en-IN")}  ·  Rows: ${list.length}  ·  Rate: ${gridLabel}/Solar ₹${electRate}/unit  ·  Diesel ₹${dieselRate}/L`;
                ws.getCell(3, 1).font = fontSmall;
                ws.getCell(3, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
                ws.getCell(3, 1).alignment = { vertical: "middle", horizontal: "left" };
                for (let c = 1; c <= COLS; c++) ws.getCell(3, c).border = borderAll;
                ws.getRow(3).height = 18;
                ws.getRow(4).height = 6;

                PLANT_REPORT_TYPE_ROW.forEach((v, i) => {
                    const cell = ws.getCell(5, i + 1);
                    cell.value = v;
                    cell.font = fontBold;
                    cell.alignment = center;
                    cell.border = borderAll;
                    cell.fill = v === "Manual" ? fillYellow : fillGreen;
                });
                ws.getRow(5).height = 18;

                (reportFormulaRow || []).forEach((v, i) => {
                    const cell = ws.getCell(6, i + 1);
                    cell.value = v;
                    cell.font = fontSmall;
                    cell.alignment = center;
                    cell.border = borderAll;
                    cell.fill = MANUAL_COLS.has(i + 1) ? fillYellow : fillBlue;
                });
                ws.getRow(6).height = 32;

                (reportHeaders || []).forEach((v, i) => {
                    const cell = ws.getCell(7, i + 1);
                    cell.value = v;
                    cell.font = fontBold;
                    cell.alignment = center;
                    cell.border = borderAll;
                    cell.fill = MANUAL_COLS.has(i + 1) ? fillYellow : fillBlue;
                });
                ws.getRow(7).height = 36;

                dataRows.forEach((values, idx) => {
                    const r = 8 + idx;
                    values.forEach((v, i) => {
                        const cell = ws.getCell(r, i + 1);
                        cell.value = v;
                        cell.font = fontData;
                        cell.alignment = center;
                        cell.border = borderAll;
                        if (MANUAL_COLS.has(i + 1)) {
                            cell.fill = fillYellow;
                        } else if (idx % 2 === 1) {
                            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
                        }
                    });
                    ws.getRow(r).height = 18;
                });

                return wb;
            };

            /** Shared filtered rows + labels for Excel export & Print (same content) */
            const buildPlantReportData = (customScope = null) => {
                const targetLoc = customScope?.location ?? selectedReportLocation;
                const targetPlant = customScope?.plant ?? filters.plant;
                const targetStart = customScope?.startDate ?? (reportRangeMode === "custom" && reportCustomStartDate ? reportCustomStartDate : filters.startDate);
                const targetEnd = customScope?.endDate ?? (reportRangeMode === "custom" && reportCustomEndDate ? reportCustomEndDate : filters.endDate);

                const list = [...dailyEntries]
                    .filter((e) => {
                        const meta = resolvePlantMeta(e.plant);
                        const ep = entryPlantKey(e.plant);
                        const fp = entryPlantKey(targetPlant);

                        // Plant filter check
                        if (targetPlant && targetPlant !== "all") {
                            const matchPlant =
                                ep === fp ||
                                String(meta.code).toLowerCase() === fp ||
                                String(meta.name).toLowerCase() === fp ||
                                String(e.plant).toLowerCase() === fp;
                            if (!matchPlant) return false;
                        }

                        // Location filter check
                        if (targetLoc && targetLoc !== "all") {
                            const loc = String(e.location || meta.location || "").toUpperCase();
                            const selLoc = String(targetLoc).toUpperCase();
                            const matchLoc =
                                loc === selLoc ||
                                loc.includes(selLoc) ||
                                selLoc.includes(loc);
                            if (!matchLoc) return false;
                        }
                        if (targetStart && e.date < targetStart) return false;
                        if (targetEnd && e.date > targetEnd) return false;
                        return true;
                    })
                    .sort((a, b) => a.date.localeCompare(b.date) || String(a.plant).localeCompare(String(b.plant)));

                // If explicit customScope was passed, only return list if it matches (no fallback to all)
                const finalRows = customScope ? list : (list.length > 0 ? list : filteredEntries);
                if (finalRows.length === 0) return null;

                const plantFilterMeta = targetPlant !== "all" ? resolvePlantMeta(targetPlant) : null;
                const locationLabel =
                    targetLoc !== "all"
                        ? targetLoc
                        : plantFilterMeta?.location || (finalRows.length === 1 ? resolvePlantMeta(finalRows[0].plant).location : "All Locations");
                const plantLabel =
                    plantFilterMeta?.name ||
                    (targetPlant !== "all" ? targetPlant : finalRows.length === 1 ? resolvePlantMeta(finalRows[0].plant).name : "All Plants");

                const electRate = Math.round((Number(activeElectRate) || 10.89) * 100) / 100;
                const solarRate = Number(activeSolarRate) || 10.89;
                const dieselRate = Number(activeDieselRate) || 90.62;

                const primaryLoc = locationLabel !== "All Locations"
                    ? locationLabel
                    : (finalRows[0] ? (finalRows[0].location || resolvePlantMeta(finalRows[0].plant).location) : "PUNE");
                const primaryPlant = targetPlant !== "all" ? targetPlant : (finalRows[0]?.plant || "");
                const gridLabel = getGridProviderLabel(primaryLoc);
                const reportMf = resolveMultiplyFactor(multiplyFactors, primaryPlant, primaryLoc, finalRows[finalRows.length - 1]?.date);
                const { headers: reportHeaders, formulaRow: reportFormulaRow } = buildPlantReportColumns(gridLabel, reportMf);

                const dataRows = finalRows.map((e) => {
                    const meta = resolvePlantMeta(e.plant);
                    const loc = e.location || meta.location;
                    const rowMf = resolveMultiplyFactor(multiplyFactors, e.plant, loc, e.date);
                    const dailyReading = Number(e.electricity_closing) || 0;
                    const opening = Number(e.electricity_opening) || 0;
                    const unitsDiff = Math.max(0, dailyReading - opening);
                    const mseb = Number(e.electricity_consumption) || CalculationEngine.calculateMSEBUnits(unitsDiff, rowMf);
                    const rowElectRate = resolveTariff(tariffs, "electricity", e.plant, loc, e.date);
                    const rowSolarRate = resolveTariff(tariffs, "solar", e.plant, loc, e.date) || rowElectRate;
                    const rowDieselRate = resolveTariff(tariffs, "diesel", e.plant, loc, e.date);
                    const solar = Number(e.solar_generated) || 0;
                    const totalUnits = mseb + solar;
                    const msebCost = Number(e.electricity_cost) || CalculationEngine.calculateElectricityCost(mseb, rowElectRate);
                    const solarCost = Number(e.solar_cost) || solar * rowSolarRate;
                    const dieselL = Number(e.diesel_used) || 0;
                    const dieselCost = Number(e.diesel_cost) || dieselL * rowDieselRate;
                    const totalCost = Number(e.total_cost) || msebCost + solarCost + dieselCost;
                    const odu = Number(e.odu) || 0;
                    const idu = Number(e.idu) || 0;
                    const prodSets = Number(e.production_qty) || CalculationEngine.calculateProductionSets(odu, idu) || 0;
                    const costPerSet = prodSets > 0
                        ? CalculationEngine.calculateCostPerSet(totalCost, prodSets)
                        : 0;

                    return [
                        formatExportDate(e.date),
                        meta.name || e.plant,
                        dailyReading || "",
                        Math.round(unitsDiff * 1000) / 1000,
                        mseb,
                        solar,
                        totalUnits,
                        `₹ ${Math.round(rowElectRate * 100) / 100}`,
                        fmtInrExport(msebCost),
                        fmtInrExport(solarCost),
                        dieselL > 0 || dieselCost > 0 ? `₹ ${rowDieselRate}` : "",
                        dieselL || "",
                        dieselL ? fmtInrExport(dieselCost) : "",
                        fmtInrExport(totalCost),
                        odu || "—",
                        idu || "—",
                        prodSets || "—",
                        prodSets ? fmtInrExport(costPerSet) : "₹ 0",
                    ];
                });

                return {
                    list: finalRows, locationLabel, plantLabel, electRate, dieselRate, dataRows,
                    gridLabel, reportMf, reportHeaders, reportFormulaRow,
                };
            };

            /** Colourful Excel export matching plant tracker sheet */
            const handleExportPlantExcel = async () => {
                try {
                    // 1. Authorization check
                    if (currentUser && currentUser.role !== "IT_ADMIN") {
                        if (selectedReportLocation !== "all" && !allowedLocations.map(x => x.toUpperCase()).includes(selectedReportLocation.toUpperCase())) {
                            setToast({ type: "error", message: "Unauthorized location export." });
                            return;
                        }
                        if (filters.plant !== "all" && !allowedPlants.some(p => p.plant_code === filters.plant)) {
                            setToast({ type: "error", message: "Unauthorized plant export." });
                            return;
                        }
                    }

                    // 2. Validate range
                    let startDate = "";
                    let endDate = "";
                    if (reportRangeMode === "month") {
                        if (reportFromMonth && reportToMonth && reportFromMonth > reportToMonth) {
                            setToast({ type: "error", message: "Invalid date range: 'From Month' cannot be after 'To Month'." });
                            return;
                        }
                        startDate = reportFromMonth ? `${reportFromMonth}-01` : "";
                        endDate = reportToMonth ? getMonthEnd(reportToMonth) : "";
                    } else {
                        startDate = reportCustomStartDate || filters.startDate;
                        endDate = reportCustomEndDate || filters.endDate;
                        if (startDate && endDate && startDate > endDate) {
                            setToast({ type: "error", message: "Invalid date range: 'From Date' cannot be after 'To Date'." });
                            return;
                        }
                    }

                    const payload = buildPlantReportData({
                        location: selectedReportLocation,
                        plant: filters.plant,
                        startDate,
                        endDate
                    });

                    if (!payload || !payload.list || payload.list.length === 0) {
                        setToast({ type: "warning", message: "No data found for the selected filters." });
                        return;
                    }
                    const { list, locationLabel, plantLabel } = payload;

                    const wb = createPlantReportWorkbook(payload);
                    const buffer = await wb.xlsx.writeBuffer();
                    const blob = new Blob([buffer], {
                        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    });
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement("a");
                    const filePlant = String(plantLabel).replace(/\s+/g, "_");
                    const fileLoc = String(locationLabel).replace(/\s+/g, "_");
                    link.href = url;
                    link.download = `UtilitySense_${fileLoc}_${filePlant}_${new Date().toISOString().split("T")[0]}.xlsx`;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    URL.revokeObjectURL(url);

                    recordAuditLog({
                        action: "EXPORT",
                        module: "Reports",
                        recordId: `${fileLoc}_${filePlant}`,
                        location: locationLabel,
                        plant: plantLabel,
                        newValue: { rowCount: list.length, range: `${startDate} to ${endDate}` },
                        status: "SUCCESS"
                    });

                    setToast({
                        type: "success",
                        message: `Excel exported successfully: ${locationLabel} · ${plantLabel} (${list.length} rows)`,
                    });
                } catch (err) {
                    console.error("Export Excel Error:", err);
                    setToast({ type: "error", message: `Excel export error: ${err.message || err}` });
                }
            };

            const [plantPrintPayload, setPlantPrintPayload] = useState(null);

            /** Print EXACT same colourful plant report as Export Excel (not other report tables) */
            const handlePrintPlantReport = () => {
                let startDate = "";
                let endDate = "";
                if (reportRangeMode === "month") {
                    startDate = reportFromMonth ? `${reportFromMonth}-01` : "";
                    endDate = reportToMonth ? getMonthEnd(reportToMonth) : "";
                } else {
                    startDate = reportCustomStartDate || filters.startDate;
                    endDate = reportCustomEndDate || filters.endDate;
                }
                const payload = buildPlantReportData({
                    location: selectedReportLocation,
                    plant: filters.plant,
                    startDate,
                    endDate
                });
                if (!payload || !payload.list || payload.list.length === 0) {
                    setToast({ type: "error", message: "Print ke liye koi entry nahi mili. Filters check karo." });
                    return;
                }
                setPlantPrintPayload(payload);
                // Wait for React to paint the print sheet, then print only that
                requestAnimationFrame(() => {
                    requestAnimationFrame(() => {
                        document.body.classList.add("printing-plant-excel");
                        const cleanup = () => {
                            document.body.classList.remove("printing-plant-excel");
                            setPlantPrintPayload(null);
                            window.removeEventListener("afterprint", cleanup);
                        };
                        window.addEventListener("afterprint", cleanup);
                        window.print();
                        setTimeout(cleanup, 1500);
                    });
                });
            };

            const handleExportCSV = () => {
                handleExportPlantExcel();
            };

            // ----------------------------------------------------
            // INTRO ANIMATION (before login)
            // ----------------------------------------------------
            if (!introDone) {
                return <IntroSplash onDone={finishIntro} logoSrc={PG_LOGO_BASE_64} />;
            }

            // ----------------------------------------------------
            // REDIRECT TO LOGIN GATE IF NOT AUTHENTICATED
            // ----------------------------------------------------
            if (!currentUser) {
                return (
                    <>
                    <div
                        className="min-h-screen flex items-center justify-center p-4 relative overflow-hidden bg-slate-100"
                        style={{
                            backgroundImage: `radial-gradient(circle at center, rgba(248, 250, 252, 0.45) 0%, rgba(226, 232, 240, 0.75) 100%), url('/login-bg.png')`,
                            backgroundSize: 'cover',
                            backgroundPosition: 'center',
                            backgroundRepeat: 'no-repeat',
                            display: 'flex',
                            justifyContent: 'center',
                            alignItems: 'center'
                        }}
                    >
                        {/* Soft backdrop blur */}
                        <div className="absolute inset-0 bg-slate-900/10 backdrop-blur-[2px] pointer-events-none" />

                        <div
                            className="w-full max-w-md bg-white/95 backdrop-blur-md rounded-3xl p-8 border border-slate-200/90 shadow-[0_25px_60px_rgba(0,0,0,0.18)] z-10"
                            style={{ position: 'relative' }}
                        >
                            <div className="flex flex-col items-center text-center mb-6">
                                <div className="h-14 w-14 mb-2 flex items-center justify-center">
                                    <img
                                        src={PG_LOGO_BASE_64}
                                        alt="PG Electroplast Ltd"
                                        className="max-h-full max-w-full object-contain"
                                    />
                                </div>
                                <h1 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-1.5 justify-center">
                                    <span className="text-sky-600">Utility</span><span>Sense</span>
                                </h1>
                                <p className="text-[10.5px] font-bold text-slate-400 mt-1 uppercase tracking-wider">
                                    Corporate Energy & Resource Governance Hub
                                </p>
                            </div>

                            {loginError && (
                                <div className="mb-4 flex items-center gap-2 rounded-xl bg-red-50 border border-red-100 p-3.5 text-xs text-red-700">
                                    <span className="material-symbols-outlined text-[16px] text-red-500">warning</span>
                                    <span className="font-semibold">{loginError}</span>
                                </div>
                            )}

                            {loginMessage && (
                                <div className="mb-4 flex items-center gap-2 rounded-xl bg-emerald-50 border border-emerald-100 p-3.5 text-xs text-emerald-800">
                                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse shrink-0" />
                                    <span className="font-semibold">{loginMessage}</span>
                                </div>
                            )}

                            {!otpSent ? (
                                <form onSubmit={handleSendOTP} className="space-y-4">
                                    <div>
                                        <label className="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Registered Corporate Email</label>
                                        <div className="relative">
                                            <input
                                                type="email"
                                                required
                                                placeholder="employee@company.com"
                                                value={loginEmail}
                                                onChange={(e) => setLoginEmail(e.target.value)}
                                                className="w-full h-11 rounded-xl border border-slate-200 pl-10 pr-4 text-xs bg-slate-50 focus:outline-none focus:ring-1 focus:ring-sky-500"
                                            />
                                            <span className="material-symbols-outlined absolute left-3.5 top-3.5 text-[18px] text-slate-400">mail</span>
                                        </div>
                                    </div>

                                    <button
                                        type="submit"
                                        disabled={loginLoading}
                                        className="w-full h-11 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl text-xs font-bold transition flex items-center justify-center gap-2 disabled:opacity-50 border-none cursor-pointer"
                                    >
                                        {loginLoading && <span className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />}
                                        <span>Send Verification OTP</span>
                                    </button>

                                </form>
                            ) : (
                                <form onSubmit={handleVerifyOTP} className="space-y-4">
                                    <div>
                                        <label className="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Verification Code (OTP)</label>
                                        <div className="relative">
                                            <input
                                                type="text"
                                                required
                                                maxLength="6"
                                                placeholder="Enter 6-digit code"
                                                value={loginOtp}
                                                onChange={(e) => setLoginOtp(e.target.value.replace(/\D/g, ''))}
                                                className="w-full h-11 rounded-xl border border-slate-200 pl-10 pr-4 text-center tracking-[0.3em] text-sm font-bold bg-slate-50 focus:outline-none focus:ring-1 focus:ring-sky-500"
                                            />
                                            <span className="material-symbols-outlined absolute left-3.5 top-3.5 text-[18px] text-slate-400">lock</span>
                                        </div>
                                    </div>

                                    <button
                                        type="submit"
                                        disabled={loginLoading}
                                        className="w-full h-11 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl text-xs font-bold transition flex items-center justify-center gap-2 disabled:opacity-50 border-none cursor-pointer"
                                    >
                                        {loginLoading && <span className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />}
                                        <span>Verify & Access System</span>
                                    </button>

                                    <div className="flex items-center justify-between text-[11px] font-semibold text-slate-500 pt-2">
                                        <button type="button" onClick={() => { setOtpSent(false); setLoginOtp(""); }} className="hover:text-[#0284c7] bg-transparent border-none cursor-pointer">
                                            Change Email
                                        </button>
                                        <button type="button" onClick={handleSendOTP} className="hover:text-[#0284c7] bg-transparent border-none cursor-pointer">
                                            Resend OTP
                                        </button>
                                    </div>
                                </form>
                            )}
                        </div>
                    </div>
                    <IdleScreensaver logoSrc={PG_LOGO_BASE_64} idleMinutes={idleMinutes} />
                </>
                );
            }

            // ----------------------------------------------------
            // SYNCING SPINNER
            // ----------------------------------------------------
            if (loading) {
                return (
                    <>
                    <div className="min-h-screen flex flex-col items-center justify-center p-6 text-sm text-slate-500 bg-white">
                        <div className="w-10 h-10 border-4 border-[#0284c7] border-t-transparent rounded-full animate-spin mb-4"></div>
                        Syncing database...
                    </div>
                    <IdleScreensaver logoSrc={PG_LOGO_BASE_64} idleMinutes={idleMinutes} />
                    </>
                );
            }

            // ----------------------------------------------------
            // ERROR STATE GATE
            // ----------------------------------------------------
            if (error) {
                return (
                    <>
                    <div className="min-h-screen flex items-center justify-center p-6 bg-red-50 text-red-700 text-sm">
                        <div className="max-w-md text-center bg-white p-8 rounded-3xl border border-red-100 shadow-xl">
                            <span className="material-symbols-outlined text-[36px] text-red-500 mb-2">error</span>
                            <p className="font-bold text-slate-900 text-base mb-1">Spreadsheet Connection Failed</p>
                            <p className="text-xs text-slate-500 mb-6">{error}</p>
                            <button
                                onClick={() => handleLogout("Database connection failed.")}
                                className="px-6 py-2 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold border-none cursor-pointer transition"
                            >
                                Log Out Profile
                            </button>
                        </div>
                    </div>
                    <IdleScreensaver logoSrc={PG_LOGO_BASE_64} idleMinutes={idleMinutes} />
                    </>
                );
            }

            return (
                <>
                <div className="min-h-screen flex flex-col bg-white">
                    {/* Header Panel */}
                    <div className="sticky top-0 z-30 bg-white border-b border-slate-100 shadow-sm shrink-0 no-print" style={{ height: '64px' }}>
                        <div className="h-full px-5 flex items-center justify-between gap-2.5">
                            {/* 1. Left Group: Logo & Title & Windmill */}
                            <div className="flex items-center gap-2 shrink-0">
                                {/* Stylized PG Electroplast Circular Logo */}
                                <div className="flex items-center select-none shrink-0 pr-1">
                                    <img src={PG_LOGO_BASE_64} className="h-11 w-auto object-contain" alt="PG Electroplast Logo" />
                                </div>
                                
                                {/* UTILITY SENSE Title Block */}
                                <div className="flex flex-col justify-center shrink-0">
                                    <h1 className="text-base sm:text-lg font-black tracking-tight text-slate-900 dark:text-white uppercase leading-none">
                                        UTILITY SENSE
                                    </h1>
                                    <p className="text-[10px] sm:text-[11px] font-extrabold text-emerald-600 dark:text-emerald-400 leading-tight mt-1 whitespace-nowrap">
                                        {(() => {
                                            if (filters.plant && filters.plant !== "all") {
                                                const p = plants.find(pl => pl.plant_code === filters.plant);
                                                if (p) return `${p.plant_display_name || p.plant_name} (${p.location})`;
                                            }
                                            if (filters.location && filters.location !== "all") {
                                                return `${filters.location} · All Plants`;
                                            }
                                            return "All Locations · All Plants";
                                        })()}
                                    </p>
                                </div>

                                {/* Rotating Windmill SVG - Repositioned to the right of title text with clean separation */}
                                <div className="ml-3.5 mr-1.5 flex items-center justify-center shrink-0 border-r border-slate-200/80 pr-2.5">
                                    <svg viewBox="0 0 100 100" className="w-9 h-9 text-[#0284c7] dark:text-[#38bdf8]">
                                        <defs>
                                            <g id="wind-blade">
                                                {/* Opaque white blade body with solid blue outline */}
                                                <path d="M 49.7,40 C 49.5,33, 49.0,22, 49.3,8 C 49.7,6, 50.3,6, 50.7,8 C 51.0,22, 50.5,33, 50.3,40 Z" fill="#ffffff" stroke="currentColor" strokeWidth="0.8" />
                                                {/* Red stripe 1 (near tip) */}
                                                <path d="M 49.4,14 L 50.6,14 L 50.5,12 L 49.5,12 Z" fill="#ef4444" />
                                                {/* Red stripe 2 (at the very tip) */}
                                                <path d="M 49.3,10 L 50.7,10 C 50.6,8.2, 50.3,7.8, 50,7.8 C 49.7,7.8, 49.4,8.2, 49.3,10 Z" fill="#ef4444" />
                                            </g>
                                        </defs>
                                        
                                        {/* Tower (Tapered base) - white fill with solid blue outline */}
                                        <path d="M 48,88 L 52,88 L 50.6,40 L 49.4,40 Z" fill="#ffffff" stroke="currentColor" strokeWidth="0.8" />
                                        
                                        {/* Nacelle (Generator body) - solid blue fill */}
                                        <path d="M 48,38 C 48,36.5 52,36.5 52,38 L 51,41 L 49,41 Z" fill="currentColor" />
                                        
                                        {/* Rotating 3 Blades */}
                                        <g className="windmill-blades">
                                            <use href="#wind-blade" />
                                            <use href="#wind-blade" transform="rotate(120, 50, 40)" />
                                            <use href="#wind-blade" transform="rotate(240, 50, 40)" />
                                        </g>
                                        
                                        {/* Center Hub Cap - solid blue fill with white reflection dot */}
                                        <circle cx="50" cy="40" r="2" fill="currentColor" />
                                        <circle cx="49.5" cy="39.5" r="0.6" fill="#ffffff" />
                                    </svg>
                                </div>
                            </div>

                            {/* 2. Middle Group: Navigation Tabs (Icons ONLY) */}
                            <div className="flex items-center gap-1 bg-slate-100/70 dark:bg-slate-800/50 p-0.5 rounded-full border border-slate-200/50 h-[34px] shrink-0">
                                <button type="button" onClick={() => setActiveTab("dashboard")} className={`h-[28px] w-[28px] rounded-full transition flex items-center justify-center cursor-pointer border-none shrink-0 ${activeTab === 'dashboard' ? 'bg-white text-[#0284c7] shadow-sm dark:bg-slate-700 dark:text-white' : 'bg-transparent text-slate-500 hover:text-slate-700'}`} title="Dashboard">
                                    <span className="material-symbols-outlined text-[18px]">dashboard</span>
                                </button>

                                <button type="button" onClick={() => setActiveTab("reports")} className={`h-[28px] w-[28px] rounded-full transition flex items-center justify-center cursor-pointer border-none shrink-0 ${activeTab === 'reports' ? 'bg-white text-[#0284c7] shadow-sm dark:bg-slate-700 dark:text-white' : 'bg-transparent text-slate-500 hover:text-slate-700'}`} title="Reports">
                                    <span className="material-symbols-outlined text-[18px]">description</span>
                                </button>

                                {currentUser.role === "IT_ADMIN" && (
                                    <React.Fragment>
                                        <button type="button" onClick={() => setActiveTab("targets")} className={`h-[28px] w-[28px] rounded-full transition flex items-center justify-center cursor-pointer border-none shrink-0 ${activeTab === 'targets' ? 'bg-white text-[#0284c7] shadow-sm dark:bg-slate-700 dark:text-white' : 'bg-transparent text-slate-500 hover:text-slate-700'}`} title="Targets">
                                            <span className="material-symbols-outlined text-[18px]">track_changes</span>
                                        </button>

                                        <button type="button" onClick={() => setActiveTab("alerts")} className={`h-[28px] w-[28px] rounded-full transition flex items-center justify-center cursor-pointer border-none shrink-0 relative ${activeTab === 'alerts' ? 'bg-white text-[#0284c7] shadow-sm dark:bg-slate-700 dark:text-white' : 'bg-transparent text-slate-500 hover:text-slate-700'}`} title="Alerts">
                                            <span className="material-symbols-outlined text-[18px]">notifications_active</span>
                                            {alertsList.length > 0 && (
                                                <span className="absolute -top-0.5 -right-0.5 px-1 py-0.2 rounded-full text-[8px] font-bold bg-rose-500 text-white leading-none">{alertsList.length}</span>
                                            )}
                                        </button>

                                        {/* Shifted Monitor Toggle Icon directly after Alerts */}
                                        {activeTab === "dashboard" && (
                                            <button
                                                type="button"
                                                onClick={handleToggleCanMonitor}
                                                className={`h-[28px] w-[28px] rounded-full border-none transition flex items-center justify-center cursor-pointer shrink-0 ${
                                                    canMonitor
                                                        ? 'bg-emerald-50 text-emerald-700 shadow-xs'
                                                        : 'bg-transparent text-slate-500 hover:text-slate-700'
                                                }`}
                                                title="Toggle Monitoring for Air, LPG, and Waste"
                                            >
                                                <span className="material-symbols-outlined text-[18px]">{canMonitor ? 'visibility' : 'visibility_off'}</span>
                                            </button>
                                        )}

                                        {/* Shifted Filters Toggle Icon directly after Monitor */}
                                        {activeTab === "dashboard" && (
                                            <button
                                                ref={filterToggleBtnRef}
                                                type="button"
                                                onClick={() => setFiltersOpen(o => !o)}
                                                className={`h-[28px] w-[28px] rounded-full border-none transition flex items-center justify-center cursor-pointer shrink-0 ${
                                                    filtersOpen
                                                        ? 'bg-sky-50 text-sky-700 shadow-xs'
                                                        : 'bg-transparent text-slate-500 hover:text-slate-700'
                                                }`}
                                                title="Toggle Filters"
                                            >
                                                <span className="material-symbols-outlined text-[18px]">tune</span>
                                            </button>
                                        )}

                                        <button type="button" onClick={() => { setActiveTab("master"); setSelectedMasterTable("plants"); }} className={`h-[28px] w-[28px] rounded-full transition flex items-center justify-center cursor-pointer border-none shrink-0 ${activeTab === 'master' ? 'bg-white text-[#0284c7] shadow-sm dark:bg-slate-700 dark:text-white' : 'bg-transparent text-slate-500 hover:text-slate-700'}`} title="Master Settings">
                                            <span className="material-symbols-outlined text-[18px]">settings</span>
                                        </button>
                                    </React.Fragment>
                                )}

                                {currentUser.role !== "IT_ADMIN" && (
                                    <React.Fragment>
                                        {activeTab === "dashboard" && (
                                            <button
                                                type="button"
                                                onClick={handleToggleCanMonitor}
                                                className={`h-[28px] w-[28px] rounded-full border-none transition flex items-center justify-center cursor-pointer shrink-0 ${
                                                    canMonitor
                                                        ? 'bg-emerald-50 text-emerald-700 shadow-xs'
                                                        : 'bg-transparent text-slate-500 hover:text-slate-700'
                                                }`}
                                                title="Toggle Monitoring for Air, LPG, and Waste"
                                            >
                                                <span className="material-symbols-outlined text-[18px]">{canMonitor ? 'visibility' : 'visibility_off'}</span>
                                            </button>
                                        )}

                                        {activeTab === "dashboard" && (
                                            <button
                                                ref={filterToggleBtnRef}
                                                type="button"
                                                onClick={() => setFiltersOpen(o => !o)}
                                                className={`h-[28px] w-[28px] rounded-full border-none transition flex items-center justify-center cursor-pointer shrink-0 ${
                                                    filtersOpen
                                                        ? 'bg-sky-50 text-sky-700 shadow-xs'
                                                        : 'bg-transparent text-slate-500 hover:text-slate-700'
                                                }`}
                                                title="Toggle Filters"
                                            >
                                                <span className="material-symbols-outlined text-[18px]">tune</span>
                                            </button>
                                        )}

                                        <button type="button" onClick={() => { setActiveTab("master"); setSelectedMasterTable("tariff_rates"); }} className={`h-[28px] w-[28px] rounded-full transition flex items-center justify-center cursor-pointer border-none shrink-0 ${activeTab === 'master' ? 'bg-white text-[#0284c7] shadow-sm dark:bg-slate-700 dark:text-white' : 'bg-transparent text-slate-500 hover:text-slate-700'}`} title="Master Settings">
                                            <span className="material-symbols-outlined text-[18px]">settings</span>
                                        </button>
                                    </React.Fragment>
                                )}
                            </div>

                            {/* 3. Right Group: Actions & User details (ml-auto shrink-0) */}
                            <div className="flex items-center gap-2.5 ml-auto shrink-0">
                                {activeTab === "dashboard" && (
                                    <button
                                        type="button"
                                        onClick={() => setActiveTab("entry")}
                                        className="h-[28px] px-3.5 rounded-full bg-[#0284c7] hover:bg-[#0369a1] text-white font-bold text-xs transition flex items-center gap-1 border-none cursor-pointer shadow-sm"
                                        title="Operations Logging"
                                    >
                                        <span className="material-symbols-outlined text-[16px]">add</span>
                                        <span>Daily Entry</span>
                                    </button>
                                )}

                                <div className="flex items-center gap-3 text-xs font-semibold text-slate-700 shrink-0">
                                    <div className="hidden lg:block text-right leading-tight">
                                        <p className="text-xs font-bold text-slate-900">{currentUser.name}</p>
                                        <p className="text-[9px] text-slate-400 font-medium font-mono">{currentUser.email}</p>
                                    </div>

                                    {/* Circular profile initials badge matching VEMS */}
                                    <div className="w-9 h-9 bg-sky-600 text-white rounded-full flex items-center justify-center font-bold text-xs uppercase shadow-sm select-none" title={`${currentUser.name} (${currentUser.email})`}>
                                        {currentUser.name ? currentUser.name.split(' ').map(n => n[0]).join('').toUpperCase() : 'U'}
                                    </div>

                                    {/* view layout toggle */}
                                    <button
                                        onClick={() => setKpiLayout(kpiLayout === "scroll" ? "grid" : "scroll")}
                                        title={kpiLayout === "scroll" ? "Switch to Stacked Grid View" : "Switch to Scrollable Row View"}
                                        className="h-9 w-9 rounded-full bg-slate-50 border border-slate-200 hover:bg-slate-100 flex items-center justify-center text-slate-500 hover:text-slate-700 cursor-pointer transition"
                                    >
                                        <span className="material-symbols-outlined text-[18px]">
                                            {kpiLayout === "scroll" ? "splitscreen" : "view_week"}
                                        </span>
                                    </button>

                                    <button
                                        onClick={() => handleLogout("Session signed out.")}
                                        title="Logout Session"
                                        className="h-9 w-9 rounded-full bg-slate-50 border border-slate-200 hover:bg-slate-100 flex items-center justify-center text-slate-500 hover:text-slate-700 cursor-pointer transition"
                                    >
                                        <span className="material-symbols-outlined text-[18px]">logout</span>
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>



                    {/* MAIN CONTENT DISPLAY */}
                    <main className={`main-content${activeTab === "reports" ? " main-content--reports" : ""}`}>

                        {/* 1. DASHBOARD COMPONENT */}
                        {activeTab === "dashboard" && (
                            <div className="flex flex-col gap-1.5 pt-1">

                                {/* Filters panel (collapsible — toggled via the "Filters" button in the top nav) */}
                                {filtersOpen && (
                                    <div ref={filterPanelRef} className="no-print bg-slate-50 border border-slate-200/60 p-2 px-3 rounded-xl flex flex-wrap items-center justify-between gap-3 shadow-xs mt-0 mb-1">
                                        <div className="flex flex-wrap items-center gap-2.5">
                                            <div className="flex items-center gap-1.5 bg-white border border-slate-200 rounded-lg px-2 py-1 shadow-xs">
                                                <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500 shrink-0">Start:</label>
                                                <input
                                                    type="date"
                                                    value={filters.startDate}
                                                    onChange={(e) => setFilters({ ...filters, startDate: e.target.value })}
                                                    className="h-6 border-none bg-transparent text-[11.5px] text-slate-800 font-bold focus:outline-none p-0"
                                                />
                                            </div>
                                            <div className="flex items-center gap-1.5 bg-white border border-slate-200 rounded-lg px-2 py-1 shadow-xs">
                                                <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500 shrink-0">End:</label>
                                                <input
                                                    type="date"
                                                    value={filters.endDate}
                                                    onChange={(e) => setFilters({ ...filters, endDate: e.target.value })}
                                                    className="h-6 border-none bg-transparent text-[11.5px] text-slate-800 font-bold focus:outline-none p-0"
                                                />
                                            </div>
                                            <div className="flex items-center gap-1.5 bg-white border border-slate-200 rounded-lg px-2 py-1 shadow-xs">
                                                <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500 shrink-0">Location:</label>
                                                <select
                                                    value={filters.location || "all"}
                                                    onChange={(e) => handleLocationFilterChange(e.target.value)}
                                                    className="h-6 border-none bg-transparent text-[11.5px] text-slate-800 font-bold focus:outline-none p-0 cursor-pointer"
                                                >
                                                    <option value="all">All Locations</option>
                                                    {reportLocations.map(loc => (
                                                        <option key={loc} value={loc}>{loc}</option>
                                                    ))}
                                                </select>
                                            </div>
                                            <div className="flex items-center gap-1.5 bg-white border border-slate-200 rounded-lg px-2 py-1 shadow-xs">
                                                <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500 shrink-0">Plant:</label>
                                                <select
                                                    value={filters.plant}
                                                    onChange={(e) => setFilters({ ...filters, plant: e.target.value })}
                                                    className="h-6 border-none bg-transparent text-[11.5px] text-slate-800 font-bold focus:outline-none p-0 cursor-pointer max-w-[160px] truncate"
                                                >
                                                    <option value="all">All Plants</option>
                                                    {filteredPlantsForDropdown.map(p => (
                                                        <option key={p.plant_code} value={p.plant_code}>{p.plant_name}</option>
                                                    ))}
                                                </select>
                                            </div>
                                        </div>

                                        <button
                                            onClick={() => setFilters(getDefaultDateFilters())}
                                            className="h-7 px-2.5 bg-white border border-slate-200 hover:bg-slate-100 rounded-lg text-[11px] font-bold text-slate-600 transition flex items-center gap-1 cursor-pointer shadow-xs"
                                            title="Reset Filters"
                                        >
                                            <span className="material-symbols-outlined text-[14px]">restart_alt</span>
                                            <span>Reset</span>
                                        </button>
                                    </div>
                                )}

                                 {/* 13 KPI Metric Row — locked to exactly one line, frozen (sticky) below the top nav while the page scrolls */}
                                {kpiLayout === "grid" ? (
                                    <div className="w-full flex flex-col gap-2 no-print" style={{ position: 'sticky', top: '64px', zIndex: 15, background: 'var(--bg)', paddingTop: '0px', paddingBottom: '4px' }}>
                                        {/* Row 1: 8 Primary Cards */}
                                        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2">
                                            <KpiCard label="Electricity (kWh)" value={fmtNum(kpiTotals.electricity)} icon="electric_bolt" tone="blue" compact={true} />
                                            <KpiCard label="Solar Gen (kWh)" value={fmtNum(kpiTotals.solarGenerated)} icon="wb_sunny" tone="amber" compact={true} />
                                            <KpiCard label="Total Consumption" value={fmtNum(aggregatedCosts.totalConsumption)} sub={`${dashboardGridLabel} + Solar (filter)`} icon="bolt" tone="orange" compact={true} />
                                            <KpiCard label="Water Consumption (KL)" value={fmtNum(displayWaterConsumption)} icon="water_drop" tone="teal" compact={true} />
                                            <KpiCard label="Diesel Consumption" value={`${fmtNum(kpiTotals.diesel)} L`} icon="local_gas_station" tone="red" compact={true} />
                                            <KpiCard label="Energy Cost" value={fmtINR(aggregatedCosts.energyCost, { compact: true })} sub={`Tariff ₹ ${Number(aggregatedCosts.electRate || 0).toFixed(2)}/unit`} icon="currency_rupee" tone="purple" compact={true} />
                                            <KpiCard label="Water Cost" value={fmtINR(displayWaterCost, { compact: true })} icon="payments" tone="teal" compact={true} />
                                            <ProductionKpiCard production={kpiTotals.production} odu={kpiTotals.odu} idu={kpiTotals.idu} compact={true} />
                                        </div>
                                        
                                        {/* Row 2: 5 Secondary Collapsible Cards */}
                                        {canMonitor && (
                                            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2 transition-all duration-300">
                                                <KpiCard label="Compressed Air" value={`${fmtNum(kpiTotals.air)} units`} icon="air" tone="indigo" compact={true} />
                                                <KpiCard label="LPG/PNG Used" value={`${fmtNum(kpiTotals.lpg)} kg`} icon="propane_tank" tone="pink" compact={true} />
                                                <KpiCard label="Waste Haz" value={`${fmtNum(kpiTotals.wasteHaz)} kg`} icon="delete_forever" tone="red" compact={true} />
                                                <KpiCard label="Waste Non-Haz" value={`${fmtNum(kpiTotals.wasteNHaz)} kg`} icon="delete" tone="gray" compact={true} />
                                                <KpiCard label="Waste Recycled" value={`${fmtNum(kpiTotals.wasteRec)} kg`} sub={`${kpiTotals.wasteNHaz > 0 ? (kpiTotals.wasteRec / (kpiTotals.wasteHaz + kpiTotals.wasteNHaz) * 100).toFixed(0) : 0}% recovery`} icon="recycling" tone="emerald" compact={true} />
                                            </div>
                                        )}
                                    </div>
                                ) : (
                                    <section
                                        className="flex flex-row overflow-x-auto gap-2 scrollbar-hide items-stretch no-print"
                                        style={{ position: 'sticky', top: '64px', zIndex: 15, background: 'var(--bg)', paddingTop: '0px', paddingBottom: '4px' }}
                                    >
                                        <KpiCard label="Electricity (kWh)" value={fmtNum(kpiTotals.electricity)} icon="electric_bolt" tone="blue" />
                                        <KpiCard label="Solar Gen (kWh)" value={fmtNum(kpiTotals.solarGenerated)} icon="wb_sunny" tone="amber" />
                                        <KpiCard label="Total Consumption" value={fmtNum(aggregatedCosts.totalConsumption)} sub={`${dashboardGridLabel} + Solar (filter)`} icon="bolt" tone="orange" />
                                        <KpiCard label="Water Consumption (KL)" value={fmtNum(displayWaterConsumption)} icon="water_drop" tone="teal" />
                                        <KpiCard label="Diesel Consumption" value={`${fmtNum(kpiTotals.diesel)} L`} icon="local_gas_station" tone="red" />
                                        <KpiCard label="Energy Cost" value={fmtINR(aggregatedCosts.energyCost, { compact: true })} sub={`Tariff ₹ ${Number(aggregatedCosts.electRate || 0).toFixed(2)}/unit`} icon="currency_rupee" tone="purple" />
                                        <KpiCard label="Water Cost" value={fmtINR(displayWaterCost, { compact: true })} icon="payments" tone="teal" />
                                        <ProductionKpiCard production={kpiTotals.production} odu={kpiTotals.odu} idu={kpiTotals.idu} />
                                        
                                        {canMonitor && (
                                            <React.Fragment>
                                                <KpiCard label="Compressed Air" value={`${fmtNum(kpiTotals.air)} units`} icon="air" tone="indigo" />
                                                <KpiCard label="LPG/PNG Used" value={`${fmtNum(kpiTotals.lpg)} kg`} icon="propane_tank" tone="pink" />
                                                <KpiCard label="Waste Haz" value={`${fmtNum(kpiTotals.wasteHaz)} kg`} icon="delete_forever" tone="red" />
                                                <KpiCard label="Waste Non-Haz" value={`${fmtNum(kpiTotals.wasteNHaz)} kg`} icon="delete" tone="gray" />
                                                <KpiCard label="Waste Recycled" value={`${fmtNum(kpiTotals.wasteRec)} kg`} sub={`${kpiTotals.wasteNHaz > 0 ? (kpiTotals.wasteRec / (kpiTotals.wasteHaz + kpiTotals.wasteNHaz) * 100).toFixed(0) : 0}% recovery`} icon="recycling" tone="emerald" />
                                            </React.Fragment>
                                        )}
                                    </section>
                                )}

                                {/* Charts — Row 1 Daily / Row 2 Monthly / Row 3 Yearly + Remaining */}
                                <div className="space-y-2.5 no-print">

                                     {/* ROW 1: Daily trends — last 7 calendar days */}
                                     <section className="grid grid-cols-1 md:grid-cols-3 gap-3">
                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-3 shadow-sm">
                                             <div className="mb-0.5">
                                                 <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Daily Electricity (kWh)</h4>
                                                 <p className="text-[9px] text-slate-400">Last 7 days grid consumption</p>
                                             </div>
                                             <ResponsiveContainer width="100%" height={215}>
                                                 <AreaChart data={last7DaysTrendsData} margin={{ top: 16, right: 28, left: 4, bottom: 4 }}>
                                                     <defs>
                                                         <linearGradient id="colorElect7" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="0%" stopColor={CHART.dailyElect} stopOpacity={0.4}/>
                                                             <stop offset="100%" stopColor={CHART.dailyElect} stopOpacity={0.05}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} interval={0} padding={{ left: 12, right: 12 }} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={48} domain={[0, 'auto']} />
                                                     <Tooltip contentStyle={{ fontSize: 9, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} />
                                                     <Area type="monotone" dataKey="electricity" name="Electricity" stroke={CHART.dailyElect} fill="url(#colorElect7)" strokeWidth={2.5} dot={{ r: 4.5, fill: CHART.dailyElect, stroke: '#ffffff', strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} animationEasing="ease-in-out">
                                                         <LabelList dataKey="electricity" position="top" offset={8} style={{ fontSize: 9, fontWeight: 700, fill: CHART.dailyElect }} formatter={(v) => fmtNum(v)} />
                                                     </Area>
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>

                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-3 shadow-sm">
                                             <div className="mb-0.5">
                                                 <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Daily Solar (kWh)</h4>
                                                 <p className="text-[9px] text-slate-400">Last 7 days solar generation</p>
                                             </div>
                                             <ResponsiveContainer width="100%" height={215}>
                                                 <AreaChart data={last7DaysTrendsData} margin={{ top: 16, right: 28, left: 4, bottom: 4 }}>
                                                     <defs>
                                                         <linearGradient id="colorSolar7" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="0%" stopColor={CHART.dailySolar} stopOpacity={0.4}/>
                                                             <stop offset="100%" stopColor={CHART.dailySolar} stopOpacity={0.05}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} interval={0} padding={{ left: 12, right: 12 }} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={48} domain={[0, 'auto']} />
                                                     <Tooltip contentStyle={{ fontSize: 9, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} />
                                                     <Area type="monotone" dataKey="solarGen" name="Solar Gen" stroke={CHART.dailySolar} fill="url(#colorSolar7)" strokeWidth={2.5} dot={{ r: 4.5, fill: CHART.dailySolar, stroke: '#ffffff', strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} animationEasing="ease-in-out">
                                                         <LabelList dataKey="solarGen" position="top" offset={8} style={{ fontSize: 9, fontWeight: 700, fill: CHART.dailySolar }} formatter={(v) => fmtNum(v)} />
                                                     </Area>
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>

                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-3 shadow-sm">
                                             <div className="mb-0.5">
                                                 <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Daily Diesel (L)</h4>
                                                 <p className="text-[9px] text-slate-400">Last 7 days DG fuel use</p>
                                             </div>
                                             <ResponsiveContainer width="100%" height={215}>
                                                 <AreaChart data={last7DaysTrendsData} margin={{ top: 16, right: 28, left: 4, bottom: 4 }}>
                                                     <defs>
                                                         <linearGradient id="colorDiesel7" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="0%" stopColor={CHART.dailyDiesel} stopOpacity={0.4}/>
                                                             <stop offset="100%" stopColor={CHART.dailyDiesel} stopOpacity={0.05}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} interval={0} padding={{ left: 12, right: 12 }} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={48} domain={[0, 'auto']} />
                                                     <Tooltip contentStyle={{ fontSize: 9, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} />
                                                     <Area type="monotone" dataKey="diesel" name="Diesel Liters" stroke={CHART.dailyDiesel} fill="url(#colorDiesel7)" strokeWidth={2.5} dot={{ r: 4.5, fill: CHART.dailyDiesel, stroke: '#ffffff', strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} animationEasing="ease-in-out">
                                                         <LabelList dataKey="diesel" position="top" offset={8} style={{ fontSize: 9, fontWeight: 700, fill: CHART.dailyDiesel }} formatter={(v) => fmtNum(v)} />
                                                     </Area>
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>
                                     </section>

                                     {/* ROW 2: Monthly — smooth area spline (sample style), unique color each */}
                                     <section className="grid grid-cols-1 md:grid-cols-3 gap-3">
                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-3 shadow-sm">
                                             <div className="mb-0.5">
                                                 <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Monthly Electricity Load</h4>
                                                 <p className="text-[9px] text-slate-400">Last 5 months grid consumption</p>
                                             </div>
                                             <ResponsiveContainer width="100%" height={215}>
                                                 <AreaChart data={last5MonthsTrendsData} margin={{ top: 16, right: 28, left: 4, bottom: 4 }}>
                                                     <defs>
                                                         <linearGradient id="colorElectM" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="0%" stopColor={CHART.monthlyElect} stopOpacity={0.45}/>
                                                             <stop offset="100%" stopColor={CHART.monthlyElect} stopOpacity={0.06}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} padding={{ left: 12, right: 12 }} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={48} domain={[0, 'auto']} />
                                                     <Tooltip contentStyle={{ fontSize: 9, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} />
                                                     <Area type="monotone" dataKey="electricity" name="Electricity" stroke={CHART.monthlyElect} fill="url(#colorElectM)" strokeWidth={3} dot={{ r: 5, fill: CHART.monthlyElect, stroke: '#ffffff', strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} animationEasing="ease-in-out">
                                                         <LabelList dataKey="electricity" position="top" offset={8} style={{ fontSize: 10, fontWeight: 700, fill: CHART.monthlyElect }} formatter={(v) => fmtNum(v)} />
                                                     </Area>
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>

                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-3 shadow-sm">
                                             <div className="mb-0.5">
                                                 <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Monthly Solar Gen</h4>
                                                 <p className="text-[9px] text-slate-400">Last 5 months solar generation</p>
                                             </div>
                                             <ResponsiveContainer width="100%" height={215}>
                                                 <AreaChart data={last5MonthsTrendsData} margin={{ top: 16, right: 28, left: 4, bottom: 4 }}>
                                                     <defs>
                                                         <linearGradient id="colorSolarM" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="0%" stopColor={CHART.monthlySolar} stopOpacity={0.45}/>
                                                             <stop offset="100%" stopColor={CHART.monthlySolar} stopOpacity={0.06}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} padding={{ left: 12, right: 12 }} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={48} domain={[0, 'auto']} />
                                                     <Tooltip contentStyle={{ fontSize: 9, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} />
                                                     <Area type="monotone" dataKey="solarGen" name="Solar Gen" stroke={CHART.monthlySolar} fill="url(#colorSolarM)" strokeWidth={3} dot={{ r: 5, fill: CHART.monthlySolar, stroke: '#ffffff', strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} animationEasing="ease-in-out">
                                                         <LabelList dataKey="solarGen" position="top" offset={8} style={{ fontSize: 10, fontWeight: 700, fill: CHART.monthlySolar }} formatter={(v) => fmtNum(v)} />
                                                     </Area>
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>

                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-3 shadow-sm">
                                             <div className="mb-0.5">
                                                 <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Monthly Diesel</h4>
                                                 <p className="text-[9px] text-slate-400">Last 5 months DG fuel use</p>
                                             </div>
                                             <ResponsiveContainer width="100%" height={215}>
                                                 <AreaChart data={last5MonthsTrendsData} margin={{ top: 16, right: 28, left: 4, bottom: 4 }}>
                                                     <defs>
                                                         <linearGradient id="colorDieselM" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="0%" stopColor={CHART.monthlyDiesel} stopOpacity={0.45}/>
                                                             <stop offset="100%" stopColor={CHART.monthlyDiesel} stopOpacity={0.06}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} padding={{ left: 12, right: 12 }} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={48} domain={[0, 'auto']} />
                                                     <Tooltip contentStyle={{ fontSize: 9, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} />
                                                     <Area type="monotone" dataKey="diesel" name="Diesel Liters" stroke={CHART.monthlyDiesel} fill="url(#colorDieselM)" strokeWidth={3} dot={{ r: 5, fill: CHART.monthlyDiesel, stroke: '#ffffff', strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} animationEasing="ease-in-out">
                                                         <LabelList dataKey="diesel" position="top" offset={8} style={{ fontSize: 10, fontWeight: 700, fill: CHART.monthlyDiesel }} formatter={(v) => fmtNum(v)} />
                                                     </Area>
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>
                                     </section>

                                     {/* ROW 3: Yearly area charts + Target pie */}
                                     <section className="grid grid-cols-1 md:grid-cols-3 gap-4">
                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-4 shadow-sm">
                                             <div className="mb-1">
                                                 <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Yearly Electricity</h4>
                                                 <p className="text-[9px] text-slate-400">Last 4 years grid load</p>
                                             </div>
                                             <ResponsiveContainer width="100%" height={250}>
                                                 <AreaChart data={last4YearsTrendsData} margin={{ top: 16, right: 28, left: 4, bottom: 4 }}>
                                                     <defs>
                                                         <linearGradient id="colorElectY" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="0%" stopColor={CHART.yearlyElect} stopOpacity={0.4}/>
                                                             <stop offset="100%" stopColor={CHART.yearlyElect} stopOpacity={0.05}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} padding={{ left: 12, right: 12 }} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={48} domain={[0, 'auto']} />
                                                     <Tooltip contentStyle={{ fontSize: 9, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} />
                                                     <Area type="monotone" dataKey="electricity" name="Electricity" stroke={CHART.yearlyElect} fill="url(#colorElectY)" strokeWidth={2.5} dot={{ r: 4.5, fill: '#ffffff', stroke: CHART.yearlyElect, strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} animationEasing="ease-in-out">
                                                         <LabelList dataKey="electricity" position="top" offset={8} style={{ fontSize: 9, fontWeight: 700, fill: CHART.yearlyElect }} formatter={(v) => fmtNum(v)} />
                                                     </Area>
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>

                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-4 shadow-sm">
                                             <div className="mb-1">
                                                 <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Yearly Cost</h4>
                                                 <p className="text-[9px] text-slate-400">Last 4 years utility cost</p>
                                             </div>
                                             <ResponsiveContainer width="100%" height={250}>
                                                 <AreaChart data={last4YearsTrendsData} margin={{ top: 16, right: 28, left: 4, bottom: 4 }}>
                                                     <defs>
                                                         <linearGradient id="colorCostY" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="0%" stopColor={CHART.yearlyCost} stopOpacity={0.4}/>
                                                             <stop offset="100%" stopColor={CHART.yearlyCost} stopOpacity={0.05}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} padding={{ left: 12, right: 12 }} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={48} domain={[0, 'auto']} />
                                                     <Tooltip contentStyle={{ fontSize: 9, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} formatter={(v) => [fmtINR(v), "Cost"]} />
                                                     <Area type="monotone" dataKey="cost" name="Cost (₹)" stroke={CHART.yearlyCost} fill="url(#colorCostY)" strokeWidth={2.5} dot={{ r: 4.5, fill: '#ffffff', stroke: CHART.yearlyCost, strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} animationEasing="ease-in-out">
                                                         <LabelList dataKey="cost" position="top" offset={8} style={{ fontSize: 9, fontWeight: 700, fill: CHART.yearlyCost }} formatter={(v) => fmtINR(v, { compact: true })} />
                                                     </Area>
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>

                                        <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-4 shadow-sm overflow-visible">
                                            <div className="mb-2">
                                                <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Target vs Actual Load</h4>
                                                <p className="text-[9px] text-slate-400">Electricity actual vs monthly target</p>
                                            </div>
                                            <ResponsiveContainer width="100%" height={250}>
                                                <PieChart margin={{ top: 8, right: 8, bottom: 8, left: 8 }}>
                                                    <Pie
                                                        data={targetVsActualData}
                                                        dataKey="value"
                                                        nameKey="name"
                                                        cx="50%"
                                                        cy="46%"
                                                        outerRadius={72}
                                                        innerRadius={40}
                                                        paddingAngle={3}
                                                        isAnimationActive={true}
                                                        animationDuration={1000}
                                                    >
                                                        {targetVsActualData.map((d) => (
                                                            <Cell key={d.name} fill={TARGET_PIE_COLORS[d.name] || "#94a3b8"} stroke="#fff" strokeWidth={2} />
                                                        ))}
                                                    </Pie>
                                                    <Tooltip contentStyle={{ fontSize: 10, borderRadius: 8, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} formatter={(v) => [fmtNum(v) + " kWh", ""]} />
                                                    <Legend
                                                        verticalAlign="bottom"
                                                        height={48}
                                                        wrapperStyle={{ fontSize: 10 }}
                                                        formatter={(value) => {
                                                            const row = targetVsActualData.find(d => d.name === value);
                                                            return row ? `${value}: ${fmtNum(row.value)} kWh` : value;
                                                        }}
                                                    />
                                                </PieChart>
                                            </ResponsiveContainer>
                                        </div>
                                    </section>

                                     {/* ROW 4: Solar / Water / Production — rich enhanced design */}
                                     <section className="grid grid-cols-1 md:grid-cols-3 gap-4">
                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-4 shadow-sm">
                                             <div className="mb-2 flex items-center justify-between">
                                                 <div>
                                                     <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Solar Gen vs Total Consumption</h4>
                                                     <p className="text-[9px] text-slate-400">Solar generation vs {dashboardGridLabel} + Solar total (Last 30 days)</p>
                                                 </div>
                                                 <span className="text-[9px] font-bold px-2 py-0.5 rounded-full bg-amber-50 dark:bg-amber-950/40 text-amber-600 dark:text-amber-400 border border-amber-200/50">Daily Trend</span>
                                             </div>
                                             <ResponsiveContainer width="100%" height={210}>
                                                 <AreaChart data={last30DaysTrendsData}>
                                                     <defs>
                                                         <linearGradient id="colorTotalR4" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="5%" stopColor="#0284c7" stopOpacity={0.35}/>
                                                             <stop offset="95%" stopColor="#0284c7" stopOpacity={0.02}/>
                                                         </linearGradient>
                                                         <linearGradient id="colorSolarR4" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="5%" stopColor="#f59e0b" stopOpacity={0.4}/>
                                                             <stop offset="95%" stopColor="#f59e0b" stopOpacity={0.02}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} minTickGap={25} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={42} tickFormatter={(v) => fmtNum(v, { compact: true })} />
                                                     <Tooltip
                                                         content={({ active, payload, label }) => {
                                                             if (!active || !payload || !payload.length) return null;
                                                             const d = payload[0].payload;
                                                             const solar = Number(d.solarGen || 0);
                                                             const total = Number(d.totalConsumption || 0);
                                                             const grid = Math.max(0, total - solar);
                                                             const solarPct = total > 0 ? ((solar / total) * 100).toFixed(1) : "0.0";
                                                             return (
                                                                 <div className="bg-slate-900/95 text-white p-2.5 rounded-xl shadow-xl border border-slate-700/60 text-[10px] space-y-1 z-50 backdrop-blur-md min-w-[170px]">
                                                                     <p className="font-extrabold border-b border-slate-700/60 pb-1 text-slate-300 uppercase tracking-wider text-[9.5px]">{label} ({d.date})</p>
                                                                     <div className="flex items-center justify-between text-amber-400 font-semibold pt-0.5">
                                                                         <span>☀️ Solar Gen:</span>
                                                                         <span className="font-extrabold">{fmtNum(solar)} kWh</span>
                                                                     </div>
                                                                     <div className="flex items-center justify-between text-sky-400 font-semibold">
                                                                         <span>⚡ Total Load:</span>
                                                                         <span className="font-extrabold">{fmtNum(total)} kWh</span>
                                                                     </div>
                                                                     <div className="flex items-center justify-between text-slate-400 font-medium">
                                                                         <span>🔌 Grid Share:</span>
                                                                         <span>{fmtNum(grid)} kWh</span>
                                                                     </div>
                                                                     <div className="flex items-center justify-between text-emerald-400 font-bold border-t border-slate-800 pt-1">
                                                                         <span>🌱 Solar Share:</span>
                                                                         <span>{solarPct}%</span>
                                                                     </div>
                                                                 </div>
                                                             );
                                                         }}
                                                     />
                                                     <Legend wrapperStyle={{ fontSize: 9, paddingTop: 4 }} />
                                                     <Area type="monotone" dataKey="totalConsumption" name="Total Consumption" stroke="#0284c7" fill="url(#colorTotalR4)" strokeWidth={2.5} dot={false} activeDot={{ r: 5 }} isAnimationActive={true} animationDuration={1000} />
                                                     <Area type="monotone" dataKey="solarGen" name="Solar Gen" stroke="#f59e0b" fill="url(#colorSolarR4)" strokeWidth={2.5} dot={false} activeDot={{ r: 5 }} isAnimationActive={true} animationDuration={1000} />
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>

                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-4 shadow-sm">
                                             <div className="mb-2 flex items-center justify-between">
                                                 <div>
                                                     <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Water Consumption (KL)</h4>
                                                     <p className="text-[9px] text-slate-400">Daily water resource logging (Last 30 days)</p>
                                                 </div>
                                                 <span className="text-[9px] font-bold px-2 py-0.5 rounded-full bg-teal-50 dark:bg-teal-950/40 text-teal-600 dark:text-teal-400 border border-teal-200/50">KL Metric</span>
                                             </div>
                                             <ResponsiveContainer width="100%" height={210}>
                                                 <AreaChart data={waterTrendsData}>
                                                     <defs>
                                                         <linearGradient id="colorWaterR4" x1="0" y1="0" x2="0" y2="1">
                                                             <stop offset="5%" stopColor="#0d9488" stopOpacity={0.4}/>
                                                             <stop offset="95%" stopColor="#0d9488" stopOpacity={0.02}/>
                                                         </linearGradient>
                                                     </defs>
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} minTickGap={25} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={36} tickFormatter={(v) => fmtNum(v, { compact: true })} />
                                                     <Tooltip
                                                         contentStyle={{ fontSize: 10, borderRadius: 10, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }}
                                                         labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }}
                                                         formatter={(v) => [`${fmtNum(v)} KL`, "Water Consumed"]}
                                                     />
                                                     <Area type="monotone" dataKey="water" name="Water (KL)" stroke="#0d9488" fill="url(#colorWaterR4)" strokeWidth={2.5} dot={false} activeDot={{ r: 6, fill: '#0d9488', stroke: '#ffffff', strokeWidth: 2 }} isAnimationActive={true} animationDuration={1000} />
                                                 </AreaChart>
                                             </ResponsiveContainer>
                                         </div>

                                         <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-4 shadow-sm">
                                             <div className="mb-2 flex items-center justify-between">
                                                 <div>
                                                     <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Production Output</h4>
                                                     <p className="text-[9px] text-slate-400">Manufactured outputs breakdown (Last 30 days)</p>
                                                 </div>
                                                 <span className="text-[9px] font-bold px-2 py-0.5 rounded-full bg-indigo-50 dark:bg-indigo-950/40 text-indigo-600 dark:text-indigo-400 border border-indigo-200/50">Units</span>
                                             </div>
                                             <ResponsiveContainer width="100%" height={210}>
                                                 <BarChart data={last30DaysTrendsData} barGap={2} barCategoryGap="18%">
                                                     <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                     <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} minTickGap={25} />
                                                     <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={42} tickFormatter={(v) => fmtNum(v, { compact: true })} />
                                                     <Tooltip contentStyle={{ fontSize: 10, borderRadius: 10, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} formatter={(v, name) => [`${fmtNum(v)} units`, name]} />
                                                     <Legend wrapperStyle={{ fontSize: 9, paddingTop: 4 }} />
                                                     <Bar dataKey="production" name="Sets" fill="#4f46e5" radius={[3, 3, 0, 0]} isAnimationActive={true} animationDuration={1000} />
                                                     <Bar dataKey="odu" name="ODU" fill="#0891b2" radius={[3, 3, 0, 0]} isAnimationActive={true} animationDuration={1000} />
                                                     <Bar dataKey="idu" name="IDU" fill="#c026d3" radius={[3, 3, 0, 0]} isAnimationActive={true} animationDuration={1000} />
                                                 </BarChart>
                                             </ResponsiveContainer>
                                         </div>
                                     </section>

                                     {/* Remaining graphs */}
                                     <section className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                                         {canMonitor && (
                                             <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-4 shadow-sm">
                                                 <div className="mb-2">
                                                     <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">LPG/PNG Fuel Used (kg)</h4>
                                                     <p className="text-[9px] text-slate-400">Gas fuels utility tracking (Last 30 days)</p>
                                                 </div>
                                                 <ResponsiveContainer width="100%" height={200}>
                                                     <AreaChart data={last30DaysTrendsData}>
                                                         <defs>
                                                             <linearGradient id="colorLpgR5" x1="0" y1="0" x2="0" y2="1">
                                                                 <stop offset="5%" stopColor={COLORS.purple} stopOpacity={0.4}/>
                                                                 <stop offset="95%" stopColor={COLORS.purple} stopOpacity={0.02}/>
                                                             </linearGradient>
                                                         </defs>
                                                         <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                         <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} minTickGap={25} />
                                                         <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={40} tickFormatter={(v) => fmtNum(v, { compact: true })} />
                                                         <Tooltip contentStyle={{ fontSize: 10, borderRadius: 10, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} formatter={(v) => [`${fmtNum(v)} kg`, "LPG/PNG"]} />
                                                         <Area type="monotone" dataKey="lpg" name="LPG/PNG kg" stroke={COLORS.purple} fill="url(#colorLpgR5)" strokeWidth={2.5} dot={false} activeDot={{ r: 6 }} isAnimationActive={true} animationDuration={1000} />
                                                     </AreaChart>
                                                 </ResponsiveContainer>
                                             </div>
                                         )}

                                         {canMonitor && (
                                             <div className="bg-white dark:bg-[#121a29] rounded-2xl border border-slate-200/70 dark:border-[#26334a] p-4 shadow-sm">
                                                 <div className="mb-2">
                                                     <h4 className="text-[11px] font-extrabold text-slate-800 dark:text-slate-200 uppercase tracking-wider">Waste Generation</h4>
                                                     <p className="text-[9px] text-slate-400">Total hazardous & non-hazardous (Last 30 days)</p>
                                                 </div>
                                                 <ResponsiveContainer width="100%" height={200}>
                                                     <BarChart data={last30DaysTrendsData}>
                                                         <CartesianGrid stroke="#e2e8f0" strokeDasharray="0" vertical={false} />
                                                         <XAxis dataKey="label" tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} minTickGap={25} />
                                                         <YAxis tick={{ fontSize: 9, fill: '#64748b' }} axisLine={false} tickLine={false} width={40} tickFormatter={(v) => fmtNum(v, { compact: true })} />
                                                         <Tooltip contentStyle={{ fontSize: 10, borderRadius: 10, background: 'var(--tooltip-bg)', border: '1px solid var(--border)', color: 'var(--text-body)' }} labelStyle={{ fontWeight: 'bold', color: 'var(--text-heading)' }} formatter={(v) => [`${fmtNum(v)} kg`, "Waste"]} />
                                                         <Bar dataKey="waste" name="Waste (kg)" fill={COLORS.red} radius={[4, 4, 0, 0]} isAnimationActive={true} animationDuration={1000} />
                                                     </BarChart>
                                                 </ResponsiveContainer>
                                             </div>
                                         )}
                                     </section>
                                </div>
                            </div>
                        )}

                        {/* 2. DAILY DATA ENTRY GRID COMPONENT */}
                        {activeTab === "entry" && (
                            <div className="space-y-6 pt-4">
                                <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-3 border-b pb-4">
                                    <div className="shrink-0">
                                        <h2 className="text-base font-extrabold text-slate-900 flex items-center gap-1.5 uppercase whitespace-nowrap">
                                            <span className="material-symbols-outlined text-[#0284c7]">edit_document</span>
                                            <span>Daily Operations Logs</span>
                                        </h2>
                                        <p className="text-xs text-slate-400 mt-0.5 whitespace-nowrap">Read, insert, update and manage operational transaction metrics</p>
                                    </div>

                                    {/* All controls in ONE single horizontal line */}
                                    <div className="flex items-center gap-1.5 flex-nowrap overflow-x-auto pb-0.5 max-w-full">
                                        <select
                                            value={entryLocationFilter}
                                            onChange={(e) => {
                                                setEntryLocationFilter(e.target.value);
                                                setEntryPlantFilter("all");
                                            }}
                                            className="h-9 rounded-xl border border-slate-200 px-2.5 text-xs bg-white text-slate-700 focus:outline-none focus:ring-1 focus:ring-sky-500 font-semibold shrink-0"
                                            title="Filter by location"
                                        >
                                            <option value="all">All Locations</option>
                                            {reportLocations.map((loc) => (
                                                <option key={loc} value={loc}>{loc}</option>
                                            ))}
                                        </select>

                                        <select
                                            value={entryPlantFilter}
                                            onChange={(e) => setEntryPlantFilter(e.target.value)}
                                            className="h-9 rounded-xl border border-slate-200 px-2.5 text-xs bg-white text-slate-700 focus:outline-none focus:ring-1 focus:ring-sky-500 font-semibold shrink-0"
                                            title="Filter by plant"
                                        >
                                            <option value="all">All Plants</option>
                                            {entryFilterPlants.map((p) => (
                                                <option key={p.plant_code} value={p.plant_code}>
                                                    {p.plant_display_name || p.plant_name || p.plant_code}
                                                </option>
                                            ))}
                                        </select>

                                        <div className="relative shrink-0">
                                            <input
                                                type="text"
                                                placeholder="Search date, plant, operator..."
                                                value={entrySearch}
                                                onChange={(e) => setEntrySearch(e.target.value)}
                                                className="w-36 sm:w-44 h-9 rounded-xl border border-slate-200 pl-8 pr-2.5 text-xs bg-slate-50 focus:outline-none focus:ring-1 focus:ring-sky-500 font-medium"
                                            />
                                            <span className="material-symbols-outlined absolute left-2.5 top-2.5 text-[15px] text-slate-400">search</span>
                                        </div>

                                        {currentUser.role === "IT_ADMIN" && (
                                            !isDailyDeleteMode ? (
                                                <button
                                                    onClick={() => setIsDailyDeleteMode(true)}
                                                    className="flex items-center gap-1 h-9 px-2.5 bg-white border border-red-200 hover:bg-red-50 text-red-600 rounded-xl text-xs font-bold transition shadow-sm cursor-pointer shrink-0 whitespace-nowrap"
                                                    title="Click to enable selection mode for deleting entries"
                                                >
                                                    <span className="material-symbols-outlined text-[16px]">delete_sweep</span>
                                                    <span>Delete All</span>
                                                </button>
                                            ) : (
                                                <div className="flex items-center gap-1 shrink-0">
                                                    <button
                                                        onClick={handleBulkDelete}
                                                        className="flex items-center gap-1 h-9 px-2.5 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold transition shadow-sm border-none cursor-pointer whitespace-nowrap"
                                                        title={selectedRowIds.size > 0 ? `Delete ${selectedRowIds.size} selected entries` : "Delete all database entries"}
                                                    >
                                                        <span className="material-symbols-outlined text-[16px]">{selectedRowIds.size > 0 ? 'delete' : 'delete_forever'}</span>
                                                        <span>{selectedRowIds.size > 0 ? `Confirm Delete (${selectedRowIds.size})` : "Delete All Entries"}</span>
                                                    </button>
                                                    <button
                                                        onClick={() => { setIsDailyDeleteMode(false); setSelectedRowIds(new Set()); }}
                                                        className="flex items-center gap-1 h-9 px-2 bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 rounded-xl text-xs font-bold transition shadow-sm cursor-pointer whitespace-nowrap"
                                                        title="Cancel selection mode"
                                                    >
                                                        <span>Cancel</span>
                                                    </button>
                                                </div>
                                            )
                                        )}

                                        {currentUser.role === "IT_ADMIN" && (
                                            <button
                                                onClick={openMassUploadModal}
                                                className="flex items-center gap-1.5 h-9 px-3 bg-sky-50 hover:bg-sky-100 text-sky-700 border border-sky-200 rounded-xl text-xs font-bold transition shadow-sm cursor-pointer shrink-0 whitespace-nowrap"
                                                title="Smart Mass Excel Importer / Bulk Data Uploader (IT Admin Only)"
                                            >
                                                <span className="material-symbols-outlined text-[17px] text-sky-600">upload_file</span>
                                                <span>Mass Upload</span>
                                            </button>
                                        )}

                                        <button
                                            onClick={() => openDailyForm()}
                                            className="flex items-center gap-1 h-9 px-3.5 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl text-xs font-bold transition shadow-sm border-none cursor-pointer shrink-0 whitespace-nowrap"
                                        >
                                            <span className="material-symbols-outlined text-[16px]">add</span>
                                            <span>Daily Entry</span>
                                        </button>
                                    </div>
                                </div>

                                {/* Data Table Grid */}
                                <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
                                    <div className="overflow-x-auto scrollbar-hide">
                                        <table className="w-full border-collapse text-left text-xs">
                                            <thead>
                                                <tr className="border-b border-slate-100 text-slate-400 font-bold uppercase tracking-wider text-[9px] bg-slate-50/50 whitespace-nowrap">
                                                    {currentUser.role === "IT_ADMIN" && isDailyDeleteMode && (
                                                        <th className="py-3 px-2.5 w-8">
                                                            <input
                                                                type="checkbox"
                                                                checked={paginatedEntries.length > 0 && paginatedEntries.every(e => selectedRowIds.has(e.id))}
                                                                onChange={handleToggleSelectAllDaily}
                                                                className="rounded text-sky-600 focus:ring-sky-500 h-4 w-4 cursor-pointer"
                                                                title="Select all visible entries"
                                                            />
                                                        </th>
                                                    )}
                                                    <th className="py-3 px-2.5">Date</th>
                                                    <th className="py-3 px-2.5">Plant</th>
                                                    <th className="py-3 px-2.5 text-right">Daily Reading (kWh)</th>
                                                    <th className="py-3 px-2.5 text-right">Units Diff</th>
                                                    <th className="py-3 px-2.5 text-right">{entryTableContext.gridLabel} Units{entryTableContext.mf ? ` (x${entryTableContext.mf})` : ""}</th>
                                                    <th className="py-3 px-2.5 text-right">Solar Units</th>
                                                    <th className="py-3 px-2.5 text-right">Total Units</th>
                                                    <th className="py-3 px-2.5 text-right">{entryTableContext.gridLabel} ₹/Unit</th>
                                                    <th className="py-3 px-2.5 text-right">₹ {entryTableContext.gridLabel}</th>
                                                    <th className="py-3 px-2.5 text-right">₹ Solar</th>
                                                    <th className="py-3 px-2.5 text-right">₹/L Diesel</th>
                                                    <th className="py-3 px-2.5 text-right">Diesel (L)</th>
                                                    <th className="py-3 px-2.5 text-right">₹ Diesel</th>
                                                    <th className="py-3 px-2.5 text-right">₹ Total</th>
                                                    <th className="py-3 px-2.5 text-right">ODU</th>
                                                    <th className="py-3 px-2.5 text-right">IDU</th>
                                                    <th className="py-3 px-2.5 text-right">Prod (Set)</th>
                                                    <th className="py-3 px-2.5 text-right">Prod/Set ₹</th>
                                                    <th className="py-3 px-2.5">Remarks</th>
                                                    <th className="py-3 px-2.5 text-right">Actions</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-100 text-slate-700">
                                                {paginatedEntries.length === 0 ? (
                                                    <tr>
                                                        <td colSpan="20" className="py-12 text-center text-slate-400 font-semibold">No operational entries logged in system</td>
                                                    </tr>
                                                ) : (
                                                    paginatedEntries.map(e => {
                                                        const rowMeta = resolvePlantMeta(e.plant);
                                                        const rowLoc = e.location || rowMeta.location;
                                                        const rowMf = resolveMultiplyFactor(multiplyFactors, rowMeta.code || e.plant, rowLoc, e.date);
                                                        const msebUnits = Number(e.electricity_consumption) || 0;
                                                        const solarUnits = Number(e.solar_generated) || 0;
                                                        const dieselL = Number(e.diesel_used) || 0;
                                                        const unitsDiff = rowMf > 0 ? msebUnits / rowMf : 0;
                                                        const totalUnits = msebUnits + solarUnits;
                                                        const msebRate = msebUnits > 0 ? (Number(e.electricity_cost) || 0) / msebUnits : null;
                                                        const dieselRate = dieselL > 0 ? (Number(e.diesel_cost) || 0) / dieselL : null;
                                                        const prodQty = Number(e.production_qty) || 0;
                                                        const prodSetCost = prodQty > 0 ? (Number(e.total_cost) || 0) / prodQty : null;
                                                        return (
                                                        <tr
                                                            key={e.id}
                                                            onClick={() => setViewingRecord(e)}
                                                            title="Click to view full entry details"
                                                            className={`hover:bg-slate-50/60 transition cursor-pointer whitespace-nowrap ${selectedRowIds.has(e.id) ? 'bg-sky-50/40 dark:bg-sky-950/20' : ''}`}
                                                        >
                                                            {currentUser.role === "IT_ADMIN" && isDailyDeleteMode && (
                                                                <td className="py-3 px-2.5" onClick={(ev) => ev.stopPropagation()}>
                                                                    <input
                                                                        type="checkbox"
                                                                        checked={selectedRowIds.has(e.id)}
                                                                        onChange={() => handleToggleRowSelect(e.id)}
                                                                        className="rounded text-sky-600 focus:ring-sky-500 h-4 w-4 cursor-pointer"
                                                                    />
                                                                </td>
                                                            )}
                                                            <td className="py-3 px-2.5 font-bold text-slate-900">{e.date}</td>
                                                            <td className="py-3 px-2.5">
                                                                <span className={`px-2 py-0.5 rounded font-bold text-[9px] uppercase ${e.plant === 'NGM' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-sky-50 text-sky-700 border border-sky-100'}`}>
                                                                    {e.plant}
                                                                </span>
                                                            </td>
                                                            <td className="py-3 px-2.5 text-right text-slate-800">{fmtNum(e.electricity_opening)}</td>
                                                            <td className="py-3 px-2.5 text-right text-slate-800">{fmtNum(unitsDiff, 2)}</td>
                                                            <td className="py-3 px-2.5 text-right font-medium text-slate-800">{fmtNum(msebUnits)}</td>
                                                            <td className="py-3 px-2.5 text-right text-amber-600 font-medium">{fmtNum(solarUnits, 2)}</td>
                                                            <td className="py-3 px-2.5 text-right text-slate-800">{fmtNum(totalUnits, 2)}</td>
                                                            <td className="py-3 px-2.5 text-right text-slate-500">{fmtNum(msebRate, 4)}</td>
                                                            <td className="py-3 px-2.5 text-right font-mono text-slate-800">{fmtMoney(e.electricity_cost)}</td>
                                                            <td className="py-3 px-2.5 text-right font-mono text-amber-700">{fmtMoney(e.solar_cost)}</td>
                                                            <td className="py-3 px-2.5 text-right text-slate-500">{fmtNum(dieselRate, 2)}</td>
                                                            <td className="py-3 px-2.5 text-right text-slate-800">{fmtNum(dieselL, 2)}</td>
                                                            <td className="py-3 px-2.5 text-right font-mono text-slate-800">{fmtMoney(e.diesel_cost)}</td>
                                                            <td className="py-3 px-2.5 text-right font-mono font-bold text-emerald-600">{fmtMoney(e.total_cost)}</td>
                                                            <td className="py-3 px-2.5 text-right text-slate-800">{fmtNum(e.odu)}</td>
                                                            <td className="py-3 px-2.5 text-right text-slate-800">{fmtNum(e.idu)}</td>
                                                            <td className="py-3 px-2.5 text-right font-semibold text-slate-800">{fmtNum(prodQty)}</td>
                                                            <td className="py-3 px-2.5 text-right font-mono text-slate-800">{fmtMoney(prodSetCost)}</td>
                                                            <td className="py-3 px-2.5 text-slate-400 max-w-[130px] truncate" title={e.remarks}>{e.remarks || "—"}</td>
                                                            <td className="py-3 px-2.5 text-right space-x-3 whitespace-nowrap" onClick={(ev) => ev.stopPropagation()}>
                                                                <button onClick={() => openDailyForm(e)} className="text-[#0284c7] hover:text-[#0369a1] font-bold bg-transparent border-none cursor-pointer transition">
                                                                    Edit
                                                                </button>
                                                                <button onClick={() => handleDeleteEntry(e.id)} className="text-red-500 hover:text-red-700 font-bold bg-transparent border-none cursor-pointer transition">
                                                                    Delete
                                                                </button>
                                                            </td>
                                                        </tr>
                                                        );
                                                    })
                                                )}
                                            </tbody>
                                        </table>
                                    </div>

                                    {/* Pagination component */}
                                    <div className="border-t border-slate-100 px-4 py-3 bg-slate-50/50 flex items-center justify-between text-xs font-semibold text-slate-500">
                                        <span>Showing page {entryPage} of {totalPages} ({searchedEntries.length} items)</span>
                                        <div className="flex gap-2">
                                            <button
                                                disabled={entryPage === 1}
                                                onClick={() => setEntryPage(p => p - 1)}
                                                className="px-3 py-1.5 bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-lg disabled:opacity-50 cursor-pointer"
                                            >
                                                Previous
                                            </button>
                                            <button
                                                disabled={entryPage === totalPages}
                                                onClick={() => setEntryPage(p => p + 1)}
                                                className="px-3 py-1.5 bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-lg disabled:opacity-50 cursor-pointer"
                                            >
                                                Next
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* 3. REPORTS COMPONENT */}
                        {activeTab === "reports" && (
                            <div className="space-y-2 -mt-1">
                                <div className="no-print flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200/70 pb-2">
                                    <div className="flex items-center gap-2 flex-wrap">
                                        <h2 className="text-sm font-extrabold text-slate-900 flex items-center gap-1.5 uppercase leading-none">
                                            <span className="material-symbols-outlined text-[#0284c7] text-[18px]">description</span>
                                            <span>Printable Reports Engine</span>
                                        </h2>
                                        <button
                                            onClick={() => setReportsFiltersOpen(o => !o)}
                                            className={`flex items-center gap-1 px-2.5 py-1 rounded-lg border text-[10px] font-extrabold cursor-pointer transition select-none ${reportsFiltersOpen ? 'bg-sky-50 text-sky-700 border-sky-200 hover:bg-sky-100' : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'}`}
                                            title="Toggle Filters"
                                        >
                                            <span className="material-symbols-outlined text-[14px] text-slate-500">tune</span>
                                            <span>Filters</span>
                                        </button>
                                    </div>

                                    <div className="flex items-center gap-1.5">
                                        <button
                                            onClick={handlePrintPlantReport}
                                            className="flex items-center gap-1 px-2.5 py-1.5 bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-lg text-[11px] font-bold transition cursor-pointer"
                                        >
                                            <span className="material-symbols-outlined text-[15px] text-slate-500">print</span>
                                            <span>Print PDF</span>
                                        </button>
                                        <button
                                            onClick={handleExportPlantExcel}
                                            className="flex items-center gap-1 px-2.5 py-1.5 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-lg text-[11px] font-bold transition cursor-pointer border-none"
                                        >
                                            <span className="material-symbols-outlined text-[15px] text-white">download</span>
                                            <span>Export Excel</span>
                                        </button>
                                    </div>
                                </div>

                                {/* Selection Controls */}
                                {reportsFiltersOpen && (
                                    <div className="no-print bg-slate-50 dark:bg-slate-900/40 border border-slate-200/80 dark:border-slate-800 py-1.5 px-2.5 rounded-xl flex flex-wrap items-end gap-2.5">
                                        <div className="flex flex-col">
                                            <label className="text-[8px] font-extrabold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-0.5">Location Filter</label>
                                            <select
                                                value={selectedReportLocation}
                                                onChange={(e) => {
                                                    setSelectedReportLocation(e.target.value);
                                                    setFilters({ ...filters, plant: "all" });
                                                }}
                                                className="h-8 border border-slate-200/90 dark:border-slate-800 rounded-lg px-2.5 text-[11px] bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200 focus:outline-none focus:border-sky-500 transition-all font-semibold shadow-sm w-40"
                                            >
                                                <option value="all">All Locations</option>
                                                {reportLocations.map(loc => (
                                                    <option key={loc} value={loc}>{loc}</option>
                                                ))}
                                            </select>
                                        </div>

                                        <div className="flex flex-col">
                                            <label className="text-[8px] font-extrabold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-0.5">Plant Filter</label>
                                            <select
                                                value={filters.plant}
                                                onChange={(e) => setFilters({ ...filters, plant: e.target.value })}
                                                className="h-8 border border-slate-200/90 dark:border-slate-800 rounded-lg px-2.5 text-[11px] bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200 focus:outline-none focus:border-sky-500 transition-all font-semibold shadow-sm w-40"
                                            >
                                                <option value="all">All Plants</option>
                                                {reportLocationPlants.map(p => (
                                                    <option key={p.plant_code} value={p.plant_code}>{p.plant_name}</option>
                                                ))}
                                            </select>
                                        </div>

                                        <div className="flex flex-col">
                                            <label className="text-[8px] font-extrabold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-0.5">Range Filter</label>
                                            <div className="flex items-center bg-slate-200/70 dark:bg-slate-800 p-0.5 rounded-lg h-8">
                                                <button
                                                    type="button"
                                                    onClick={() => setReportRangeMode("custom")}
                                                    className={`px-2 py-1 text-[10px] font-bold rounded-md transition border-none cursor-pointer ${reportRangeMode === "custom" ? "bg-white dark:bg-slate-700 text-sky-700 dark:text-sky-300 shadow-sm" : "bg-transparent text-slate-600 dark:text-slate-400 hover:text-slate-900"}`}
                                                >
                                                    Custom Dates
                                                </button>
                                                <button
                                                    type="button"
                                                    onClick={() => setReportRangeMode("month")}
                                                    className={`px-2 py-1 text-[10px] font-bold rounded-md transition border-none cursor-pointer ${reportRangeMode === "month" ? "bg-white dark:bg-slate-700 text-sky-700 dark:text-sky-300 shadow-sm" : "bg-transparent text-slate-600 dark:text-slate-400 hover:text-slate-900"}`}
                                                >
                                                    Month Range
                                                </button>
                                            </div>
                                        </div>

                                        {reportRangeMode === "month" ? (
                                            <React.Fragment>
                                                <div className="flex flex-col">
                                                    <label className="text-[8px] font-extrabold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-0.5">From Month</label>
                                                    <input
                                                        type="month"
                                                        value={reportFromMonth}
                                                        onChange={(e) => setReportFromMonth(e.target.value)}
                                                        className="h-8 border border-slate-200/90 dark:border-slate-800 rounded-lg px-2 text-[11px] bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200 focus:outline-none focus:border-sky-500 font-semibold shadow-sm w-32"
                                                    />
                                                </div>
                                                <div className="flex flex-col">
                                                    <label className="text-[8px] font-extrabold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-0.5">To Month</label>
                                                    <input
                                                        type="month"
                                                        value={reportToMonth}
                                                        onChange={(e) => setReportToMonth(e.target.value)}
                                                        className="h-8 border border-slate-200/90 dark:border-slate-800 rounded-lg px-2 text-[11px] bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200 focus:outline-none focus:border-sky-500 font-semibold shadow-sm w-32"
                                                    />
                                                </div>
                                            </React.Fragment>
                                        ) : (
                                            <React.Fragment>
                                                <div className="flex flex-col">
                                                    <label className="text-[8px] font-extrabold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-0.5">From Date</label>
                                                    <input
                                                        type="date"
                                                        value={reportCustomStartDate || filters.startDate}
                                                        onChange={(e) => setReportCustomStartDate(e.target.value)}
                                                        className="h-8 border border-slate-200/90 dark:border-slate-800 rounded-lg px-2 text-[11px] bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200 focus:outline-none focus:border-sky-500 font-semibold shadow-sm w-32"
                                                    />
                                                </div>
                                                <div className="flex flex-col">
                                                    <label className="text-[8px] font-extrabold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-0.5">To Date</label>
                                                    <input
                                                        type="date"
                                                        value={reportCustomEndDate || filters.endDate}
                                                        onChange={(e) => setReportCustomEndDate(e.target.value)}
                                                        className="h-8 border border-slate-200/90 dark:border-slate-800 rounded-lg px-2 text-[11px] bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200 focus:outline-none focus:border-sky-500 font-semibold shadow-sm w-32"
                                                    />
                                                </div>
                                            </React.Fragment>
                                        )}

                                        <div className="flex flex-col">
                                            <label className="text-[8px] font-extrabold uppercase tracking-wider text-slate-400 dark:text-slate-500 mb-0.5">Select Report Type</label>
                                            <select
                                                value={selectedReportType}
                                                onChange={(e) => setSelectedReportType(e.target.value)}
                                                className="h-8 border border-slate-200/90 dark:border-slate-800 rounded-lg px-2.5 text-[11px] bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200 focus:outline-none focus:border-sky-500 transition-all font-bold shadow-sm w-56"
                                            >
                                                <option value="daily">Daily Operations Report</option>
                                                <option value="monthly">Monthly Utility Summary</option>
                                                <option value="utility">Utility-wise Consumption Report</option>
                                                <option value="waste">Waste Generation Report</option>
                                                <option value="targets">Target vs Actual Compliance</option>
                                            </select>
                                        </div>
                                    </div>
                                )}

                                {/* On-screen report preview only — never print these (Print uses Excel sheet layout) */}
                                <div className="no-print">
                                {/* Printable header */}
                                <div className="hidden print-only mb-4 border-b pb-2">
                                    <h1 className="text-xl font-bold text-slate-900 uppercase">UTILITY SENSE Analytics Report</h1>
                                    <p className="text-[10px] text-slate-400 mt-1">Generated: {new Date().toLocaleString()} · Security Class: Restricted</p>
                                </div>

                                {/* Render Report dynamically */}
                                {selectedReportType === "daily" && (
                                    <PrintableReportTable
                                        title="Daily Operations Statement"
                                        headers={["Date", "Plant", "Department", "Shift", "Operator", "Elect (kWh)", "Water (KL)", "Diesel (L)", "LPG (kg)", "Production"]}
                                        data={filteredEntries.map(e => [
                                            e.date, e.plant, e.department, e.shift, e.operator_name,
                                            fmtNum(e.electricity_consumption), fmtNum(e.water_consumption), fmtNum(e.diesel_used), fmtNum(e.lpg_used), fmtNum(e.production_qty)
                                        ])}
                                        totalRow={["Report Total", `${filteredEntries.length} rows`, "", "", "", fmtNum(kpiTotals.electricity), fmtNum(kpiTotals.water), fmtNum(kpiTotals.diesel), fmtNum(kpiTotals.lpg), fmtNum(kpiTotals.production)]}
                                    />
                                )}

                                {selectedReportType === "monthly" && (
                                    <PrintableReportTable
                                        title="Monthly Utility Aggregate"
                                        headers={["Period Month", "Electricity Load (kWh)", "Water Cons (KL)", "DG Fuel Used (L)", "Gas Used (kg)", "Estimated Energy Cost"]}
                                        data={monthlyTrendsData.map(m => [
                                            m.period, fmtNum(m.electricity), "—", "—", "—", fmtINR(m.cost)
                                        ])}
                                        totalRow={["Report Sum", fmtNum(kpiTotals.electricity), fmtNum(kpiTotals.water), fmtNum(kpiTotals.diesel), fmtNum(kpiTotals.lpg), fmtINR(aggregatedCosts.energyCost)]}
                                    />
                                )}

                                {selectedReportType === "utility" && (
                                    <PrintableReportTable
                                        title="Utility-wise Consumption Breakdown"
                                        headers={["Utility Source", "Cons/Used Volume", "Tariff Rate Average", "Unit", "Total cost (INR)"]}
                                        data={[
                                            [`Industrial Electricity (${reportGridLabel} Grid)`, fmtNum(kpiTotals.electricity), `₹ ${activeElectRate}`, "kWh", fmtINR(kpiTotals.electricityCost)],
                                            ["Municipal/Borewell Water", fmtNum(kpiTotals.water), `₹ ${activeWaterRate}`, "KL", fmtINR(kpiTotals.waterCost)],
                                            ["Diesel HSD generator fuel", fmtNum(kpiTotals.diesel), `₹ ${activeDieselRate}`, "Litre", fmtINR(kpiTotals.dieselCost)],
                                            ["LPG Gas utility", fmtNum(kpiTotals.lpg), `₹ ${activeLpgRate}`, "kg", fmtINR(kpiTotals.lpgCost)]
                                        ]}
                                        totalRow={["Sum Utilities Costs", "", "", "", fmtINR(aggregatedCosts.totalCost)]}
                                    />
                                )}

                                {selectedReportType === "waste" && (
                                    <PrintableReportTable
                                        title="Waste generation audit logs"
                                        headers={["Date", "Plant", "Hazardous Waste (kg)", "Non-Hazardous Waste (kg)", "Recycled Waste (kg)", "Recycling Recovery %"]}
                                        data={filteredEntries.map(e => [
                                            e.date, e.plant, fmtNum(e.waste_hazardous), fmtNum(e.waste_non_hazardous), fmtNum(e.waste_recycled),
                                            e.waste_non_hazardous > 0 ? (e.waste_recycled / (e.waste_hazardous + e.waste_non_hazardous) * 100).toFixed(0) + "%" : "0%"
                                        ])}
                                        totalRow={["Report Total", "", fmtNum(kpiTotals.wasteHaz), fmtNum(kpiTotals.wasteNHaz), fmtNum(kpiTotals.wasteRec), ""]}
                                    />
                                )}

                                {selectedReportType === "targets" && (
                                    <PrintableReportTable
                                        title="Target vs Actual operations compliance"
                                        headers={["Operating Plant", "Electricity Target (kWh)", "Electricity Actual (kWh)", "Status", "Water Target (KL)", "Water Actual (KL)"]}
                                        data={plants.map(p => {
                                            const targetRow = targets.find(t => t.plant_code === p.plant_code);
                                            const matchRows = filteredEntries.filter(x => x.plant === p.plant_code);
                                            const electActual = matchRows.reduce((a, c) => a + Number(c.electricity_consumption), 0);
                                            const waterActual = matchRows.reduce((a, c) => a + Number(c.water_consumption), 0);
                                            const eTarget = targetRow ? Number(targetRow.electricity_target_kwh) * 20 : 10000;
                                            const wTarget = targetRow ? Number(targetRow.water_target_kl) * 20 : 300;
                                            return [
                                                p.plant_code, fmtNum(eTarget), fmtNum(electActual),
                                                electActual > eTarget ? "âš ï¸ Exceeded" : "✅ Compliant",
                                                fmtNum(wTarget), fmtNum(waterActual)
                                            ];
                                        })}
                                    />
                                )}
                                </div>
                            </div>
                        )}

                        {/* 4. TARGETS COMPONENT */}
                        {activeTab === "targets" && (
                            <div className="space-y-6 pt-4">
                                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b pb-4">
                                    <div>
                                        <h2 className="text-base font-extrabold text-slate-900 flex items-center gap-1.5 uppercase">
                                            <span className="material-symbols-outlined text-[#0284c7]">track_changes</span>
                                            <span>Monthly Targets Configurations</span>
                                        </h2>
                                        <p className="text-xs text-slate-400 mt-0.5">Setup target limits for plants to trigger auto anomaly email alerts</p>
                                    </div>

                                    <button
                                        onClick={() => { setActiveTab("master"); setSelectedMasterTable("target_values"); }}
                                        className="px-4 py-2 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-bold cursor-pointer transition flex items-center gap-1.5"
                                    >
                                        <span className="material-symbols-outlined text-[15px]">edit</span>
                                        <span>Edit Targets Config</span>
                                    </button>
                                </div>

                                {/* Targets display grids */}
                                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                                    {targets.map(t => (
                                        <div key={t.target_id} className={`bg-white dark:bg-[#121a29] rounded-2xl border-l border-r border-b border-slate-200/90 dark:border-[#26334a] p-5 shadow-[0_3px_10px_-3px_rgba(15,23,42,0.06)] transition-all hover:shadow-md hover:-translate-y-1.5 duration-300 ${t.status === 'Active' ? 'border-t-4 border-t-emerald-500' : 'border-t-4 border-t-slate-400'}`}>
                                            <div className="flex items-center justify-between mb-4 pb-3 border-b border-slate-100 dark:border-slate-800">
                                                <div>
                                                    <h3 className="text-sm font-extrabold text-slate-900 dark:text-slate-100 uppercase tracking-tight">{t.plant_code} · {t.department_code}</h3>
                                                    <p className="text-[10px] text-slate-400 font-bold font-mono mt-0.5 uppercase">Month: {MONTHS[t.month - 1]} {t.year}</p>
                                                </div>
                                                <span className={`px-2.5 py-0.5 text-[9px] font-extrabold uppercase rounded-full ${t.status === "Active" ? "bg-emerald-50 text-emerald-700 border border-emerald-200/80" : "bg-slate-50 text-slate-500 border border-slate-200/80"}`}>
                                                    {t.status}
                                                </span>
                                            </div>

                                            <div className="space-y-3">
                                                <div className="flex items-center justify-between bg-slate-50/50 dark:bg-slate-900/20 p-2 rounded-xl border border-slate-100 dark:border-slate-800/40">
                                                    <div className="flex items-center gap-2">
                                                        <span className="material-symbols-outlined text-[16px] text-sky-500 font-bold">electric_bolt</span>
                                                        <span className="text-xs font-semibold text-slate-500 dark:text-slate-400">Electricity Target:</span>
                                                    </div>
                                                    <span className="text-xs font-extrabold text-slate-900 dark:text-white font-mono bg-sky-50 dark:bg-sky-950/20 px-2 py-0.5 rounded border border-sky-100/50 dark:border-sky-900/20">{fmtNum(t.electricity_target_kwh)} kWh</span>
                                                </div>

                                                <div className="flex items-center justify-between bg-slate-50/50 dark:bg-slate-900/20 p-2 rounded-xl border border-slate-100 dark:border-slate-800/40">
                                                    <div className="flex items-center gap-2">
                                                        <span className="material-symbols-outlined text-[16px] text-teal-500 font-bold">water_drop</span>
                                                        <span className="text-xs font-semibold text-slate-500 dark:text-slate-400">Water Limit:</span>
                                                    </div>
                                                    <span className="text-xs font-extrabold text-slate-900 dark:text-white font-mono bg-teal-50 dark:bg-teal-950/20 px-2 py-0.5 rounded border border-teal-100/50 dark:border-teal-900/20">{fmtNum(t.water_target_kl)} KL</span>
                                                </div>

                                                <div className="flex items-center justify-between bg-slate-50/50 dark:bg-slate-900/20 p-2 rounded-xl border border-slate-100 dark:border-slate-800/40">
                                                    <div className="flex items-center gap-2">
                                                        <span className="material-symbols-outlined text-[16px] text-red-500 font-bold">local_gas_station</span>
                                                        <span className="text-xs font-semibold text-slate-500 dark:text-slate-400">Diesel DG Limit:</span>
                                                    </div>
                                                    <span className="text-xs font-extrabold text-slate-900 dark:text-white font-mono bg-red-50 dark:bg-red-950/20 px-2 py-0.5 rounded border border-red-100/50 dark:border-red-900/20">{fmtNum(t.diesel_target_l)} L</span>
                                                </div>

                                                <div className="flex items-center justify-between bg-slate-50/50 dark:bg-slate-900/20 p-2 rounded-xl border border-slate-100 dark:border-slate-800/40">
                                                    <div className="flex items-center gap-2">
                                                        <span className="material-symbols-outlined text-[16px] text-pink-500 font-bold">propane_tank</span>
                                                        <span className="text-xs font-semibold text-slate-500 dark:text-slate-400">LPG Gas Limit:</span>
                                                    </div>
                                                    <span className="text-xs font-extrabold text-slate-900 dark:text-white font-mono bg-pink-50 dark:bg-pink-950/20 px-2 py-0.5 rounded border border-pink-100/50 dark:border-pink-900/20">{fmtNum(t.lpg_target_kg)} kg</span>
                                                </div>

                                                <div className="flex items-center justify-between bg-slate-50/50 dark:bg-slate-900/20 p-2 rounded-xl border border-slate-100 dark:border-slate-800/40">
                                                    <div className="flex items-center gap-2">
                                                        <span className="material-symbols-outlined text-[16px] text-amber-500 font-bold">wb_sunny</span>
                                                        <span className="text-xs font-semibold text-slate-500 dark:text-slate-400">Solar Gen Min:</span>
                                                    </div>
                                                    <span className="text-xs font-extrabold text-amber-700 dark:text-amber-400 font-mono bg-amber-50 dark:bg-amber-950/20 px-2 py-0.5 rounded border border-amber-100/50 dark:border-amber-900/20">{fmtNum(t.solar_generation_target_kwh)} kWh</span>
                                                </div>
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        )}

                        {/* 5. ALERTS COMPONENT */}
                        {activeTab === "alerts" && (
                            <div className="space-y-6 pt-4">
                                <div className="border-b pb-4">
                                    <h2 className="text-base font-extrabold text-slate-900 flex items-center gap-1.5 uppercase">
                                        <span className="material-symbols-outlined text-[#0284c7]">notifications_active</span>
                                        <span>Utility Exceedance Alerts Log</span>
                                    </h2>
                                    <p className="text-xs text-slate-400 mt-0.5">Real-time target limits alerts and duplicate entries warning logs</p>
                                </div>

                                {/* Alerts List */}
                                <div className="space-y-3.5">
                                    {alertsList.length === 0 ? (
                                        <div className="bg-emerald-50 border border-emerald-100 rounded-2xl p-5 text-center text-emerald-800 text-xs">
                                            <span className="material-symbols-outlined text-[24px] mb-1">check_circle</span>
                                            <p className="font-bold">System Status Healthy</p>
                                            <p className="text-[10px] text-emerald-600 mt-0.5">No target limit exceedances or duplicates detected in operations entries</p>
                                        </div>
                                    ) : (
                                        alertsList.map((alert, idx) => (
                                            <div key={idx} className={`flex items-start gap-4 p-4 border rounded-2xl bg-white shadow-sm transition hover:shadow-md ${alert.level === 'Critical' ? 'border-red-200' : 'border-amber-200'}`}>
                                                <span className={`material-symbols-outlined text-[24px] ${alert.level === 'Critical' ? 'text-red-500' : 'text-amber-500'}`}>
                                                    {alert.level === 'Critical' ? 'emergency' : 'warning'}
                                                </span>
                                                <div className="flex-1">
                                                    <div className="flex items-center justify-between gap-4">
                                                        <h4 className="text-xs font-bold text-slate-900">{alert.type}</h4>
                                                        <span className={`px-2 py-0.5 text-[9px] font-extrabold uppercase rounded-full ${alert.level === 'Critical' ? 'bg-red-50 text-red-700 border border-red-100' : 'bg-amber-50 text-amber-700 border border-amber-100'}`}>
                                                            {alert.level}
                                                        </span>
                                                    </div>
                                                    <p className="text-xs text-slate-500 mt-1 leading-relaxed">{alert.message}</p>
                                                    <p className="text-[9px] text-slate-400 font-mono mt-1.5">{alert.date}</p>
                                                </div>
                                            </div>
                                        ))
                                    )}
                                </div>
                            </div>
                        )}

                        {/* 6. MASTER CONFIGS TABLE EDITOR COMPONENT */}
                        {activeTab === "master" && (
                            <div className="space-y-6 pt-4">
                                <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-3 border-b pb-4">
                                    <div className="shrink-0">
                                        <h2 className="text-base font-extrabold text-slate-900 flex items-center gap-1.5 uppercase whitespace-nowrap">
                                            <span className="material-symbols-outlined text-[#0284c7]">settings</span>
                                            <span>
                                                {selectedMasterTable === "users" ? "User Management" :
                                                 selectedMasterTable === "email_automation" ? "Automatic Email Reports" :
                                                 selectedMasterTable === "audit_logs" ? "Security & Operations Audit Logs" :
                                                 selectedMasterTable === "multiply_factors" ? "Multiply Factor Settings" :
                                                 selectedMasterTable === "tariff_rates" ? "Tariff Rates Settings" : "Master Configuration Settings"}
                                            </span>
                                        </h2>
                                        <p className="text-xs text-slate-400 mt-0.5">
                                            {currentUser.role === "IT_ADMIN"
                                                ? "Corporate governance, automated reports dispatch and audit traceability"
                                                : "Manage tariff rates and multiply factor for your assigned location/plant"}
                                        </p>
                                    </div>

                                    {/* All header controls in ONE single horizontal line */}
                                    <div className="flex items-center gap-1.5 flex-nowrap overflow-x-auto pb-0.5 max-w-full">
                                        {/* Theme Picker */}
                                        <div className="relative shrink-0">
                                            <button
                                                onClick={() => { setThemePickerOpen(o => !o); setIdlePickerOpen(false); }}
                                                className="flex items-center gap-1.5 h-9 px-2.5 border border-slate-200 rounded-xl bg-white text-xs font-bold text-slate-700 hover:bg-slate-50 cursor-pointer shadow-sm whitespace-nowrap"
                                                title="Change Theme"
                                            >
                                                <span className="material-symbols-outlined text-[17px] text-amber-500">palette</span>
                                                <span>Theme</span>
                                            </button>
                                            {themePickerOpen && (
                                                <React.Fragment>
                                                    <div className="fixed inset-0 z-40" onClick={() => setThemePickerOpen(false)}></div>
                                                    <div className="theme-popover absolute right-0 top-11 z-50 w-52 rounded-2xl border border-slate-200 bg-white shadow-lg p-2.5">
                                                        <p className="text-[9px] font-bold uppercase tracking-wider text-slate-400 px-2 pb-2 pt-1">Choose Theme</p>
                                                        {THEME_OPTIONS.map(opt => (
                                                            <button
                                                                key={opt.id}
                                                                onClick={() => { setTheme(opt.id); setThemePickerOpen(false); }}
                                                                className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-xl text-xs font-semibold transition cursor-pointer ${theme === opt.id ? "bg-sky-50 text-sky-700" : "text-slate-600 hover:bg-slate-50"}`}
                                                            >
                                                                <span
                                                                    className="h-5 w-5 rounded-full border border-black/10 shrink-0"
                                                                    style={{ background: opt.swatch }}
                                                                ></span>
                                                                {opt.label}
                                                                {theme === opt.id && (
                                                                    <span className="material-symbols-outlined text-[16px] ml-auto text-sky-600">check</span>
                                                                )}
                                                            </button>
                                                        ))}
                                                    </div>
                                                </React.Fragment>
                                            )}
                                        </div>

                                        {/* Idle Screen Saver Selector */}
                                        <div className="relative shrink-0">
                                            <button
                                                onClick={() => { setIdlePickerOpen(o => !o); setThemePickerOpen(false); }}
                                                className="flex items-center gap-1.5 h-9 px-2.5 border border-slate-200 rounded-xl bg-white text-xs font-bold text-slate-700 hover:bg-slate-50 cursor-pointer shadow-sm whitespace-nowrap"
                                                title="Idle Screensaver Timeout"
                                            >
                                                <span className="material-symbols-outlined text-[17px] text-sky-500">timer</span>
                                                <span>Screensaver ({idleMinutes}m)</span>
                                            </button>
                                            {idlePickerOpen && (
                                                <React.Fragment>
                                                    <div className="fixed inset-0 z-40" onClick={() => setIdlePickerOpen(false)}></div>
                                                    <div className="absolute right-0 top-11 z-50 w-48 rounded-2xl border border-slate-200 bg-white shadow-lg p-2.5">
                                                        <p className="text-[9px] font-bold uppercase tracking-wider text-slate-400 px-2 pb-2 pt-1">Idle Screen</p>
                                                        {IDLE_MINUTE_OPTIONS.map(opt => (
                                                            <button
                                                                key={opt.value}
                                                                type="button"
                                                                onClick={() => { setIdleMinutes(opt.value); setIdlePickerOpen(false); }}
                                                                className={`w-full flex items-center gap-2 px-2.5 py-2 rounded-xl text-xs font-semibold transition cursor-pointer border-none ${idleMinutes === opt.value ? "bg-sky-50 text-sky-700" : "bg-transparent text-slate-600 hover:bg-slate-50"}`}
                                                            >
                                                                {opt.label}
                                                                {idleMinutes === opt.value && (
                                                                    <span className="material-symbols-outlined text-[16px] ml-auto text-sky-600">check</span>
                                                                )}
                                                            </button>
                                                        ))}
                                                    </div>
                                                </React.Fragment>
                                            )}
                                        </div>

                                        {selectedMasterTable !== "users" && (
                                            <select
                                                value={selectedMasterTable}
                                                onChange={(e) => setSelectedMasterTable(e.target.value)}
                                                className="h-9 border border-slate-200 rounded-xl px-2 text-xs bg-[#f8fafc] font-bold text-slate-700 focus:outline-none shrink-0"
                                            >
                                                {currentUser.role === "IT_ADMIN" ? (
                                                    <React.Fragment>
                                                        <option value="plants">Plants Node</option>
                                                        <option value="tariff_rates">Tariff Rates (₹)</option>
                                                        <option value="multiply_factors">Multiply Factor (MF)</option>
                                                        <option value="target_values">Target Values Settings</option>
                                                        <option value="email_automation">Automatic Email Reports</option>
                                                        <option value="audit_logs">Audit Logs (Security & Activity)</option>
                                                    </React.Fragment>
                                                ) : (
                                                    <React.Fragment>
                                                        <option value="tariff_rates">Tariff Rates (₹)</option>
                                                        <option value="multiply_factors">Multiply Factor (MF)</option>
                                                    </React.Fragment>
                                                )}
                                            </select>
                                        )}

                                        {currentUser.role === "IT_ADMIN" && (
                                            <button
                                                onClick={() => setSelectedMasterTable(selectedMasterTable === "email_automation" ? "plants" : "email_automation")}
                                                className={`flex items-center gap-1.5 h-9 px-2.5 border rounded-xl text-xs font-bold cursor-pointer shadow-sm transition shrink-0 whitespace-nowrap ${selectedMasterTable === "email_automation" ? "bg-sky-50 border-sky-200 text-sky-700" : "bg-white border-slate-200 text-slate-700 hover:bg-slate-50"}`}
                                                title="Automatic Location/Plant-wise Scheduled Email Reports"
                                            >
                                                <span className="material-symbols-outlined text-[17px]">forward_to_inbox</span>
                                                <span>Auto Email</span>
                                            </button>
                                        )}

                                        {currentUser.role === "IT_ADMIN" && (
                                            <button
                                                onClick={() => setSelectedMasterTable(selectedMasterTable === "audit_logs" ? "plants" : "audit_logs")}
                                                className={`flex items-center gap-1.5 h-9 px-2.5 border rounded-xl text-xs font-bold cursor-pointer shadow-sm transition shrink-0 whitespace-nowrap ${selectedMasterTable === "audit_logs" ? "bg-indigo-50 border-indigo-200 text-indigo-700" : "bg-white border-slate-200 text-slate-700 hover:bg-slate-50"}`}
                                                title="Tamper-Resistant Audit Logs"
                                            >
                                                <span className="material-symbols-outlined text-[17px]">history</span>
                                                <span>Audit Logs</span>
                                            </button>
                                        )}

                                        {currentUser.role === "IT_ADMIN" && (
                                            <button
                                                onClick={() => setSelectedMasterTable(selectedMasterTable === "users" ? "plants" : "users")}
                                                className={`flex items-center gap-1.5 h-9 px-2.5 border rounded-xl text-xs font-bold cursor-pointer shadow-sm transition shrink-0 whitespace-nowrap ${selectedMasterTable === "users" ? "bg-sky-50 border-sky-200 text-sky-700" : "bg-white border-slate-200 text-slate-700 hover:bg-slate-50"}`}
                                                title="User Management"
                                            >
                                                <span className="material-symbols-outlined text-[17px]">group</span>
                                                <span>Users</span>
                                            </button>
                                        )}

                                        {selectedMasterTable !== "audit_logs" && selectedMasterTable !== "email_automation" && (
                                            <React.Fragment>
                                                {currentUser.role === "IT_ADMIN" && (
                                                    !isMasterDeleteMode ? (
                                                        <button
                                                            onClick={() => setIsMasterDeleteMode(true)}
                                                            className="flex items-center gap-1 h-9 px-2.5 bg-white border border-red-200 hover:bg-red-50 text-red-600 rounded-xl text-xs font-bold transition shadow-sm cursor-pointer shrink-0 whitespace-nowrap"
                                                            title="Click to enable selection mode for deleting master records"
                                                        >
                                                            <span className="material-symbols-outlined text-[16px]">delete_sweep</span>
                                                            <span>Bulk Delete</span>
                                                        </button>
                                                    ) : (
                                                        <div className="flex items-center gap-1 shrink-0">
                                                            {selectedMasterRowKeys.size > 0 && (
                                                                <button
                                                                    onClick={handleBulkDeleteMaster}
                                                                    className="flex items-center gap-1 h-9 px-2.5 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold transition shadow-sm border-none cursor-pointer whitespace-nowrap"
                                                                    title={`Delete ${selectedMasterRowKeys.size} selected master records`}
                                                                >
                                                                    <span className="material-symbols-outlined text-[16px]">delete</span>
                                                                    <span>Confirm Delete ({selectedMasterRowKeys.size})</span>
                                                                </button>
                                                            )}
                                                            <button
                                                                onClick={() => { setIsMasterDeleteMode(false); setSelectedMasterRowKeys(new Set()); }}
                                                                className="flex items-center gap-1 h-9 px-2 bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 rounded-xl text-xs font-bold transition shadow-sm cursor-pointer whitespace-nowrap"
                                                                title="Cancel selection mode"
                                                            >
                                                                <span>Cancel</span>
                                                            </button>
                                                        </div>
                                                    )
                                                )}

                                                <button
                                                    onClick={() => openMasterForm()}
                                                    className="flex items-center gap-1 h-9 px-3 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl text-xs font-bold cursor-pointer transition border-none shadow-sm shrink-0 whitespace-nowrap"
                                                >
                                                    <span className="material-symbols-outlined text-[16px]">add</span>
                                                    <span>Add Record</span>
                                                </button>
                                            </React.Fragment>
                                        )}
                                    </div>
                                </div>

                                {/* Config Records Table / Specialized Panels */}
                                {selectedMasterTable === "email_automation" && currentUser.role === "IT_ADMIN" ? (
                                    <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden space-y-0">
                                        {/* Multi-Schedule Header Bar */}
                                        <div className="p-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4 bg-slate-50/50">
                                            <div>
                                                <h3 className="text-sm font-extrabold text-slate-900 flex items-center gap-2">
                                                    <span className="material-symbols-outlined text-sky-600">forward_to_inbox</span>
                                                    <span>Automatic Location/Plant-wise Scheduled Email Reports</span>
                                                </h3>
                                                <p className="text-xs text-slate-500 mt-0.5">
                                                    Configure multiple independent monthly Excel report dispatches per location & plant with custom To/CC recipients.
                                                </p>
                                            </div>
                                            <div className="flex items-center gap-2">
                                                <button
                                                    type="button"
                                                    onClick={() => setIsScheduleLogsOpen(true)}
                                                    className="flex items-center gap-1.5 h-9 px-3.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold transition shadow-sm cursor-pointer whitespace-nowrap"
                                                    title="View history of scheduled report dispatches"
                                                >
                                                    <span className="material-symbols-outlined text-[17px] text-indigo-600">list_alt</span>
                                                    <span>Execution Logs ({emailScheduleLogs.length})</span>
                                                </button>
                                                <button
                                                    type="button"
                                                    onClick={openCreateScheduleModal}
                                                    className="flex items-center gap-1.5 h-9 px-4 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl text-xs font-bold transition shadow-sm border-none cursor-pointer whitespace-nowrap"
                                                >
                                                    <span className="material-symbols-outlined text-[17px]">add</span>
                                                    <span>Create Email Schedule</span>
                                                </button>
                                            </div>
                                        </div>

                                        {/* Schedules Table */}
                                        <div className="overflow-x-auto">
                                            <table className="w-full border-collapse text-left text-xs">
                                                <thead>
                                                    <tr className="border-b border-slate-100 text-slate-400 font-bold uppercase tracking-wider text-[9px] bg-slate-50/70">
                                                        <th className="py-3 px-4">Schedule</th>
                                                        <th className="py-3 px-4">Location</th>
                                                        <th className="py-3 px-4">Plant</th>
                                                        <th className="py-3 px-4">To Recipients</th>
                                                        <th className="py-3 px-4">CC Recipients</th>
                                                        <th className="py-3 px-4">Timing</th>
                                                        <th className="py-3 px-4">Status</th>
                                                        <th className="py-3 px-4">Last Run</th>
                                                        <th className="py-3 px-4 text-right">Actions</th>
                                                    </tr>
                                                </thead>
                                                <tbody className="divide-y divide-slate-100 text-slate-700">
                                                    {emailSchedules.length === 0 ? (
                                                        <tr>
                                                            <td colSpan={9} className="py-12 text-center text-slate-400 font-medium">
                                                                <div className="flex flex-col items-center gap-2">
                                                                    <span className="material-symbols-outlined text-[36px] text-slate-300">mail_off</span>
                                                                    <p>No automated email schedules configured yet.</p>
                                                                    <button
                                                                        type="button"
                                                                        onClick={openCreateScheduleModal}
                                                                        className="mt-1 px-3 py-1.5 bg-sky-50 text-sky-700 hover:bg-sky-100 font-bold rounded-xl text-xs border border-sky-200 cursor-pointer"
                                                                    >
                                                                        + Create First Schedule
                                                                    </button>
                                                                </div>
                                                            </td>
                                                        </tr>
                                                    ) : (
                                                        emailSchedules.map((sched) => {
                                                            const toCount = (sched.to_recipients || "").split(",").map(x => x.trim()).filter(Boolean).length;
                                                            const ccCount = (sched.cc_recipients || "").split(",").map(x => x.trim()).filter(Boolean).length;
                                                            const isRunning = runningScheduleId === sched.id;

                                                            return (
                                                                <tr key={sched.id} className="hover:bg-slate-50/60 transition">
                                                                    <td className="py-3 px-4">
                                                                        <div className="font-bold text-slate-900">{sched.name || sched.id}</div>
                                                                        <div className="text-[10px] text-slate-400 font-mono">{sched.report_type || "Monthly Utility Report"}</div>
                                                                    </td>
                                                                    <td className="py-3 px-4 font-bold text-slate-800 uppercase">
                                                                        {sched.location}
                                                                    </td>
                                                                    <td className="py-3 px-4 font-bold text-sky-700">
                                                                        {sched.plant}
                                                                    </td>
                                                                    <td className="py-3 px-4" title={sched.to_recipients}>
                                                                        <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-sky-50 text-sky-800 border border-sky-100">
                                                                            {toCount} recipient{toCount !== 1 ? 's' : ''}
                                                                        </span>
                                                                    </td>
                                                                    <td className="py-3 px-4" title={sched.cc_recipients}>
                                                                        {ccCount > 0 ? (
                                                                            <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-700">
                                                                                {ccCount} CC
                                                                            </span>
                                                                        ) : (
                                                                            <span className="text-[10px] text-slate-300">—</span>
                                                                        )}
                                                                    </td>
                                                                    <td className="py-3 px-4 font-mono text-[11px] text-slate-600 whitespace-nowrap">
                                                                        Day {sched.schedule_day || 1} @ {sched.schedule_time || "10:00"}
                                                                    </td>
                                                                    <td className="py-3 px-4">
                                                                        <label className="relative inline-flex items-center cursor-pointer">
                                                                            <input
                                                                                type="checkbox"
                                                                                checked={sched.enabled !== false}
                                                                                onChange={() => handleToggleScheduleStatus(sched.id, sched.enabled !== false)}
                                                                                className="sr-only peer"
                                                                            />
                                                                            <div className="w-8 h-4.5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-3.5 after:w-3.5 after:transition-all peer-checked:bg-emerald-500"></div>
                                                                            <span className={`ml-2 text-[10px] font-extrabold ${sched.enabled !== false ? "text-emerald-600" : "text-slate-400"}`}>
                                                                                {sched.enabled !== false ? "ON" : "OFF"}
                                                                            </span>
                                                                        </label>
                                                                    </td>
                                                                    <td className="py-3 px-4 whitespace-nowrap text-[10.5px]">
                                                                        {sched.last_run_at ? (
                                                                            <div>
                                                                                <span className={`px-1.5 py-0.5 rounded text-[9px] font-extrabold mr-1 ${sched.last_status === "SUCCESS" ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>
                                                                                    {sched.last_status || "DONE"}
                                                                                </span>
                                                                                <span className="text-slate-400 font-mono">{new Date(sched.last_run_at).toLocaleDateString("en-IN")}</span>
                                                                            </div>
                                                                        ) : (
                                                                            <span className="text-slate-300">Never executed</span>
                                                                        )}
                                                                    </td>
                                                                    <td className="py-3 px-4 text-right space-x-1.5 whitespace-nowrap">
                                                                        <button
                                                                            type="button"
                                                                            disabled={isRunning || isEmailSending}
                                                                            onClick={() => handleExecuteSchedule({ schedule: sched, isManualTest: true })}
                                                                            className="px-2.5 py-1 text-[10.5px] font-extrabold text-sky-700 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded-lg cursor-pointer transition disabled:opacity-50 inline-flex items-center gap-1"
                                                                            title="Dispatch test report email right now"
                                                                        >
                                                                            {isRunning ? <span className="w-3 h-3 border-2 border-sky-600 border-t-transparent rounded-full animate-spin" /> : <span className="material-symbols-outlined text-[13px]">send</span>}
                                                                            <span>Run Now</span>
                                                                        </button>
                                                                        <button
                                                                            type="button"
                                                                            onClick={() => openEditScheduleModal(sched)}
                                                                            className="px-2 py-1 text-[10.5px] font-extrabold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg cursor-pointer border-none transition"
                                                                        >
                                                                            Edit
                                                                        </button>
                                                                        <button
                                                                            type="button"
                                                                            onClick={() => handleDeleteSchedule(sched.id)}
                                                                            className="px-2 py-1 text-[10.5px] font-extrabold text-red-600 bg-red-50 hover:bg-red-100 rounded-lg cursor-pointer border-none transition"
                                                                        >
                                                                            Delete
                                                                        </button>
                                                                    </td>
                                                                </tr>
                                                            );
                                                        })
                                                    )}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                ) : selectedMasterTable === "audit_logs" && currentUser.role === "IT_ADMIN" ? (
                                    <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden space-y-0">
                                        {/* Audit Log Controls Bar */}
                                        <div className="p-4 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-50/50">
                                            <div className="flex items-center gap-2 flex-1">
                                                <div className="relative flex-1 max-w-xs">
                                                    <span className="material-symbols-outlined absolute left-2.5 top-2 text-slate-400 text-[18px]">search</span>
                                                    <input
                                                        type="text"
                                                        placeholder="Search audit logs by user, action, module..."
                                                        value={auditSearchQuery}
                                                        onChange={(e) => setAuditSearchQuery(e.target.value)}
                                                        className="w-full h-8 pl-8 pr-3 bg-white border border-slate-200 rounded-xl text-xs text-slate-700 focus:outline-none focus:border-sky-500 shadow-sm"
                                                    />
                                                </div>
                                                <select
                                                    value={auditModuleFilter}
                                                    onChange={(e) => setAuditModuleFilter(e.target.value)}
                                                    className="h-8 border border-slate-200 rounded-xl px-2.5 text-xs bg-white text-slate-700 font-bold focus:outline-none shadow-sm"
                                                >
                                                    <option value="all">All Modules</option>
                                                    <option value="Daily Operations">Daily Operations</option>
                                                    <option value="Reports">Reports</option>
                                                    <option value="Email Automation">Email Automation</option>
                                                    <option value="Auth">Authentication</option>
                                                    <option value="plants">Plants Node</option>
                                                    <option value="tariff_rates">Tariff Rates</option>
                                                    <option value="multiply_factors">Multiply Factors</option>
                                                    <option value="users">User Management</option>
                                                </select>
                                            </div>
                                            <div className="text-[11px] text-slate-400 font-semibold">
                                                Total Records: <span className="text-slate-700 font-bold">{auditLogs.length}</span> (Read-Only & Tamper-Resistant)
                                            </div>
                                        </div>

                                        {/* Audit Logs Table */}
                                        <div className="overflow-x-auto">
                                            <table className="w-full border-collapse text-left text-xs">
                                                <thead>
                                                    <tr className="border-b border-slate-100 text-slate-400 font-bold uppercase tracking-wider text-[9px] bg-slate-50/70">
                                                        <th className="py-3 px-4">Timestamp</th>
                                                        <th className="py-3 px-4">User</th>
                                                        <th className="py-3 px-4">Action</th>
                                                        <th className="py-3 px-4">Module</th>
                                                        <th className="py-3 px-4">Record / Target</th>
                                                        <th className="py-3 px-4">Status</th>
                                                        <th className="py-3 px-4 text-right">Details</th>
                                                    </tr>
                                                </thead>
                                                <tbody className="divide-y divide-slate-100 text-slate-700">
                                                    {(() => {
                                                        const filtered = auditLogs.filter(log => {
                                                            if (auditModuleFilter !== "all" && log.module !== auditModuleFilter) return false;
                                                            if (auditSearchQuery) {
                                                                const q = auditSearchQuery.toLowerCase();
                                                                const text = `${log.user_email} ${log.user_name} ${log.action} ${log.module} ${log.record_id} ${log.location} ${log.plant}`.toLowerCase();
                                                                if (!text.includes(q)) return false;
                                                            }
                                                            return true;
                                                        });

                                                        if (filtered.length === 0) {
                                                            return (
                                                                <tr>
                                                                    <td colSpan={7} className="py-12 text-center text-slate-400 font-medium">
                                                                        No audit logs recorded for this query.
                                                                    </td>
                                                                </tr>
                                                            );
                                                        }

                                                        return filtered.slice(0, 100).map((log, idx) => (
                                                            <tr key={idx} className="hover:bg-slate-50/50 transition">
                                                                <td className="py-2.5 px-4 font-mono text-[11px] text-slate-500 whitespace-nowrap">
                                                                    {log.created_at ? new Date(log.created_at).toLocaleString("en-IN") : "—"}
                                                                </td>
                                                                <td className="py-2.5 px-4">
                                                                    <div className="font-bold text-slate-800">{log.user_name || log.user_email}</div>
                                                                    <div className="text-[10px] text-slate-400">{log.user_email}</div>
                                                                </td>
                                                                <td className="py-2.5 px-4">
                                                                    <span className={`px-2 py-0.5 rounded text-[9px] font-extrabold uppercase ${
                                                                        log.action === "CREATE" ? "bg-emerald-50 text-emerald-700 border border-emerald-100" :
                                                                        log.action === "UPDATE" ? "bg-sky-50 text-sky-700 border border-sky-100" :
                                                                        log.action === "DELETE" ? "bg-red-50 text-red-700 border border-red-100" :
                                                                        log.action === "EXPORT" ? "bg-purple-50 text-purple-700 border border-purple-100" :
                                                                        log.action.includes("LOGIN") ? "bg-indigo-50 text-indigo-700 border border-indigo-100" :
                                                                        "bg-amber-50 text-amber-700 border border-amber-100"
                                                                    }`}>
                                                                        {log.action}
                                                                    </span>
                                                                </td>
                                                                <td className="py-2.5 px-4 font-bold text-slate-700 text-[11px]">{log.module}</td>
                                                                <td className="py-2.5 px-4 font-mono text-[11px] text-slate-600">
                                                                    {log.plant ? `${log.plant} (${log.location || ""})` : log.record_id || "—"}
                                                                </td>
                                                                <td className="py-2.5 px-4">
                                                                    <span className={`px-2 py-0.5 rounded text-[9px] font-bold ${log.status === "SUCCESS" ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>
                                                                        {log.status}
                                                                    </span>
                                                                </td>
                                                                <td className="py-2.5 px-4 text-right">
                                                                    <button
                                                                        type="button"
                                                                        onClick={() => setSelectedAuditDetail(log)}
                                                                        className="px-2.5 py-1 text-[10px] font-extrabold text-sky-600 hover:text-sky-700 bg-sky-50 hover:bg-sky-100 rounded-lg cursor-pointer border-none transition"
                                                                    >
                                                                        Inspect
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        ));
                                                    })()}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                ) : (
                                <div className="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
                                    <div className="overflow-x-auto">
                                        <table className="w-full border-collapse text-left text-xs">
                                            <thead>
                                                <tr className="border-b border-slate-100 text-slate-400 font-bold uppercase tracking-wider text-[9px] bg-slate-50/50">
                                                    {currentUser.role === "IT_ADMIN" && isMasterDeleteMode && (
                                                        <th className="py-3 px-4 w-8">
                                                            {(() => {
                                                                const masterList = getFilteredMasterList(selectedMasterTable);
                                                                const pkCol = getTablePrimaryKey(selectedMasterTable);
                                                                const visibleKeys = masterList.map(row => String(row[pkCol])).filter(Boolean);
                                                                const isAllSelected = visibleKeys.length > 0 && visibleKeys.every(k => selectedMasterRowKeys.has(k));
                                                                return (
                                                                    <input
                                                                        type="checkbox"
                                                                        checked={isAllSelected}
                                                                        onChange={handleToggleSelectAllMaster}
                                                                        className="rounded text-sky-600 focus:ring-sky-500 h-4 w-4 cursor-pointer"
                                                                        title="Select all visible master records"
                                                                    />
                                                                );
                                                            })()}
                                                        </th>
                                                    )}
                                                    {getMasterTableHeaders(selectedMasterTable).map((h, idx) => {
                                                        let label = h.replace("_", " ");
                                                        if (selectedMasterTable === "users") {
                                                            if (h === "id") label = "Employee ID";
                                                            else if (h === "email") label = "Company Email";
                                                            else if (h === "allowed_locations") label = "Locations Access";
                                                            else if (h === "allowed_plants") label = "Plants Access";
                                                        }
                                                        return (
                                                            <th key={idx} className="py-3 px-4 uppercase text-[9px] font-bold tracking-wider">{label}</th>
                                                        );
                                                    })}
                                                    <th className="py-3 px-4 text-right">Actions</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-100 text-slate-700">
                                                {getFilteredMasterList(selectedMasterTable).length === 0 ? (
                                                    <tr>
                                                        <td colSpan={getMasterTableHeaders(selectedMasterTable).length + 1} className="py-8 text-center text-slate-400 font-medium">
                                                            No configuration rows found
                                                        </td>
                                                    </tr>
                                                ) : (
                                                    getFilteredMasterList(selectedMasterTable).map((row, rIdx) => {
                                                        const pk = getTablePrimaryKey(selectedMasterTable);
                                                        const pkValStr = String(row[pk]);
                                                        return (
                                                            <tr key={rIdx} className={`hover:bg-slate-50/20 transition ${selectedMasterRowKeys.has(pkValStr) ? 'bg-sky-50/40' : ''}`}>
                                                                {currentUser.role === "IT_ADMIN" && isMasterDeleteMode && (
                                                                    <td className="py-3 px-4">
                                                                        <input
                                                                            type="checkbox"
                                                                            checked={selectedMasterRowKeys.has(pkValStr)}
                                                                            onChange={() => handleToggleMasterRowSelect(row[pk])}
                                                                            className="rounded text-sky-600 focus:ring-sky-500 h-4 w-4 cursor-pointer"
                                                                        />
                                                                    </td>
                                                                )}
                                                                {getMasterTableHeaders(selectedMasterTable).map((h, cIdx) => (
                                                                    <td key={cIdx} className="py-3 px-4 font-medium text-xs">
                                                                        {h === "status" ? (
                                                                            <span className={`px-2 py-0.5 rounded text-[9px] font-bold ${row[h] === 'Active' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' : 'bg-slate-50 text-slate-500 border border-slate-100'}`}>
                                                                                {row[h]}
                                                                            </span>
                                                                        ) : h === "role" ? (
                                                                            <span className={`px-2 py-0.5 rounded text-[9px] font-bold ${row[h] === 'IT_ADMIN' ? 'bg-indigo-50 text-indigo-700 border border-indigo-100' : 'bg-slate-100 text-slate-700'}`}>
                                                                                {row[h]}
                                                                            </span>
                                                                        ) : h === "type" && selectedMasterTable === "tariff_rates" ? (
                                                                            <span className="capitalize font-bold text-slate-700 dark:text-slate-300">
                                                                                {row[h]}
                                                                            </span>
                                                                        ) : (h === "allowed_locations" || h === "allowed_plants") && row[h] === "all" ? (
                                                                            <span className="px-2 py-0.5 rounded text-[9px] font-bold bg-sky-50 text-sky-700 border border-sky-100">
                                                                                ALL
                                                                            </span>
                                                                        ) : row[h] !== undefined && row[h] !== null ? String(row[h]) : ""}
                                                                    </td>
                                                                ))}
                                                                <td className="py-3 px-4 text-right space-x-2.5 whitespace-nowrap">
                                                                    <button onClick={() => openMasterForm(row)} className="px-2.5 py-1.5 text-[11px] font-extrabold text-sky-600 hover:text-sky-700 bg-sky-50 dark:bg-sky-950/20 hover:bg-sky-100 dark:hover:bg-sky-900/30 rounded-xl cursor-pointer border-none transition shadow-sm">
                                                                        Edit
                                                                    </button>
                                                                    <button onClick={() => handleDeleteMasterRecord(row[pk])} className="px-2.5 py-1.5 text-[11px] font-extrabold text-red-600 hover:text-red-700 bg-red-50 dark:bg-red-950/20 hover:bg-red-100 dark:hover:bg-red-900/30 rounded-xl cursor-pointer border-none transition shadow-sm">
                                                                        Delete
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        );
                                                    })
                                                )}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                                )}
                            </div>
                        )}

                        {/* FOOTER */}
                        <footer className="pt-8 pb-4 text-center text-[10px] text-slate-400 tracking-wider no-print border-t border-slate-100/50 mt-12 uppercase font-semibold">
                            UTILITY SENSE v3.0 · Secured Apps Script Data Hub · Corporate Governance Panel
                        </footer>
                    </main>

                    {/* ----------------------------------------------------
                        MODALS
                    ---------------------------------------------------- */}

                    {isFormOpen && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
                            <div className="relative w-full max-w-4xl overflow-hidden rounded-2xl bg-white border border-slate-200 shadow-xl transition-all flex flex-col max-h-[90vh]">
                                {/* Header */}
                                <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4 bg-white">
                                    <div className="flex items-center gap-3">
                                        <div className="h-9 w-9 rounded-lg bg-sky-50 border border-sky-100 flex items-center justify-center text-[#0284c7] flex-shrink-0">
                                            <span className="material-symbols-outlined text-[20px]">bolt</span>
                                        </div>
                                        <div>
                                            <h3 className="text-sm font-bold text-slate-800">
                                                {editingRecord ? "Modify Operational Entry" : "New Operational Entry"}
                                            </h3>
                                            <p className="text-[11px] text-slate-500 mt-0.5">Log daily readings, consumption, and production costs</p>
                                        </div>
                                    </div>
                                    <div className="flex items-center gap-2">
                                        <button type="button" onClick={() => setIsFormOpen(false)} className="h-8 w-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600 transition border-none bg-transparent cursor-pointer">
                                            <span className="material-symbols-outlined text-[18px]">close</span>
                                        </button>
                                    </div>
                                </div>

                                <form onSubmit={handleEntryFormSubmit} className="flex-1 overflow-y-auto px-6 py-5 flex flex-col md:flex-row gap-5 text-xs bg-slate-50/60">
                                    {/* Left Side: Inputs Form */}
                                    <div className="flex-1 space-y-4">

                                        {/* Row 1: Key Metadata */}
                                        <div className="bg-white rounded-xl border border-slate-200 p-4">
                                            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                                                <div>
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Date</label>
                                                    <input
                                                        type="date"
                                                        required
                                                        value={entryFormValues.date}
                                                        onChange={(e) => updateFormCalculations({ date: e.target.value })}
                                                        className={`w-full h-9 rounded-lg border px-2.5 bg-white text-slate-700 focus:outline-none focus:ring-2 transition ${isDuplicateDateEntry ? "border-red-400 focus:ring-red-500/30 text-red-700 bg-red-50/40" : "border-slate-300 focus:ring-sky-500/30 focus:border-sky-400"}`}
                                                    />
                                                    {isDuplicateDateEntry && (
                                                        <p className="mt-1 text-[10px] font-bold text-red-600 flex items-center gap-1 leading-tight animate-pulse">
                                                            <span className="material-symbols-outlined text-[13px] shrink-0">warning</span>
                                                            <span>An entry already exists for this date. Duplicate entries are not allowed.</span>
                                                        </p>
                                                    )}
                                                </div>
                                                <div>
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Location</label>
                                                    <select
                                                        value={entryFormValues.location}
                                                        onChange={(e) => {
                                                            const newLoc = e.target.value;
                                                            const filtered = allowedPlants.filter(p => p.location === newLoc);
                                                            const newPlant = filtered[0]?.plant_code || "";
                                                            updateFormCalculations({ location: newLoc, plant: newPlant });
                                                        }}
                                                        className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 font-medium focus:outline-none focus:ring-2 focus:ring-sky-500/30 focus:border-sky-400 transition"
                                                    >
                                                        {Array.from(new Set(allowedPlants.map(p => p.location).filter(Boolean))).map(loc => (
                                                            <option key={loc} value={loc}>{loc}</option>
                                                        ))}
                                                    </select>
                                                </div>
                                                <div>
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Plant</label>
                                                    <select
                                                        value={entryFormValues.plant}
                                                        onChange={(e) => updateFormCalculations({ plant: e.target.value })}
                                                        className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 font-medium focus:outline-none focus:ring-2 focus:ring-sky-500/30 focus:border-sky-400 transition"
                                                    >
                                                        {allowedPlants.filter(p => p.location === entryFormValues.location).map(p => (
                                                            <option key={p.plant_code} value={p.plant_code}>{p.plant_name}</option>
                                                        ))}
                                                    </select>
                                                </div>
                                                <div>
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Operator</label>
                                                    <input type="text" required placeholder="Name" value={entryFormValues.operator_name} onChange={(e) => updateFormCalculations({ operator_name: e.target.value })} className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-sky-500/30 focus:border-sky-400 transition" />
                                                </div>
                                            </div>
                                        </div>

                                        {/* Section 2: Electricity Grid Logging */}
                                        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                                            <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100 bg-sky-50/50">
                                                <div className="flex items-center gap-2">
                                                    <span className="material-symbols-outlined text-[16px] text-[#0284c7]">electric_bolt</span>
                                                    <span className="text-[11px] font-semibold text-slate-700">Electricity Grid Meter (kWh)</span>
                                                </div>
                                                <label className="flex items-center gap-1.5 cursor-pointer select-none text-[10px] font-medium text-slate-500 hover:text-slate-700 transition">
                                                    <input
                                                        type="checkbox"
                                                        checked={entryFormValues.meter_changed}
                                                        onChange={(e) => updateFormCalculations({ meter_changed: e.target.checked })}
                                                        className="rounded border-slate-300 text-[#0284c7] focus:ring-[#0284c7] h-3.5 w-3.5"
                                                    />
                                                    <span>Meter Replaced / Changed</span>
                                                </label>
                                            </div>

                                            <div className="p-4 grid grid-cols-1 sm:grid-cols-3 gap-3">
                                                <div>
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Previous Reading (kWh)</label>
                                                    <input
                                                        type="number"
                                                        value={entryFormValues.electricity_opening}
                                                        onChange={(e) => updateFormCalculations({ electricity_opening: Number(e.target.value) || 0 })}
                                                        className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-slate-50 text-slate-500 focus:outline-none focus:ring-2 focus:ring-sky-500/30 focus:border-sky-400 transition"
                                                    />
                                                </div>
                                                <div>
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">
                                                        Daily Reading (kWh)
                                                        {isFetchingExcelData && <span className="ml-1 text-[8px] text-sky-600 animate-pulse">(Syncing...)</span>}
                                                    </label>
                                                    <input
                                                        type="number"
                                                        required
                                                        step="any"
                                                        min="0"
                                                        placeholder={isFetchingExcelData ? "Syncing..." : "Enter meter reading"}
                                                        disabled={isFetchingExcelData}
                                                        value={isFetchingExcelData ? "" : entryFormValues.electricity_closing}
                                                        onChange={(e) => updateFormCalculations({ electricity_closing: e.target.value })}
                                                        className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-sky-500/30 focus:border-sky-400 transition disabled:bg-slate-50 disabled:text-slate-400"
                                                    />
                                                </div>
                                                {entryFormValues.meter_changed && (
                                                    <div>
                                                        <label className="block text-[10px] font-semibold text-red-500 mb-1.5">Custom Net Difference</label>
                                                        <input
                                                            type="number"
                                                            required
                                                            step="any"
                                                            min="0"
                                                            placeholder="Reset difference value"
                                                            value={entryFormValues.custom_difference}
                                                            onChange={(e) => updateFormCalculations({ custom_difference: e.target.value })}
                                                            className="w-full h-9 rounded-lg border border-red-300 px-2.5 bg-red-50/40 text-red-800 font-medium focus:outline-none focus:ring-2 focus:ring-red-400/30 focus:border-red-400 transition"
                                                        />
                                                    </div>
                                                )}
                                            </div>
                                        </div>

                                        {/* Section 3: Solar — Previous show-only; Daily × ₹10.89 */}
                                        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                                            <div className="flex items-center gap-2 px-4 py-3 border-b border-slate-100 bg-amber-50/50">
                                                <span className="material-symbols-outlined text-[16px] text-amber-600">solar_power</span>
                                                <span className="text-[11px] font-semibold text-slate-700">Solar Generated Units (kWh)</span>
                                                {isFetchingExcelData && (
                                                    <span className="ml-auto text-[9px] font-bold text-amber-600 animate-pulse uppercase tracking-wider flex items-center gap-1">
                                                        <span className="material-symbols-outlined text-[12px] animate-spin">sync</span>
                                                        Fetching...
                                                    </span>
                                                )}
                                            </div>
                                            <div className="p-4 grid grid-cols-1 sm:grid-cols-2 gap-3">
                                                <div>
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Previous Reading (kWh)</label>
                                                    <input
                                                        type="number"
                                                        readOnly
                                                        disabled
                                                        value={entryFormValues.solar_opening || 0}
                                                        className="w-full h-9 rounded-lg border border-slate-200 px-2.5 bg-slate-100 text-slate-500 cursor-not-allowed select-none"
                                                    />
                                                </div>
                                                <div>
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">
                                                        Daily Reading (kWh)
                                                        <span className="ml-1 font-normal text-slate-400">× ₹{activeSolarRate}/unit</span>
                                                        {isFetchingExcelData && <span className="ml-1 text-[8px] text-amber-600 animate-pulse">(Syncing...)</span>}
                                                    </label>
                                                    <input
                                                        type="number"
                                                        required
                                                        step="any"
                                                        min="0"
                                                        placeholder={isFetchingExcelData ? "Syncing..." : "Enter solar units"}
                                                        disabled={isFetchingExcelData}
                                                        value={isFetchingExcelData ? "" : entryFormValues.solar_closing}
                                                        onChange={(e) => updateFormCalculations({ solar_closing: e.target.value })}
                                                        className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-amber-500/30 focus:border-amber-400 transition disabled:bg-slate-50 disabled:text-slate-400"
                                                    />
                                                </div>
                                            </div>
                                        </div>

                                        {/* Section 4: Diesel below Solar */}
                                        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                                            <div className="flex items-center gap-2 px-4 py-3 border-b border-slate-100 bg-red-50/50">
                                                <span className="material-symbols-outlined text-[16px] text-red-500">local_gas_station</span>
                                                <span className="text-[11px] font-semibold text-slate-700">Diesel Consumption</span>
                                                {isFetchingExcelData && (
                                                    <span className="ml-auto text-[9px] font-bold text-red-500 animate-pulse uppercase tracking-wider flex items-center gap-1">
                                                        <span className="material-symbols-outlined text-[12px] animate-spin">sync</span>
                                                        Fetching...
                                                    </span>
                                                )}
                                            </div>
                                            <div className="p-4">
                                                <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Diesel Qty (Liters)</label>
                                                <input
                                                    type="number"
                                                    required
                                                    step="any"
                                                    min="0"
                                                    placeholder={isFetchingExcelData ? "Syncing..." : "Enter diesel liters"}
                                                    disabled={isFetchingExcelData}
                                                    value={isFetchingExcelData ? "" : entryFormValues.diesel_used}
                                                    onChange={(e) => updateFormCalculations({ diesel_used: e.target.value })}
                                                    className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-red-500/30 focus:border-red-400 transition disabled:bg-slate-50 disabled:text-slate-400"
                                                />
                                            </div>
                                        </div>

                                        {/* Section 5: Production Input (Auto-calculated from OneDrive production_summary) */}
                                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                            {/* ODU */}
                                            <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                                                <div className="flex items-center gap-2 px-4 py-3 border-b border-slate-100 bg-orange-50/50">
                                                    <span className="material-symbols-outlined text-[16px] text-orange-600">precision_manufacturing</span>
                                                    <span className="text-[11px] font-semibold text-slate-700">ODU Count</span>
                                                    <span className="ml-auto text-[9px] font-extrabold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full flex items-center gap-1">
                                                        <span className="material-symbols-outlined text-[11px]">sync</span>
                                                        Auto-Sync
                                                    </span>
                                                </div>
                                                <div className="p-4">
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">ODU (Auto-calculated)</label>
                                                    <input
                                                        type="number"
                                                        readOnly={true}
                                                        disabled={true}
                                                        value={entryFormValues.odu || 0}
                                                        className="w-full h-9 rounded-lg border border-slate-200 px-2.5 bg-slate-100 text-slate-700 font-bold cursor-not-allowed select-none"
                                                    />
                                                </div>
                                            </div>

                                            {/* IDU */}
                                            <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                                                <div className="flex items-center gap-2 px-4 py-3 border-b border-slate-100 bg-orange-50/50">
                                                    <span className="material-symbols-outlined text-[16px] text-orange-600">precision_manufacturing</span>
                                                    <span className="text-[11px] font-semibold text-slate-700">IDU Count</span>
                                                    <span className="ml-auto text-[9px] font-extrabold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full flex items-center gap-1">
                                                        <span className="material-symbols-outlined text-[11px]">sync</span>
                                                        Auto-Sync
                                                    </span>
                                                </div>
                                                <div className="p-4">
                                                    <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">IDU (Auto-calculated)</label>
                                                    <input
                                                        type="number"
                                                        readOnly={true}
                                                        disabled={true}
                                                        value={entryFormValues.idu || 0}
                                                        className="w-full h-9 rounded-lg border border-slate-200 px-2.5 bg-slate-100 text-slate-700 font-bold cursor-not-allowed select-none"
                                                    />
                                                </div>
                                            </div>
                                        </div>

                                        {/* Section 5: Waste Management (collapsible, closed by default) */}
                                        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                                            <button
                                                type="button"
                                                onClick={() => setWasteSectionOpen(o => !o)}
                                                className="w-full flex items-center justify-between gap-2 px-4 py-3 bg-slate-50/60 hover:bg-slate-100/60 transition cursor-pointer border-none"
                                            >
                                                <div className="flex items-center gap-2">
                                                    <span className="material-symbols-outlined text-[16px] text-slate-500">recycling</span>
                                                    <span className="text-[11px] font-semibold text-slate-700">Waste Management (Optional)</span>
                                                    {(Number(entryFormValues.waste_hazardous) > 0 || Number(entryFormValues.waste_non_hazardous) > 0 || Number(entryFormValues.waste_recycled) > 0) && (
                                                        <span className="text-[9px] font-bold uppercase tracking-wide px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-600 border border-emerald-200/60">Filled</span>
                                                    )}
                                                </div>
                                                <span className={`material-symbols-outlined text-[18px] text-slate-400 transition-transform duration-200 ${wasteSectionOpen ? "rotate-180" : ""}`}>expand_more</span>
                                            </button>
                                            {wasteSectionOpen && (
                                                <div className="p-4 border-t border-slate-100 grid grid-cols-1 sm:grid-cols-3 gap-4">
                                                    <div>
                                                        <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Hazardous Waste (kg)</label>
                                                        <input
                                                            type="number"
                                                            step="any"
                                                            min="0"
                                                            placeholder="0"
                                                            value={entryFormValues.waste_hazardous}
                                                            onChange={(e) => updateFormCalculations({ waste_hazardous: e.target.value })}
                                                            className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-red-500/30 focus:border-red-400 transition"
                                                        />
                                                    </div>
                                                    <div>
                                                        <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Non-Hazardous Waste (kg)</label>
                                                        <input
                                                            type="number"
                                                            step="any"
                                                            min="0"
                                                            placeholder="0"
                                                            value={entryFormValues.waste_non_hazardous}
                                                            onChange={(e) => updateFormCalculations({ waste_non_hazardous: e.target.value })}
                                                            className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-slate-500/30 focus:border-slate-400 transition"
                                                        />
                                                    </div>
                                                    <div>
                                                        <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Recycled Waste (kg)</label>
                                                        <input
                                                            type="number"
                                                            step="any"
                                                            min="0"
                                                            placeholder="0"
                                                            value={entryFormValues.waste_recycled}
                                                            onChange={(e) => updateFormCalculations({ waste_recycled: e.target.value })}
                                                            className="w-full h-9 rounded-lg border border-slate-300 px-2.5 bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-emerald-500/30 focus:border-emerald-400 transition"
                                                        />
                                                    </div>
                                                </div>
                                            )}
                                        </div>

                                        {/* Remarks Section (kept last) */}
                                        <div className="bg-white rounded-xl border border-slate-200 p-4">
                                            <label className="block text-[10px] font-semibold text-slate-500 mb-1.5">Remarks & Observations</label>
                                            <textarea rows="2" placeholder="Optional comments..." value={entryFormValues.remarks} onChange={(e) => updateFormCalculations({ remarks: e.target.value })} className="w-full border border-slate-300 p-2.5 rounded-lg bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-sky-500/30 focus:border-sky-400 resize-none transition" />
                                        </div>

                                    </div>

                                    {/* Right Side: Summary Panel */}
                                    <div className="w-full md:w-80 flex flex-col gap-3">

                                        {/* Cost / Audit summary card */}
                                        <div className="flex-1 flex flex-col bg-white rounded-xl border border-slate-200 p-4">
                                            <div className="flex items-center gap-2 border-b border-slate-100 pb-2.5 mb-3">
                                                <span className="material-symbols-outlined text-[15px] text-slate-400">receipt_long</span>
                                                <span className="text-[10px] font-semibold text-slate-500 uppercase tracking-wide">Cost Summary</span>
                                            </div>

                                            <div className="flex-1 space-y-3.5 text-[11px]">

                                                {/* Grid difference and MSEB units */}
                                                <div className="space-y-1.5 pb-3 border-b border-slate-100">
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-slate-500">Meter Difference</span>
                                                        <span className="text-slate-800 font-semibold">
                                                            {entryFormValues.meter_changed
                                                                ? `${Number(entryFormValues.custom_difference) || 0} (Manual)`
                                                                : `${Math.max(0, (Number(entryFormValues.electricity_closing) || 0) - (Number(entryFormValues.electricity_opening) || 0))}`}
                                                        </span>
                                                    </div>
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-slate-500">{entryRateContext.gridLabel} Units (×{entryRateContext.mf})</span>
                                                        <span className="text-[#0284c7] font-semibold">{fmtNum(entryFormValues.electricity_consumption)} kWh</span>
                                                    </div>
                                                </div>

                                                {/* Cost breakdowns */}
                                                <div className="space-y-2 pb-3 border-b border-slate-100">
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-slate-500">Electricity Cost</span>
                                                        <span className="text-slate-800 font-medium">₹ {fmtNum(entryFormValues.electricity_cost)}</span>
                                                    </div>
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-slate-500">Solar Cost</span>
                                                        <span className="text-slate-800 font-medium">₹ {fmtNum(entryFormValues.solar_cost)}</span>
                                                    </div>
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-slate-500">Diesel Cost</span>
                                                        <span className="text-slate-800 font-medium">₹ {fmtNum(entryFormValues.diesel_cost)}</span>
                                                    </div>
                                                    <p className="text-[9px] text-slate-400 pt-1">Rate: Grid: ₹{entryRateContext.electRate}/unit · Solar: ₹{entryRateContext.solarRate}/unit · Diesel: ₹{entryRateContext.dieselRate}/liter</p>
                                                </div>

                                                {/* Total cost */}
                                                <div className="bg-sky-50 border border-sky-100 rounded-lg p-3 flex justify-between items-center">
                                                    <span className="text-[10px] font-semibold text-slate-600 uppercase tracking-wide">Total Consumption</span>
                                                    <span className="text-[#0369a1] font-bold text-sm">₹ {fmtNum(entryFormValues.total_cost)}</span>
                                                </div>

                                                {/* Production Set formula */}
                                                <div className="space-y-1.5 pb-3 border-b border-slate-100">
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-slate-500">ODU Count</span>
                                                        <span className="text-slate-800 font-medium">{entryFormValues.odu || 0}</span>
                                                    </div>
                                                    <div className="flex justify-between items-center">
                                                        <span className="text-slate-500">IDU Count</span>
                                                        <span className="text-slate-800 font-medium">{entryFormValues.idu || 0}</span>
                                                    </div>
                                                    <div className="flex justify-between items-center pt-1">
                                                        <span className="text-slate-700 font-semibold">Production Sets</span>
                                                        <span className="text-orange-600 font-bold">{entryFormValues.production_set} Sets</span>
                                                    </div>
                                                </div>

                                                {/* Cost per set */}
                                                <div className="bg-emerald-50 border border-emerald-100 rounded-lg p-3 text-center">
                                                    <p className="text-[9px] text-emerald-700 font-semibold uppercase tracking-wide">Cost Per Production Set</p>
                                                    <p className="text-lg font-bold text-emerald-600 mt-1">₹ {fmtNum(entryFormValues.cost_per_set)} <span className="text-[10px] font-normal text-slate-400">/ set</span></p>
                                                </div>

                                            </div>
                                        </div>

                                        {/* Actions footer */}
                                        <div className="flex items-center justify-end gap-2">
                                            <button
                                                type="button"
                                                onClick={() => setIsFormOpen(false)}
                                                className="px-4 py-2 bg-white border border-slate-300 hover:bg-slate-50 text-slate-600 rounded-lg text-xs font-semibold transition cursor-pointer"
                                            >
                                                Cancel
                                            </button>
                                            <button
                                                type="submit"
                                                disabled={actionLoading || isDuplicateDateEntry}
                                                className="flex items-center gap-1.5 px-4 py-2 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-lg text-xs font-semibold transition shadow-sm cursor-pointer disabled:opacity-50 border-none"
                                            >
                                                {actionLoading && <span className="w-3 h-3 border-2 border-white border-t-transparent rounded-full animate-spin" />}
                                                <span>Save Entry</span>
                                            </button>
                                        </div>

                                    </div>
                                </form>
                            </div>
                        </div>
                    )}

                    {/* View Entry Details Popup (read-only, opened by clicking a row) */}
                    {viewingRecord && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
                            <div className="relative w-full max-w-4xl overflow-hidden rounded-2xl bg-white border border-slate-200 shadow-xl transition-all flex flex-col max-h-[90vh]">
                                {/* Header */}
                                <div className="flex items-center justify-between border-b border-slate-200 px-6 py-4 bg-white">
                                    <div className="flex items-center gap-3">
                                        <div className="h-9 w-9 rounded-lg bg-sky-50 border border-sky-100 flex items-center justify-center text-[#0284c7] flex-shrink-0">
                                            <span className="material-symbols-outlined text-[20px]">visibility</span>
                                        </div>
                                        <div>
                                            <h3 className="text-sm font-bold text-slate-800">
                                                Entry Details — {viewingRecord.date} · {viewingRecord.plant} {viewingRecord.location ? `(${viewingRecord.location})` : ""}
                                            </h3>
                                            <p className="text-[11px] text-slate-500 mt-0.5">Full read-only snapshot of everything recorded for this operational log</p>
                                        </div>
                                    </div>
                                    <button onClick={() => setViewingRecord(null)} className="h-8 w-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600 transition border-none bg-transparent cursor-pointer">
                                        <span className="material-symbols-outlined text-[18px]">close</span>
                                    </button>
                                </div>

                                <div className="flex-1 overflow-y-auto px-6 py-5 space-y-4 text-xs bg-slate-50/60">

                                    {/* Meta */}
                                    <div className="bg-white rounded-xl border border-slate-200 p-4">
                                        <div className="flex items-center gap-2 mb-3 text-slate-600">
                                            <span className="material-symbols-outlined text-[16px]">event_note</span>
                                            <span className="text-[11px] font-semibold uppercase">Log Metadata</span>
                                        </div>
                                        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                                            <DetailField label="Date" value={viewingRecord.date} />
                                            <DetailField label="Location" value={viewingRecord.location} />
                                            <DetailField label="Plant" value={viewingRecord.plant} />
                                            <DetailField label="Department" value={viewingRecord.department} />
                                            <DetailField label="Shift" value={viewingRecord.shift} />
                                            <DetailField label="Operator" value={viewingRecord.operator_name} />
                                            <DetailField label="Created At" value={viewingRecord.created_at} />
                                            <DetailField label="Created By" value={viewingRecord.created_by} />
                                        </div>
                                    </div>

                                    {/* Cost Summary & Production Efficiency (simplified: only the figures that matter for review) */}
                                    <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                                        <div className="flex items-center gap-2 px-4 py-3 border-b border-slate-100 bg-slate-50">
                                            <span className="material-symbols-outlined text-[16px] text-slate-500">receipt_long</span>
                                            <span className="text-[11px] font-semibold text-slate-700">Cost Summary & Production Efficiency</span>
                                        </div>
                                        <div className="p-4 grid grid-cols-2 sm:grid-cols-4 gap-4">
                                            <DetailField label="Meter Difference" value={fmtNum(Math.max(0, (Number(viewingRecord.electricity_closing) || 0) - (Number(viewingRecord.electricity_opening) || 0)))} />
                                            <DetailField label={`${getGridProviderLabel(viewingRecord.location || plants.find(p => p.plant_code === viewingRecord.plant)?.location)} Units`} value={fmtNum(viewingRecord.electricity_consumption) + " kWh"} highlight />
                                            <DetailField label="Electricity Cost" value={"₹ " + fmtNum(viewingRecord.electricity_cost)} />
                                            <DetailField label="Solar Cost" value={"₹ " + fmtNum(viewingRecord.solar_cost)} />
                                            <DetailField label="Diesel Used (L)" value={fmtNum(viewingRecord.diesel_used)} />
                                            <DetailField label="Diesel Cost" value={"₹ " + fmtNum(viewingRecord.diesel_cost)} />
                                            <DetailField label="ODU Count" value={fmtNum(Number(viewingRecord.odu) || 0)} />
                                            <DetailField label="IDU Count" value={fmtNum(Number(viewingRecord.idu) || 0)} />
                                            <DetailField label="Production Sets" value={fmtNum(viewingRecord.production_set) + " Sets"} highlight />
                                            <DetailField label="Cost Per Set" value={viewingRecord.cost_per_set !== undefined && viewingRecord.cost_per_set !== "" ? "₹ " + fmtNum(viewingRecord.cost_per_set) : "—"} highlight />
                                            <DetailField label="Hazardous Waste (kg)" value={fmtNum(Number(viewingRecord.waste_hazardous) || 0)} />
                                            <DetailField label="Non-Hazardous Waste (kg)" value={fmtNum(Number(viewingRecord.waste_non_hazardous) || 0)} />
                                            <DetailField label="Recycled Waste (kg)" value={fmtNum(Number(viewingRecord.waste_recycled) || 0)} />
                                        </div>
                                        <div className="px-4 pb-4">
                                            <div className="bg-sky-50 border border-sky-100 rounded-lg p-3 flex justify-between items-center">
                                                <span className="text-[10px] font-semibold text-slate-600 uppercase tracking-wide">Total Consumption</span>
                                                <span className="text-[#0369a1] font-bold text-sm">₹ {fmtNum(viewingRecord.total_cost)}</span>
                                            </div>
                                        </div>
                                    </div>

                                    {/* Remarks */}
                                    <div className="bg-white rounded-xl border border-slate-200 p-4">
                                        <p className="text-[9px] font-semibold text-slate-400 uppercase tracking-wide mb-1">Remarks & Observations</p>
                                        <p className="text-xs text-slate-700">{viewingRecord.remarks || "—"}</p>
                                    </div>

                                </div>

                                {/* Footer Actions */}
                                <div className="flex items-center justify-end gap-2 border-t border-slate-200 px-6 py-4 bg-white">
                                    <button
                                        onClick={() => setViewingRecord(null)}
                                        className="px-4 py-2 bg-white border border-slate-300 hover:bg-slate-50 text-slate-600 rounded-lg text-xs font-semibold transition cursor-pointer"
                                    >
                                        Close
                                    </button>
                                    <button
                                        onClick={() => { const rec = viewingRecord; setViewingRecord(null); openDailyForm(rec); }}
                                        className="flex items-center gap-1.5 px-4 py-2 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-lg text-xs font-semibold transition shadow-sm cursor-pointer border-none"
                                    >
                                        <span className="material-symbols-outlined text-[16px]">edit</span>
                                        <span>Edit This Entry</span>
                                    </button>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* 2. Unified Master Record Form Modal */}
                    {isMasterFormOpen && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
                            <div className="w-full max-w-2xl rounded-3xl bg-white border border-slate-200 shadow-2xl transition-all scale-100 flex flex-col max-h-[90vh]">
                                <div className="flex items-center justify-between px-6 py-4 border-b">
                                    <h3 className="text-sm font-bold text-slate-900 uppercase">
                                        {editingMasterRecord ? `Edit ${selectedMasterTable} Record` : `Register New ${selectedMasterTable}`}
                                    </h3>
                                    <button onClick={() => setIsMasterFormOpen(false)} className="text-slate-400 hover:text-slate-600 bg-transparent border-none cursor-pointer">
                                        <span className="material-symbols-outlined text-[18px]">close</span>
                                    </button>
                                </div>

                                <form onSubmit={handleMasterFormSubmit} className="flex-1 overflow-y-auto px-6 py-5 flex flex-col justify-between text-xs space-y-5">
                                    {selectedMasterTable === "users" ? (
                                        <div className="space-y-4">
                                            {/* Employee ID & Name */}
                                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                <div>
                                                    <label className="block text-[9px] font-bold text-slate-500 uppercase tracking-wider mb-1">Employee ID *</label>
                                                    <input
                                                        type="text"
                                                        required
                                                        disabled={editingMasterRecord !== null}
                                                        placeholder="Enter employee ID"
                                                        value={masterFormValues.id || ""}
                                                        onChange={(e) => setMasterFormValues({ ...masterFormValues, id: e.target.value.toUpperCase() })}
                                                        className="w-full h-10 rounded-xl border border-slate-200 px-3 text-xs bg-white focus:outline-none disabled:opacity-50 uppercase"
                                                    />
                                                </div>
                                                <div>
                                                    <label className="block text-[9px] font-bold text-slate-500 uppercase tracking-wider mb-1">Name *</label>
                                                    <input
                                                        type="text"
                                                        required
                                                        placeholder="Enter name"
                                                        value={masterFormValues.name || ""}
                                                        onChange={(e) => setMasterFormValues({ ...masterFormValues, name: e.target.value.toUpperCase() })}
                                                        className="w-full h-10 rounded-xl border border-slate-200 px-3 text-xs bg-white focus:outline-none uppercase"
                                                    />
                                                </div>
                                            </div>

                                            {/* Company Email & Role */}
                                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                                <div>
                                                    <label className="block text-[9px] font-bold text-slate-500 uppercase tracking-wider mb-1">Company Email *</label>
                                                    <input
                                                        type="email"
                                                        required
                                                        placeholder="Enter company email"
                                                        value={masterFormValues.email || ""}
                                                        onChange={(e) => setMasterFormValues({ ...masterFormValues, email: e.target.value })}
                                                        className="w-full h-10 rounded-xl border border-slate-200 px-3 text-xs bg-white focus:outline-none"
                                                    />
                                                </div>
                                                <div>
                                                    <label className="block text-[9px] font-bold text-slate-500 uppercase tracking-wider mb-1">Role *</label>
                                                    {editingMasterRecord && masterFormValues.role === "IT_ADMIN" ? (
                                                        <input
                                                            type="text"
                                                            value="IT_ADMIN"
                                                            disabled
                                                            readOnly
                                                            className="w-full h-10 rounded-xl border border-slate-200 px-3 text-xs bg-slate-100 text-slate-500 font-bold cursor-not-allowed"
                                                        />
                                                    ) : (
                                                        <select
                                                            value={masterFormValues.role || "USER"}
                                                            onChange={(e) => setMasterFormValues({ ...masterFormValues, role: e.target.value })}
                                                            className="w-full h-10 rounded-xl border border-slate-200 px-3 text-xs bg-white focus:outline-none font-bold"
                                                        >
                                                            <option value="USER">USER</option>
                                                            <option value="IT_ADMIN">IT_ADMIN</option>
                                                        </select>
                                                    )}
                                                </div>
                                            </div>

                                            {/* Locations Checkboxes */}
                                            <div>
                                                <label className="block text-[9px] font-bold text-slate-500 uppercase tracking-wider mb-1.5">Locations (select one or more)</label>
                                                <div className="flex flex-wrap gap-2 p-3 bg-slate-50 rounded-2xl border border-slate-200/50">
                                                    {Array.from(new Set(plants.map(p => p.location).filter(Boolean))).map(loc => {
                                                        const allowedLocsStr = masterFormValues.allowed_locations || "";
                                                        const selectedLocs = allowedLocsStr === "all"
                                                            ? Array.from(new Set(plants.map(p => p.location).filter(Boolean)))
                                                            : (allowedLocsStr ? allowedLocsStr.split(",").map(x => x.trim()) : []);
                                                        const isChecked = selectedLocs.includes(loc);

                                                        const handleLocToggle = () => {
                                                            let nextLocs = [...selectedLocs];
                                                            if (nextLocs.includes(loc)) {
                                                                nextLocs = nextLocs.filter(x => x !== loc);
                                                            } else {
                                                                nextLocs.push(loc);
                                                            }
                                                            const allLocs = Array.from(new Set(plants.map(p => p.location).filter(Boolean)));
                                                            const allowedL = nextLocs.length === allLocs.length ? "all" : nextLocs.join(",");

                                                            // Auto filter plants to match selected locations
                                                            const nextPlants = plants.filter(p => nextLocs.includes(p.location)).map(p => p.plant_code);
                                                            const allowedP = nextPlants.length === plants.length ? "all" : nextPlants.join(",");

                                                            setMasterFormValues(prev => ({
                                                                ...prev,
                                                                allowed_locations: allowedL,
                                                                allowed_plants: allowedP
                                                            }));
                                                        };

                                                        return (
                                                            <label key={loc} className={`flex items-center gap-2 px-3 py-1.5 border rounded-xl cursor-pointer transition select-none text-xs font-bold ${isChecked ? 'bg-sky-50 border-sky-300 text-sky-800' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}>
                                                                <input
                                                                    type="checkbox"
                                                                    checked={isChecked}
                                                                    onChange={handleLocToggle}
                                                                    className="rounded text-sky-600 focus:ring-sky-500 h-4 w-4"
                                                                />
                                                                <span>{loc}</span>
                                                            </label>
                                                        );
                                                    })}
                                                </div>
                                            </div>

                                            {/* Plants Checkboxes */}
                                            <div>
                                                <label className="block text-[9px] font-bold text-slate-500 uppercase tracking-wider mb-1.5">Plants (select multiple)</label>
                                                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 p-3 bg-slate-50 rounded-2xl border border-slate-200/50 max-h-48 overflow-y-auto">
                                                    {(() => {
                                                        const allowedLocsStr = masterFormValues.allowed_locations || "";
                                                        const selectedLocs = allowedLocsStr === "all"
                                                            ? Array.from(new Set(plants.map(p => p.location).filter(Boolean)))
                                                            : (allowedLocsStr ? allowedLocsStr.split(",").map(x => x.trim()) : []);

                                                        const visiblePlants = plants.filter(p => selectedLocs.includes(p.location));

                                                        const allowedPlantsStr = masterFormValues.allowed_plants || "";
                                                        const selectedPlants = allowedPlantsStr === "all"
                                                            ? plants.map(p => p.plant_code)
                                                            : (allowedPlantsStr ? allowedPlantsStr.split(",").map(x => x.trim()) : []);

                                                        return visiblePlants.length > 0 ? visiblePlants.map(p => {
                                                            const isChecked = selectedPlants.includes(p.plant_code);

                                                            const handlePlantToggle = () => {
                                                                let nextPlants = [...selectedPlants];
                                                                if (nextPlants.includes(p.plant_code)) {
                                                                    nextPlants = nextPlants.filter(x => x !== p.plant_code);
                                                                } else {
                                                                    nextPlants.push(p.plant_code);
                                                                }
                                                                const allowedP = nextPlants.length === plants.length ? "all" : nextPlants.join(",");
                                                                setMasterFormValues(prev => ({
                                                                    ...prev,
                                                                    allowed_plants: allowedP
                                                                }));
                                                            };

                                                            return (
                                                                <label key={p.plant_code} className={`flex items-center gap-2.5 p-2.5 border rounded-xl cursor-pointer transition select-none ${isChecked ? 'bg-slate-100 border-slate-300 text-slate-900 font-bold' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}>
                                                                    <input
                                                                        type="checkbox"
                                                                        checked={isChecked}
                                                                        onChange={handlePlantToggle}
                                                                        className="rounded text-sky-600 focus:ring-sky-500 h-4 w-4"
                                                                    />
                                                                    <div>
                                                                        <p className="text-[10px] font-extrabold text-slate-950">{p.plant_code}</p>
                                                                        <p className="text-[9px] text-slate-400 font-medium mt-0.5">{p.plant_display_name}</p>
                                                                    </div>
                                                                </label>
                                                            );
                                                        }) : (
                                                            <div className="col-span-2 text-center text-slate-400 py-4 font-bold">Please select at least one Location above.</div>
                                                        );
                                                    })()}
                                                </div>
                                            </div>
                                        </div>
                                    ) : (
                                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                            {/* Generic Form fields loop */}
                                            {getMasterTableHeaders(selectedMasterTable).map((h, idx) => {
                                                const isPk = h === getTablePrimaryKey(selectedMasterTable);
                                                return (
                                                    <div key={idx} className={getMasterTableHeaders(selectedMasterTable).length <= 4 ? "sm:col-span-2" : ""}>
                                                        <label className="block text-[11px] font-extrabold text-slate-600 dark:text-slate-400 uppercase tracking-wider mb-1.5">
                                                            {h.replace("_", " ")} {isPk ? "(Primary Key) *" : ""}
                                                        </label>

                                                        {h === "status" ? (
                                                            <select
                                                                value={masterFormValues[h]}
                                                                onChange={(e) => setMasterFormValues({ ...masterFormValues, [h]: e.target.value })}
                                                                className="w-full h-10 rounded-xl border border-slate-200 px-3 text-[12.5px] font-semibold bg-white focus:outline-none"
                                                            >
                                                                <option value="Active">Active</option>
                                                                <option value="Disabled">Disabled</option>
                                                            </select>
                                                        ) : h === "role" ? (
                                                            <select
                                                                value={masterFormValues[h]}
                                                                onChange={(e) => setMasterFormValues({ ...masterFormValues, [h]: e.target.value })}
                                                                className="w-full h-10 rounded-xl border border-slate-200 px-3 text-[12.5px] font-semibold bg-white focus:outline-none"
                                                            >
                                                                <option value="USER">USER</option>
                                                                <option value="IT_ADMIN">IT_ADMIN</option>
                                                            </select>
                                                        ) : h === "type" && selectedMasterTable === "tariff_rates" ? (
                                                            <select
                                                                value={masterFormValues[h]}
                                                                onChange={(e) => setMasterFormValues({ ...masterFormValues, [h]: e.target.value })}
                                                                className="w-full h-10 rounded-xl border border-slate-200 px-3 text-[12.5px] bg-white focus:outline-none font-bold"
                                                            >
                                                                <option value="electricity">Electricity</option>
                                                                <option value="solar">Solar</option>
                                                                <option value="water">Water</option>
                                                                <option value="diesel">Diesel</option>
                                                                <option value="lpg">LPG</option>
                                                            </select>
                                                        ) : h === "location" && (selectedMasterTable === "tariff_rates" || selectedMasterTable === "multiply_factors") ? (
                                                            <select
                                                                value={masterFormValues[h] || ""}
                                                                onChange={(e) => {
                                                                    const newLoc = e.target.value;
                                                                    let newPlantCode = masterFormValues.plant_code || "";
                                                                    if (newPlantCode) {
                                                                        const matchPlant = plants.find(p => p.plant_code === newPlantCode);
                                                                        if (matchPlant && matchPlant.location.toUpperCase() !== newLoc.toUpperCase()) {
                                                                            newPlantCode = "";
                                                                        }
                                                                    }
                                                                    setMasterFormValues({ ...masterFormValues, location: newLoc, plant_code: newPlantCode });
                                                                }}
                                                                className="w-full h-10 rounded-xl border border-slate-200 px-3 text-[12.5px] font-semibold bg-white focus:outline-none"
                                                            >
                                                                <option value="">Select Location</option>
                                                                {allowedLocations.map(loc => (
                                                                    <option key={loc} value={loc}>{loc}</option>
                                                                ))}
                                                            </select>
                                                        ) : h === "plant_code" && (selectedMasterTable === "tariff_rates" || selectedMasterTable === "multiply_factors") ? (
                                                            <select
                                                                value={masterFormValues[h] || ""}
                                                                onChange={(e) => setMasterFormValues({ ...masterFormValues, [h]: e.target.value || null })}
                                                                className="w-full h-10 rounded-xl border border-slate-200 px-3 text-[12.5px] font-semibold bg-white focus:outline-none"
                                                            >
                                                                <option value="">Global Location Fallback (No Plant)</option>
                                                                {(currentUser.role === "IT_ADMIN" ? plants : allowedPlants)
                                                                    .filter(p => !masterFormValues.location || p.location.toUpperCase() === masterFormValues.location.toUpperCase())
                                                                    .map(p => (
                                                                        <option key={p.plant_code} value={p.plant_code}>{p.plant_code} - {p.plant_name}</option>
                                                                    ))
                                                                }
                                                            </select>
                                                        ) : (
                                                            <React.Fragment>
                                                                <input
                                                                    type={h === "rate" || h === "year" || h === "month" || h.includes("target") || h === "factor" ? "number" : h.includes("date") ? "date" : "text"}
                                                                    required={!isPk || !["tariff_rates", "multiply_factors"].includes(selectedMasterTable)}
                                                                    placeholder={isPk && ["tariff_rates", "multiply_factors"].includes(selectedMasterTable) ? "Auto-generated (leave blank)" : `Enter ${h}`}
                                                                    value={masterFormValues[h]}
                                                                    onChange={(e) => setMasterFormValues({ ...masterFormValues, [h]: e.target.value })}
                                                                    className="w-full h-10 rounded-xl border border-slate-200 px-3 text-[12.5px] font-semibold bg-white focus:outline-none"
                                                                />
                                                                {isPk && editingMasterRecord !== null && (
                                                                    <p className="text-[10px] text-amber-600 dark:text-amber-400 font-extrabold mt-1.5 flex items-center gap-1">
                                                                        <span className="material-symbols-outlined text-[13px] font-bold">warning</span>
                                                                        <span>Note: Changing plant code will also update related daily entries, meters and targets to the new code. Data is not deleted.</span>
                                                                    </p>
                                                                )}
                                                            </React.Fragment>
                                                        )}
                                                    </div>
                                                );
                                            })}
                                        </div>
                                    )}

                                    <div className="flex items-center justify-end gap-2.5 pt-4 border-t border-slate-100">
                                        <button
                                            type="button"
                                            onClick={() => setIsMasterFormOpen(false)}
                                            className="px-4 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-50 transition bg-transparent cursor-pointer"
                                        >
                                            Cancel
                                        </button>
                                        <button
                                            type="submit"
                                            disabled={actionLoading}
                                            className="flex items-center gap-2 px-5 py-2 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl text-xs font-bold transition shadow-sm cursor-pointer disabled:opacity-50 border-none"
                                        >
                                            {actionLoading && <span className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />}
                                            <span>Save Record</span>
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>
                    )}

                    {/* Excel-format print sheet — only visible while printing via Print PDF */}
                    {plantPrintPayload && (
                        <div id="plant-excel-print-root" className="plant-excel-print-root" aria-hidden="true">
                            <div className="pex-title">UTILITY SENSE — Daily Utility Consumption Report</div>
                            <div className="pex-meta">
                                <div className="pex-label">LOCATION</div>
                                <div className="pex-value">{plantPrintPayload.locationLabel}</div>
                                <div className="pex-label">PLANT</div>
                                <div className="pex-value">{plantPrintPayload.plantLabel}</div>
                            </div>
                            <div className="pex-sub">
                                Generated: {new Date().toLocaleString("en-IN")} · Rows: {plantPrintPayload.list.length} · Rate: {plantPrintPayload.gridLabel}/Solar ₹{plantPrintPayload.electRate}/unit · Diesel ₹{plantPrintPayload.dieselRate}/L
                            </div>
                            <table className="pex-table">
                                <thead>
                                    <tr>
                                        {PLANT_REPORT_TYPE_ROW.map((v, i) => (
                                            <th key={`t${i}`} className={v === "Manual" ? "pex-manual" : "pex-green"}>{v}</th>
                                        ))}
                                    </tr>
                                    <tr>
                                        {(plantPrintPayload?.reportFormulaRow || []).map((v, i) => (
                                            <th key={`f${i}`} className={`${PLANT_REPORT_MANUAL_COLS.has(i) ? "pex-manual" : "pex-auto"} pex-formula`}>{v}</th>
                                        ))}
                                    </tr>
                                    <tr>
                                        {(plantPrintPayload?.reportHeaders || []).map((v, i) => (
                                            <th key={`h${i}`} className={PLANT_REPORT_MANUAL_COLS.has(i) ? "pex-manual" : "pex-auto"}>{v}</th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {plantPrintPayload.dataRows.map((row, idx) => (
                                        <tr key={idx}>
                                            {row.map((v, i) => (
                                                <td
                                                    key={i}
                                                    className={PLANT_REPORT_MANUAL_COLS.has(i) ? "pex-manual" : (idx % 2 ? "pex-alt" : "")}
                                                >
                                                    {v}
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}

                    {/* Audit Log Details Inspection Modal */}
                    {selectedAuditDetail && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
                            <div className="relative w-full max-w-2xl overflow-hidden rounded-2xl bg-white border border-slate-200 shadow-2xl flex flex-col max-h-[85vh]">
                                <div className="flex items-center justify-between border-b border-slate-100 px-6 py-4 bg-slate-50/70">
                                    <div className="flex items-center gap-2.5">
                                        <div className="h-8 w-8 rounded-lg bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
                                            <span className="material-symbols-outlined text-[18px]">manage_search</span>
                                        </div>
                                        <div>
                                            <h3 className="text-sm font-bold text-slate-800">
                                                Audit Event Details
                                            </h3>
                                            <p className="text-[10px] text-slate-400">
                                                {selectedAuditDetail.created_at ? new Date(selectedAuditDetail.created_at).toLocaleString("en-IN") : "—"}
                                            </p>
                                        </div>
                                    </div>
                                    <button
                                        type="button"
                                        onClick={() => setSelectedAuditDetail(null)}
                                        className="h-8 w-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600 transition border-none bg-transparent cursor-pointer"
                                    >
                                        <span className="material-symbols-outlined text-[18px]">close</span>
                                    </button>
                                </div>

                                <div className="p-6 overflow-y-auto space-y-4 text-xs">
                                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200/70">
                                        <div>
                                            <span className="text-[9px] font-bold uppercase text-slate-400">Action</span>
                                            <p className="font-bold text-slate-800 mt-0.5">{selectedAuditDetail.action}</p>
                                        </div>
                                        <div>
                                            <span className="text-[9px] font-bold uppercase text-slate-400">Module</span>
                                            <p className="font-bold text-slate-800 mt-0.5">{selectedAuditDetail.module}</p>
                                        </div>
                                        <div>
                                            <span className="text-[9px] font-bold uppercase text-slate-400">User</span>
                                            <p className="font-bold text-slate-800 mt-0.5">{selectedAuditDetail.user_name || selectedAuditDetail.user_email}</p>
                                        </div>
                                        <div>
                                            <span className="text-[9px] font-bold uppercase text-slate-400">Status</span>
                                            <p className="font-bold text-emerald-600 mt-0.5">{selectedAuditDetail.status}</p>
                                        </div>
                                    </div>

                                    {selectedAuditDetail.plant && (
                                        <div className="grid grid-cols-2 gap-3 bg-sky-50/50 p-3 rounded-xl border border-sky-100 text-[11px]">
                                            <div>
                                                <span className="font-bold text-slate-500">Plant:</span> <span className="font-extrabold text-sky-800">{selectedAuditDetail.plant}</span>
                                            </div>
                                            <div>
                                                <span className="font-bold text-slate-500">Location:</span> <span className="font-extrabold text-sky-800">{selectedAuditDetail.location || "—"}</span>
                                            </div>
                                        </div>
                                    )}

                                    {selectedAuditDetail.old_value && (
                                        <div>
                                            <span className="block font-bold text-slate-700 mb-1.5 flex items-center gap-1">
                                                <span className="material-symbols-outlined text-[14px] text-amber-500">history</span>
                                                <span>Previous Value (Before Edit)</span>
                                            </span>
                                            <pre className="p-3 bg-slate-900 text-emerald-400 rounded-xl overflow-x-auto text-[11px] font-mono leading-relaxed">
                                                {JSON.stringify(selectedAuditDetail.old_value, null, 2)}
                                            </pre>
                                        </div>
                                    )}

                                    {selectedAuditDetail.new_value && (
                                        <div>
                                            <span className="block font-bold text-slate-700 mb-1.5 flex items-center gap-1">
                                                <span className="material-symbols-outlined text-[14px] text-sky-500">update</span>
                                                <span>Updated / New Value</span>
                                            </span>
                                            <pre className="p-3 bg-slate-900 text-sky-300 rounded-xl overflow-x-auto text-[11px] font-mono leading-relaxed">
                                                {JSON.stringify(selectedAuditDetail.new_value, null, 2)}
                                            </pre>
                                        </div>
                                    )}

                                    {selectedAuditDetail.ip_device && (
                                        <div className="text-[10px] text-slate-400 border-t pt-2 font-mono truncate">
                                            Client Agent: {selectedAuditDetail.ip_device}
                                        </div>
                                    )}
                                </div>

                                <div className="px-6 py-3 bg-slate-50 border-t border-slate-100 flex justify-end">
                                    <button
                                        type="button"
                                        onClick={() => setSelectedAuditDetail(null)}
                                        className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl text-xs font-bold border-none cursor-pointer transition"
                                    >
                                        Close
                                    </button>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* Create / Edit Email Schedule Modal */}
                    {isScheduleModalOpen && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
                            <div className="relative w-full max-w-3xl overflow-hidden rounded-2xl bg-white border border-slate-200 shadow-2xl flex flex-col max-h-[90vh]">
                                <div className="flex items-center justify-between border-b border-slate-100 px-6 py-4 bg-slate-50/70">
                                    <div className="flex items-center gap-2.5">
                                        <div className="h-9 w-9 rounded-xl bg-sky-50 border border-sky-100 flex items-center justify-center text-sky-600">
                                            <span className="material-symbols-outlined text-[20px]">forward_to_inbox</span>
                                        </div>
                                        <div>
                                            <h3 className="text-sm font-bold text-slate-800">
                                                {editingSchedule ? "Edit Automated Email Schedule" : "Create New Email Schedule"}
                                            </h3>
                                            <p className="text-[11px] text-slate-500">
                                                Configure independent monthly Excel report dispatch for specific Location and Plant
                                            </p>
                                        </div>
                                    </div>
                                    <button
                                        type="button"
                                        onClick={() => setIsScheduleModalOpen(false)}
                                        className="h-8 w-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600 transition border-none bg-transparent cursor-pointer"
                                    >
                                        <span className="material-symbols-outlined text-[18px]">close</span>
                                    </button>
                                </div>

                                <form onSubmit={handleSaveSchedule} className="p-6 overflow-y-auto space-y-4 text-xs">
                                    {/* Row 1: Schedule Name & Status */}
                                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 items-end">
                                        <div className="sm:col-span-2">
                                            <label className="block text-[10.5px] font-bold text-slate-600 mb-1">Schedule Name</label>
                                            <input
                                                type="text"
                                                required
                                                placeholder="e.g. Bhiwadi Plant A Monthly Report"
                                                value={scheduleFormValues.name}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, name: e.target.value })}
                                                className="w-full h-9 border border-slate-200 rounded-xl px-3 bg-white text-slate-800 font-semibold focus:outline-none focus:border-sky-500 shadow-sm"
                                            />
                                        </div>
                                        <div className="p-2.5 bg-slate-50 rounded-xl border border-slate-200/70 flex items-center justify-between">
                                            <span className="font-bold text-slate-700 text-[11px]">Schedule Status</span>
                                            <label className="relative inline-flex items-center cursor-pointer">
                                                <input
                                                    type="checkbox"
                                                    checked={scheduleFormValues.enabled}
                                                    onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, enabled: e.target.checked })}
                                                    className="sr-only peer"
                                                />
                                                <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-emerald-500"></div>
                                                <span className={`ml-2 text-[10px] font-extrabold ${scheduleFormValues.enabled ? "text-emerald-600" : "text-slate-400"}`}>
                                                    {scheduleFormValues.enabled ? "ON" : "OFF"}
                                                </span>
                                            </label>
                                        </div>
                                    </div>

                                    {/* Row 2: Cascading Location & Plant Selection */}
                                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-sky-50/50 p-3.5 rounded-xl border border-sky-100">
                                        <div>
                                            <label className="block text-[10.5px] font-bold text-slate-700 mb-1">Location</label>
                                            <select
                                                required
                                                value={scheduleFormValues.location}
                                                onChange={(e) => {
                                                    const newLoc = e.target.value;
                                                    const filtered = plants.filter(p => p.location.toUpperCase() === newLoc.toUpperCase());
                                                    const newPlant = filtered[0]?.plant_code || "";
                                                    setScheduleFormValues({
                                                        ...scheduleFormValues,
                                                        location: newLoc,
                                                        plant: newPlant
                                                    });
                                                }}
                                                className="w-full h-9 border border-slate-200 rounded-xl px-3 bg-white text-slate-800 font-bold focus:outline-none focus:border-sky-500 shadow-sm"
                                            >
                                                <option value="">Select Location</option>
                                                {Array.from(new Set(plants.map(p => p.location.toUpperCase()))).map(loc => (
                                                    <option key={loc} value={loc}>{loc}</option>
                                                ))}
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-[10.5px] font-bold text-slate-700 mb-1">Plant (Strictly for {scheduleFormValues.location || "Selected Location"})</label>
                                            <select
                                                required
                                                value={scheduleFormValues.plant}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, plant: e.target.value })}
                                                className="w-full h-9 border border-slate-200 rounded-xl px-3 bg-white text-slate-800 font-bold focus:outline-none focus:border-sky-500 shadow-sm"
                                            >
                                                <option value="">Select Plant</option>
                                                {plants
                                                    .filter(p => !scheduleFormValues.location || p.location.toUpperCase() === scheduleFormValues.location.toUpperCase())
                                                    .map(p => (
                                                        <option key={p.plant_code} value={p.plant_code}>{p.plant_code} - {p.plant_name}</option>
                                                    ))
                                                }
                                            </select>
                                        </div>
                                    </div>

                                    {/* Row 3: TO, CC, BCC Recipients */}
                                    <div className="space-y-3">
                                        <div>
                                            <label className="block text-[10.5px] font-bold text-slate-700 mb-1">
                                                To Recipients <span className="text-red-500">*</span> (Comma Separated)
                                            </label>
                                            <input
                                                type="text"
                                                required
                                                placeholder="fh@example.com, manager@example.com, planthead@example.com"
                                                value={scheduleFormValues.to_recipients}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, to_recipients: e.target.value })}
                                                className="w-full h-9 border border-slate-200 rounded-xl px-3 bg-white text-slate-800 font-medium focus:outline-none focus:border-sky-500 shadow-sm text-xs"
                                            />
                                        </div>

                                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                            <div>
                                                <label className="block text-[10.5px] font-bold text-slate-600 mb-1">
                                                    CC Recipients (Optional)
                                                </label>
                                                <input
                                                    type="text"
                                                    placeholder="hod@example.com, it@example.com, finance@example.com"
                                                    value={scheduleFormValues.cc_recipients}
                                                    onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, cc_recipients: e.target.value })}
                                                    className="w-full h-9 border border-slate-200 rounded-xl px-3 bg-white text-slate-800 font-medium focus:outline-none focus:border-sky-500 shadow-sm text-xs"
                                                />
                                            </div>
                                            <div>
                                                <label className="block text-[10.5px] font-bold text-slate-600 mb-1">
                                                    BCC Recipients (Optional)
                                                </label>
                                                <input
                                                    type="text"
                                                    placeholder="audit@example.com"
                                                    value={scheduleFormValues.bcc_recipients}
                                                    onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, bcc_recipients: e.target.value })}
                                                    className="w-full h-9 border border-slate-200 rounded-xl px-3 bg-white text-slate-800 font-medium focus:outline-none focus:border-sky-500 shadow-sm text-xs"
                                                />
                                            </div>
                                        </div>
                                    </div>

                                    {/* Row 4: Frequency, Day, Time, Report Type */}
                                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200/70">
                                        <div>
                                            <label className="block text-[10px] font-bold text-slate-500 mb-1 uppercase">Report</label>
                                            <select
                                                value={scheduleFormValues.report_type}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, report_type: e.target.value })}
                                                className="w-full h-8.5 border border-slate-200 rounded-lg px-2 bg-white text-slate-800 font-bold focus:outline-none text-[11px]"
                                            >
                                                <option value="Monthly Utility Report">Monthly Utility Report</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-[10px] font-bold text-slate-500 mb-1 uppercase">Frequency</label>
                                            <select
                                                value={scheduleFormValues.frequency}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, frequency: e.target.value })}
                                                className="w-full h-8.5 border border-slate-200 rounded-lg px-2 bg-white text-slate-800 font-bold focus:outline-none text-[11px]"
                                            >
                                                <option value="Monthly">Monthly</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-[10px] font-bold text-slate-500 mb-1 uppercase">Day of Month</label>
                                            <select
                                                value={scheduleFormValues.schedule_day}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, schedule_day: Number(e.target.value) || 1 })}
                                                className="w-full h-8.5 border border-slate-200 rounded-lg px-2 bg-white text-slate-800 font-bold focus:outline-none text-[11px]"
                                            >
                                                {Array.from({ length: 28 }, (_, i) => i + 1).map(d => (
                                                    <option key={d} value={d}>{d}{d === 1 ? 'st' : d === 2 ? 'nd' : d === 3 ? 'rd' : 'th'} of month</option>
                                                ))}
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-[10px] font-bold text-slate-500 mb-1 uppercase">Time</label>
                                            <input
                                                type="time"
                                                required
                                                value={scheduleFormValues.schedule_time}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, schedule_time: e.target.value })}
                                                className="w-full h-8.5 border border-slate-200 rounded-lg px-2 bg-white text-slate-800 font-bold focus:outline-none text-[11px]"
                                            />
                                        </div>
                                    </div>

                                    {/* Row 5: Custom Subject & Body with Dynamic Variable Helpers */}
                                    <div className="space-y-3">
                                        <div>
                                            <label className="block text-[10.5px] font-bold text-slate-700 mb-1">Email Subject Template</label>
                                            <input
                                                type="text"
                                                required
                                                value={scheduleFormValues.subject_template}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, subject_template: e.target.value })}
                                                placeholder="Monthly Utility Report - {Location} - {Plant} - {Month}"
                                                className="w-full h-9 border border-slate-200 rounded-xl px-3 bg-white text-slate-800 font-bold focus:outline-none focus:border-sky-500 shadow-sm"
                                            />
                                        </div>

                                        <div>
                                            <label className="block text-[10.5px] font-bold text-slate-700 mb-1">Custom Email Body</label>
                                            <textarea
                                                rows={4}
                                                value={scheduleFormValues.body_template}
                                                onChange={(e) => setScheduleFormValues({ ...scheduleFormValues, body_template: e.target.value })}
                                                placeholder="Dear Team,\n\nPlease find attached the monthly UtilitySense report for {Location} - {Plant} for {Month}."
                                                className="w-full border border-slate-200 rounded-xl p-3 bg-white text-slate-800 font-mono text-xs focus:outline-none focus:border-sky-500 shadow-sm leading-relaxed"
                                            />
                                        </div>

                                        <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200/80 flex items-center justify-between text-[11px]">
                                            <span className="font-bold text-slate-500">Insert Placeholders:</span>
                                            <div className="flex flex-wrap gap-1.5">
                                                {['{Location}', '{Plant}', '{Month}', '{Year}', '{ReportDate}'].map(tag => (
                                                    <button
                                                        key={tag}
                                                        type="button"
                                                        onClick={() => setScheduleFormValues({
                                                            ...scheduleFormValues,
                                                            body_template: `${scheduleFormValues.body_template} ${tag}`
                                                        })}
                                                        className="px-1.5 py-0.5 bg-white border border-slate-300 hover:border-sky-400 hover:text-sky-700 rounded text-slate-600 font-mono text-[10px] cursor-pointer"
                                                    >
                                                        {tag}
                                                    </button>
                                                ))}
                                            </div>
                                        </div>
                                    </div>

                                    <div className="flex items-center justify-end gap-2.5 pt-4 border-t border-slate-100">
                                        <button
                                            type="button"
                                            onClick={() => setIsScheduleModalOpen(false)}
                                            className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold border-none cursor-pointer transition"
                                        >
                                            Cancel
                                        </button>
                                        <button
                                            type="submit"
                                            disabled={actionLoading}
                                            className="flex items-center gap-2 px-5 py-2 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl text-xs font-bold transition shadow-sm border-none cursor-pointer disabled:opacity-50"
                                        >
                                            {actionLoading && <span className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />}
                                            <span>{editingSchedule ? "Update Schedule" : "Save Email Schedule"}</span>
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>
                    )}

                    {/* Execution Logs Modal */}
                    {isScheduleLogsOpen && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
                            <div className="relative w-full max-w-4xl overflow-hidden rounded-2xl bg-white border border-slate-200 shadow-2xl flex flex-col max-h-[85vh]">
                                <div className="flex items-center justify-between border-b border-slate-100 px-6 py-4 bg-slate-50/70">
                                    <div className="flex items-center gap-2.5">
                                        <div className="h-8 w-8 rounded-lg bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
                                            <span className="material-symbols-outlined text-[18px]">history_edu</span>
                                        </div>
                                        <div>
                                            <h3 className="text-sm font-bold text-slate-800">
                                                Automated Schedule Execution Logs
                                            </h3>
                                            <p className="text-[10px] text-slate-400">
                                                Complete history of automated and manual email report dispatches
                                            </p>
                                        </div>
                                    </div>
                                    <button
                                        type="button"
                                        onClick={() => setIsScheduleLogsOpen(false)}
                                        className="h-8 w-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600 transition border-none bg-transparent cursor-pointer"
                                    >
                                        <span className="material-symbols-outlined text-[18px]">close</span>
                                    </button>
                                </div>

                                <div className="p-4 overflow-y-auto flex-1">
                                    {emailScheduleLogs.length === 0 ? (
                                        <div className="py-16 text-center text-slate-400 font-medium">
                                            No automated report executions logged yet.
                                        </div>
                                    ) : (
                                        <div className="overflow-x-auto">
                                            <table className="w-full border-collapse text-left text-xs">
                                                <thead>
                                                    <tr className="border-b border-slate-100 text-slate-400 font-bold uppercase tracking-wider text-[9px] bg-slate-50">
                                                        <th className="py-2.5 px-3">Execution Time</th>
                                                        <th className="py-2.5 px-3">Schedule Name</th>
                                                        <th className="py-2.5 px-3">Location / Plant</th>
                                                        <th className="py-2.5 px-3">Recipients</th>
                                                        <th className="py-2.5 px-3">Status</th>
                                                        <th className="py-2.5 px-3">Message / Error</th>
                                                    </tr>
                                                </thead>
                                                <tbody className="divide-y divide-slate-100 text-slate-700">
                                                    {emailScheduleLogs.map((log, idx) => (
                                                        <tr key={idx} className="hover:bg-slate-50/50">
                                                            <td className="py-2 px-3 font-mono text-[11px] text-slate-500 whitespace-nowrap">
                                                                {log.executed_at ? new Date(log.executed_at).toLocaleString("en-IN") : "—"}
                                                            </td>
                                                            <td className="py-2 px-3 font-bold text-slate-800">{log.schedule_name || log.schedule_id}</td>
                                                            <td className="py-2 px-3 font-semibold text-sky-700">{log.location} → {log.plant}</td>
                                                            <td className="py-2 px-3 text-[11px] text-slate-600 max-w-xs truncate" title={`To: ${log.to_recipients || ''} | CC: ${log.cc_recipients || ''}`}>
                                                                {log.to_recipients || "—"}
                                                            </td>
                                                            <td className="py-2 px-3">
                                                                <span className={`px-2 py-0.5 rounded text-[9px] font-bold ${log.status === "SUCCESS" ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>
                                                                    {log.status}
                                                                </span>
                                                            </td>
                                                            <td className="py-2 px-3 text-[11px] text-slate-500 font-mono truncate max-w-xs">
                                                                {log.error_message || "Excel report generated & delivered."}
                                                            </td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </div>
                                    )}
                                </div>

                                <div className="px-6 py-3 bg-slate-50 border-t border-slate-100 flex justify-end">
                                    <button
                                        type="button"
                                        onClick={() => setIsScheduleLogsOpen(false)}
                                        className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl text-xs font-bold border-none cursor-pointer transition"
                                    >
                                        Close
                                    </button>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* SMART MASS EXCEL IMPORTER MODAL (IT ADMIN ONLY) */}
                    {isMassUploadModalOpen && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-5 bg-slate-900/60 backdrop-blur-sm">
                            <div className="relative w-full max-w-5xl overflow-hidden rounded-2xl bg-white border border-slate-200 shadow-2xl flex flex-col max-h-[92vh]">
                                {/* Header */}
                                <div className="flex items-center justify-between border-b border-slate-100 px-6 py-4 bg-slate-50/80">
                                    <div className="flex items-center gap-3">
                                        <div className="h-10 w-10 rounded-xl bg-sky-50 border border-sky-100 flex items-center justify-center text-sky-600">
                                            <span className="material-symbols-outlined text-[22px]">upload_file</span>
                                        </div>
                                        <div>
                                            <div className="flex items-center gap-2">
                                                <h3 className="text-sm font-extrabold text-slate-900">
                                                    Smart Mass Excel Importer & Bulk Data Uploader
                                                </h3>
                                                <span className="px-2 py-0.5 rounded-full text-[9px] font-extrabold bg-indigo-50 text-indigo-700 border border-indigo-100">
                                                    IT Admin Only
                                                </span>
                                            </div>
                                            <p className="text-[11px] text-slate-500 mt-0.5">
                                                Intelligent column recognition, automatic Location/Plant mapping, duplicate-date protection & safe transactional import
                                            </p>
                                        </div>
                                    </div>
                                    <div className="flex items-center gap-2">
                                        <button
                                            type="button"
                                            onClick={() => setIsImportHistoryOpen(true)}
                                            className="flex items-center gap-1 h-8 px-3 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold transition shadow-xs cursor-pointer"
                                            title="View past import logs"
                                        >
                                            <span className="material-symbols-outlined text-[16px] text-indigo-600">history</span>
                                            <span>Import History</span>
                                        </button>
                                        <button
                                            type="button"
                                            onClick={() => setIsMassUploadModalOpen(false)}
                                            className="h-8 w-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600 transition border-none bg-transparent cursor-pointer"
                                        >
                                            <span className="material-symbols-outlined text-[18px]">close</span>
                                        </button>
                                    </div>
                                </div>

                                {/* Body */}
                                <div className="p-6 overflow-y-auto space-y-5 text-xs flex-1">
                                    {/* 1. File Upload / Selection Area */}
                                    {!massUploadFileName ? (
                                        <div className="border-2 border-dashed border-sky-200 hover:border-sky-400 rounded-2xl p-8 bg-sky-50/30 transition text-center flex flex-col items-center justify-center">
                                            <div className="h-12 w-12 rounded-full bg-sky-100 text-sky-600 flex items-center justify-center mb-3">
                                                <span className="material-symbols-outlined text-[28px]">cloud_upload</span>
                                            </div>
                                            <h4 className="text-sm font-bold text-slate-800 mb-1">
                                                Choose or Drag & Drop Excel File (.xlsx, .xls)
                                            </h4>
                                            <p className="text-[11px] text-slate-500 max-w-md mb-4">
                                                Supports multi-utility columns (Electricity, Solar, Diesel, ODU/IDU Production, Waste). System will auto-detect columns, locations, and plants.
                                            </p>
                                            <label className="flex items-center gap-2 px-5 py-2.5 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl font-bold cursor-pointer transition shadow-sm">
                                                {isAnalyzingExcel ? (
                                                    <>
                                                        <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                                                        <span>Analyzing Spreadsheet...</span>
                                                    </>
                                                ) : (
                                                    <>
                                                        <span className="material-symbols-outlined text-[18px]">folder_open</span>
                                                        <span>Browse Excel File</span>
                                                    </>
                                                )}
                                                <input
                                                    type="file"
                                                    accept=".xlsx, .xls, .csv"
                                                    onChange={handleMassExcelFileUpload}
                                                    className="hidden"
                                                    disabled={isAnalyzingExcel}
                                                />
                                            </label>
                                        </div>
                                    ) : (
                                        <div className="bg-slate-50 p-4 rounded-xl border border-slate-200/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                            <div className="flex items-center gap-3">
                                                <div className="h-10 w-10 rounded-xl bg-emerald-50 border border-emerald-100 flex items-center justify-center text-emerald-600 shrink-0">
                                                    <span className="material-symbols-outlined text-[22px]">description</span>
                                                </div>
                                                <div>
                                                    <div className="font-bold text-slate-900 text-sm flex items-center gap-2">
                                                        <span>{massUploadFileName}</span>
                                                        <span className="text-[10px] font-extrabold px-2 py-0.5 rounded bg-slate-200 text-slate-700">
                                                            {rawUploadData.length} Rows
                                                        </span>
                                                    </div>
                                                    <div className="text-[11px] text-slate-500 flex items-center gap-2 mt-0.5">
                                                        <span>Sheet:</span>
                                                        {massUploadSheets.length > 1 ? (
                                                            <select
                                                                value={selectedUploadSheet}
                                                                onChange={(e) => handleSheetChange(e.target.value)}
                                                                className="border border-slate-300 rounded px-1.5 py-0.5 bg-white font-semibold text-slate-700 text-xs"
                                                            >
                                                                {massUploadSheets.map(s => <option key={s} value={s}>{s}</option>)}
                                                            </select>
                                                        ) : (
                                                            <span className="font-semibold text-slate-700">{selectedUploadSheet || "Sheet1"}</span>
                                                        )}
                                                    </div>
                                                </div>
                                            </div>
                                            <label className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-200 hover:bg-slate-100 text-slate-700 rounded-xl text-xs font-bold transition shadow-xs cursor-pointer shrink-0">
                                                <span className="material-symbols-outlined text-[16px] text-slate-500">sync</span>
                                                <span>Upload Different File</span>
                                                <input
                                                    type="file"
                                                    accept=".xlsx, .xls, .csv"
                                                    onChange={handleMassExcelFileUpload}
                                                    className="hidden"
                                                />
                                            </label>
                                        </div>
                                    )}

                                    {/* 2. Intelligent Detection & Scope Overrides (When file is loaded) */}
                                    {analyzedImportRows.length > 0 && (
                                        <div className="space-y-4">
                                            {/* Scope & Auto-Detection Card */}
                                            <div className="p-4 bg-sky-50/40 rounded-xl border border-sky-100 space-y-3">
                                                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                                    <div>
                                                        <h4 className="text-xs font-extrabold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                                                            <span className="material-symbols-outlined text-sky-600 text-[18px]">tune</span>
                                                            <span>Location & Plant Routing Scope</span>
                                                        </h4>
                                                        <p className="text-[11px] text-slate-500">
                                                            Auto-detected from Excel columns or override manually for the entire file
                                                        </p>
                                                    </div>
                                                    <div className="flex items-center gap-2.5">
                                                        <div>
                                                            <label className="block text-[10px] font-bold text-slate-600 mb-0.5 uppercase">Location Scope</label>
                                                            <select
                                                                value={locationOverride}
                                                                onChange={(e) => handleLocationOverrideChange(e.target.value)}
                                                                className="h-8 border border-slate-200 rounded-lg px-2 bg-white text-slate-800 font-bold focus:outline-none text-xs"
                                                            >
                                                                <option value="auto">Auto Detect from Excel</option>
                                                                {Array.from(new Set(plants.map(p => p.location.toUpperCase()))).map(loc => (
                                                                    <option key={loc} value={loc}>{loc}</option>
                                                                ))}
                                                            </select>
                                                        </div>
                                                        <div>
                                                            <label className="block text-[10px] font-bold text-slate-600 mb-0.5 uppercase">Plant Scope</label>
                                                            <select
                                                                value={plantOverride}
                                                                onChange={(e) => handlePlantOverrideChange(e.target.value)}
                                                                className="h-8 border border-slate-200 rounded-lg px-2 bg-white text-slate-800 font-bold focus:outline-none text-xs"
                                                            >
                                                                <option value="auto">Auto Detect from Excel</option>
                                                                {plants
                                                                    .filter(p => locationOverride === "auto" || p.location.toUpperCase() === locationOverride.toUpperCase())
                                                                    .map(p => (
                                                                        <option key={p.plant_code} value={p.plant_code}>{p.plant_code} - {p.plant_name}</option>
                                                                    ))
                                                                }
                                                            </select>
                                                        </div>
                                                    </div>
                                                </div>

                                                {/* Detected Utility Columns Badges */}
                                                <div className="pt-2 border-t border-sky-100 flex flex-wrap items-center gap-2">
                                                    <span className="font-bold text-slate-600 text-[11px]">Detected Fields:</span>
                                                    {columnMappings.date && (
                                                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">
                                                            ✓ Date: <b className="font-mono">{columnMappings.date}</b>
                                                        </span>
                                                    )}
                                                    {columnMappings.electricity_closing && (
                                                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-sky-50 text-sky-700 border border-sky-200">
                                                            ✓ Electricity: <b className="font-mono">{columnMappings.electricity_closing}</b>
                                                        </span>
                                                    )}
                                                    {columnMappings.solar && (
                                                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-50 text-amber-700 border border-amber-200">
                                                            ✓ Solar: <b className="font-mono">{columnMappings.solar}</b>
                                                        </span>
                                                    )}
                                                    {columnMappings.diesel && (
                                                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-50 text-rose-700 border border-rose-200">
                                                            ✓ Diesel: <b className="font-mono">{columnMappings.diesel}</b>
                                                        </span>
                                                    )}
                                                    {(columnMappings.odu || columnMappings.idu || columnMappings.production_set) && (
                                                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-purple-50 text-purple-700 border border-purple-200">
                                                            ✓ Production: <b className="font-mono">{columnMappings.odu || columnMappings.production_set}</b>
                                                        </span>
                                                    )}
                                                    {(columnMappings.waste_hazardous || columnMappings.waste_non_hazardous) && (
                                                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-teal-50 text-teal-700 border border-teal-200">
                                                            ✓ Waste Data
                                                        </span>
                                                    )}
                                                </div>
                                            </div>

                                            {/* 3. KPI Analysis Cards */}
                                            {(() => {
                                                const total = analyzedImportRows.length;
                                                const valid = analyzedImportRows.filter(r => r.status === "VALID").length;
                                                const duplicate = analyzedImportRows.filter(r => r.status === "DUPLICATE").length;
                                                const invalid = analyzedImportRows.filter(r => r.status === "INVALID").length;

                                                const filteredRows = analyzedImportRows.filter(r => {
                                                    if (importFilterTab === "valid") return r.status === "VALID";
                                                    if (importFilterTab === "duplicate") return r.status === "DUPLICATE";
                                                    if (importFilterTab === "invalid") return r.status === "INVALID";
                                                    return true;
                                                });

                                                return (
                                                    <div className="space-y-4">
                                                        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                                                            <button
                                                                type="button"
                                                                onClick={() => setImportFilterTab("all")}
                                                                className={`p-3 rounded-xl border text-left cursor-pointer transition ${importFilterTab === "all" ? "bg-slate-100 border-slate-400 ring-2 ring-slate-400/20" : "bg-white border-slate-200 hover:bg-slate-50"}`}
                                                            >
                                                                <span className="text-[10px] font-bold text-slate-500 uppercase">Total Rows</span>
                                                                <div className="text-xl font-black text-slate-900 mt-0.5">{total}</div>
                                                            </button>
                                                            <button
                                                                type="button"
                                                                onClick={() => setImportFilterTab("valid")}
                                                                className={`p-3 rounded-xl border text-left cursor-pointer transition ${importFilterTab === "valid" ? "bg-emerald-50 border-emerald-400 ring-2 ring-emerald-400/20" : "bg-white border-slate-200 hover:bg-emerald-50/50"}`}
                                                            >
                                                                <span className="text-[10px] font-bold text-emerald-600 uppercase flex items-center gap-1">
                                                                    <span className="h-2 w-2 rounded-full bg-emerald-500" />
                                                                    <span>Ready to Import</span>
                                                                </span>
                                                                <div className="text-xl font-black text-emerald-700 mt-0.5">{valid}</div>
                                                            </button>
                                                            <button
                                                                type="button"
                                                                onClick={() => setImportFilterTab("duplicate")}
                                                                className={`p-3 rounded-xl border text-left cursor-pointer transition ${importFilterTab === "duplicate" ? "bg-amber-50 border-amber-400 ring-2 ring-amber-400/20" : "bg-white border-slate-200 hover:bg-amber-50/50"}`}
                                                            >
                                                                <span className="text-[10px] font-bold text-amber-600 uppercase flex items-center gap-1">
                                                                    <span className="h-2 w-2 rounded-full bg-amber-500" />
                                                                    <span>Duplicates (Skip)</span>
                                                                </span>
                                                                <div className="text-xl font-black text-amber-700 mt-0.5">{duplicate}</div>
                                                            </button>
                                                            <button
                                                                type="button"
                                                                onClick={() => setImportFilterTab("invalid")}
                                                                className={`p-3 rounded-xl border text-left cursor-pointer transition ${importFilterTab === "invalid" ? "bg-red-50 border-red-400 ring-2 ring-red-400/20" : "bg-white border-slate-200 hover:bg-red-50/50"}`}
                                                            >
                                                                <span className="text-[10px] font-bold text-red-600 uppercase flex items-center gap-1">
                                                                    <span className="h-2 w-2 rounded-full bg-red-500" />
                                                                    <span>Errors / Invalid</span>
                                                                </span>
                                                                <div className="text-xl font-black text-red-700 mt-0.5">{invalid}</div>
                                                            </button>
                                                        </div>

                                                        {/* 4. Import Preview Table */}
                                                        <div className="border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                                                            <div className="p-3 bg-slate-50 border-b border-slate-200/80 flex items-center justify-between text-xs">
                                                                <span className="font-bold text-slate-700">
                                                                    Previewing {filteredRows.length} of {total} records ({importFilterTab.toUpperCase()})
                                                                </span>
                                                                <span className="text-[10px] text-slate-400">
                                                                    Showing first 50 rows in preview
                                                                </span>
                                                            </div>
                                                            <div className="overflow-x-auto max-h-64">
                                                                <table className="w-full border-collapse text-left text-[11px]">
                                                                    <thead>
                                                                        <tr className="border-b border-slate-200 text-slate-500 font-bold uppercase tracking-wider text-[9px] bg-slate-50/80 sticky top-0">
                                                                            <th className="py-2 px-3">Row</th>
                                                                            <th className="py-2 px-3">Date</th>
                                                                            <th className="py-2 px-3">Location</th>
                                                                            <th className="py-2 px-3">Plant</th>
                                                                            <th className="py-2 px-3 text-right">MSEB Units</th>
                                                                            <th className="py-2 px-3 text-right">Solar (kWh)</th>
                                                                            <th className="py-2 px-3 text-right">Diesel (L)</th>
                                                                            <th className="py-2 px-3 text-right">Production</th>
                                                                            <th className="py-2 px-3 text-right">Total Cost</th>
                                                                            <th className="py-2 px-3">Status</th>
                                                                            <th className="py-2 px-3">Validation Details</th>
                                                                        </tr>
                                                                    </thead>
                                                                    <tbody className="divide-y divide-slate-100 text-slate-700">
                                                                        {filteredRows.slice(0, 50).map((r) => (
                                                                            <tr key={r.rowIndex} className={`hover:bg-slate-50/70 transition ${r.status === "INVALID" ? "bg-red-50/20" : r.status === "DUPLICATE" ? "bg-amber-50/20" : ""}`}>
                                                                                <td className="py-2 px-3 font-mono text-slate-400">{r.rowIndex}</td>
                                                                                <td className="py-2 px-3 font-bold text-slate-900 whitespace-nowrap">{r.date || "—"}</td>
                                                                                <td className="py-2 px-3 font-semibold text-slate-800 uppercase">{r.location || "—"}</td>
                                                                                <td className="py-2 px-3 font-bold text-sky-700">{r.plant || "—"}</td>
                                                                                <td className="py-2 px-3 text-right font-mono text-slate-800">{fmtNum(r.electricity_consumption)}</td>
                                                                                <td className="py-2 px-3 text-right font-mono text-amber-700">{fmtNum(r.solar_generated)}</td>
                                                                                <td className="py-2 px-3 text-right font-mono text-slate-700">{fmtNum(r.diesel_used)}</td>
                                                                                <td className="py-2 px-3 text-right font-semibold text-slate-800">{fmtNum(r.production_set)}</td>
                                                                                <td className="py-2 px-3 text-right font-mono font-bold text-emerald-600">{fmtMoney(r.total_cost)}</td>
                                                                                <td className="py-2 px-3 whitespace-nowrap">
                                                                                    {r.status === "VALID" && (
                                                                                        <span className="px-2 py-0.5 rounded text-[9px] font-extrabold bg-emerald-50 text-emerald-700 border border-emerald-100">
                                                                                            ✓ READY
                                                                                        </span>
                                                                                    )}
                                                                                    {r.status === "DUPLICATE" && (
                                                                                        <span className="px-2 py-0.5 rounded text-[9px] font-extrabold bg-amber-50 text-amber-700 border border-amber-200">
                                                                                            ⚠ DUPLICATE (SKIP)
                                                                                        </span>
                                                                                    )}
                                                                                    {r.status === "INVALID" && (
                                                                                        <span className="px-2 py-0.5 rounded text-[9px] font-extrabold bg-red-50 text-red-700 border border-red-200">
                                                                                            ❌ ERROR
                                                                                        </span>
                                                                                    )}
                                                                                </td>
                                                                                <td className="py-2 px-3 text-[10.5px] max-w-xs truncate" title={[...r.errors, ...r.warnings].join("; ")}>
                                                                                    {r.errors.length > 0 ? (
                                                                                        <span className="text-red-600 font-semibold">{r.errors[0]}</span>
                                                                                    ) : r.warnings.length > 0 ? (
                                                                                        <span className="text-amber-700 font-medium">{r.warnings[0]}</span>
                                                                                    ) : (
                                                                                        <span className="text-emerald-600">All checks passed</span>
                                                                                    )}
                                                                                </td>
                                                                            </tr>
                                                                        ))}
                                                                    </tbody>
                                                                </table>
                                                            </div>
                                                        </div>
                                                    </div>
                                                );
                                            })()}
                                        </div>
                                    )}

                                    {/* 5. Post-Import Success Summary */}
                                    {importResultSummary && (
                                        <div className="p-4 bg-emerald-50 rounded-xl border border-emerald-200 text-emerald-900 space-y-2">
                                            <div className="flex items-center gap-2 font-black text-sm text-emerald-800">
                                                <span className="material-symbols-outlined text-emerald-600">task_alt</span>
                                                <span>Mass Import Processed Successfully</span>
                                            </div>
                                            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 pt-1 text-xs">
                                                <div>Total Processed: <b>{importResultSummary.totalRows}</b></div>
                                                <div>New Created: <b className="text-emerald-700">{importResultSummary.importedCount}</b></div>
                                                <div>Duplicates Protected: <b className="text-amber-700">{importResultSummary.skippedDuplicates}</b></div>
                                                <div>Invalid Skipped: <b className="text-red-700">{importResultSummary.invalidCount}</b></div>
                                            </div>
                                        </div>
                                    )}
                                </div>

                                {/* Footer Actions */}
                                <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3">
                                    <button
                                        type="button"
                                        onClick={() => setIsImportHistoryOpen(true)}
                                        className="flex items-center gap-1.5 text-xs font-bold text-indigo-700 hover:text-indigo-900 bg-transparent border-none cursor-pointer"
                                    >
                                        <span className="material-symbols-outlined text-[16px]">list_alt</span>
                                        <span>View Previous Imports ({importHistory.length})</span>
                                    </button>

                                    <div className="flex items-center gap-2.5">
                                        <button
                                            type="button"
                                            onClick={() => setIsMassUploadModalOpen(false)}
                                            className="px-4 py-2 bg-white border border-slate-200 hover:bg-slate-100 text-slate-700 rounded-xl text-xs font-bold cursor-pointer transition shadow-xs"
                                        >
                                            Close
                                        </button>

                                        {analyzedImportRows.length > 0 && (
                                            <button
                                                type="button"
                                                disabled={isImportingData || analyzedImportRows.filter(r => r.status === "VALID").length === 0}
                                                onClick={handleExecuteMassImport}
                                                className="flex items-center gap-2 px-5 py-2 bg-[#0284c7] hover:bg-[#0369a1] text-white rounded-xl text-xs font-bold transition shadow-sm border-none cursor-pointer disabled:opacity-50"
                                            >
                                                {isImportingData ? (
                                                    <>
                                                        <span className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                                                        <span>Importing Transactions...</span>
                                                    </>
                                                ) : (
                                                    <>
                                                        <span className="material-symbols-outlined text-[16px]">check_circle</span>
                                                        <span>
                                                            Confirm & Import ({analyzedImportRows.filter(r => r.status === "VALID").length} Valid Rows)
                                                        </span>
                                                    </>
                                                )}
                                            </button>
                                        )}
                                    </div>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* IMPORT HISTORY MODAL */}
                    {isImportHistoryOpen && (
                        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
                            <div className="relative w-full max-w-4xl overflow-hidden rounded-2xl bg-white border border-slate-200 shadow-2xl flex flex-col max-h-[85vh]">
                                <div className="flex items-center justify-between border-b border-slate-100 px-6 py-4 bg-slate-50/80">
                                    <div className="flex items-center gap-2.5">
                                        <div className="h-8 w-8 rounded-lg bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
                                            <span className="material-symbols-outlined text-[18px]">history</span>
                                        </div>
                                        <div>
                                            <h3 className="text-sm font-bold text-slate-800">
                                                Mass Excel Import History & Audit Trail
                                            </h3>
                                            <p className="text-[10px] text-slate-400">
                                                Complete log of previous bulk uploads, rows imported, and duplicate protections
                                            </p>
                                        </div>
                                    </div>
                                    <button
                                        type="button"
                                        onClick={() => setIsImportHistoryOpen(false)}
                                        className="h-8 w-8 rounded-lg flex items-center justify-center text-slate-400 hover:bg-slate-100 hover:text-slate-600 transition border-none bg-transparent cursor-pointer"
                                    >
                                        <span className="material-symbols-outlined text-[18px]">close</span>
                                    </button>
                                </div>

                                <div className="p-4 overflow-y-auto flex-1">
                                    {importHistory.length === 0 ? (
                                        <div className="py-16 text-center text-slate-400 font-medium">
                                            No mass imports logged yet.
                                        </div>
                                    ) : (
                                        <div className="overflow-x-auto">
                                            <table className="w-full border-collapse text-left text-xs">
                                                <thead>
                                                    <tr className="border-b border-slate-100 text-slate-400 font-bold uppercase tracking-wider text-[9px] bg-slate-50">
                                                        <th className="py-2.5 px-3">Date / Time</th>
                                                        <th className="py-2.5 px-3">File Name</th>
                                                        <th className="py-2.5 px-3">Uploaded By</th>
                                                        <th className="py-2.5 px-3">Scope</th>
                                                        <th className="py-2.5 px-3 text-right">Total</th>
                                                        <th className="py-2.5 px-3 text-right text-emerald-700">Imported</th>
                                                        <th className="py-2.5 px-3 text-right text-amber-700">Protected</th>
                                                        <th className="py-2.5 px-3 text-right text-red-700">Failed</th>
                                                        <th className="py-2.5 px-3">Status</th>
                                                    </tr>
                                                </thead>
                                                <tbody className="divide-y divide-slate-100 text-slate-700">
                                                    {importHistory.map((h, idx) => (
                                                        <tr key={idx} className="hover:bg-slate-50/50">
                                                            <td className="py-2 px-3 font-mono text-[11px] text-slate-500 whitespace-nowrap">
                                                                {h.created_at ? new Date(h.created_at).toLocaleString("en-IN") : "—"}
                                                            </td>
                                                            <td className="py-2 px-3 font-bold text-slate-800">{h.file_name}</td>
                                                            <td className="py-2 px-3 text-slate-600">{h.uploaded_by}</td>
                                                            <td className="py-2 px-3 text-sky-700 font-semibold">{h.location} / {h.plant}</td>
                                                            <td className="py-2 px-3 text-right font-mono">{h.total_rows}</td>
                                                            <td className="py-2 px-3 text-right font-mono font-bold text-emerald-700">{h.imported_rows}</td>
                                                            <td className="py-2 px-3 text-right font-mono text-amber-700">{h.skipped_rows}</td>
                                                            <td className="py-2 px-3 text-right font-mono text-red-700">{h.failed_rows}</td>
                                                            <td className="py-2 px-3">
                                                                <span className={`px-2 py-0.5 rounded text-[9px] font-bold ${h.status === "COMPLETED" ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>
                                                                    {h.status || "COMPLETED"}
                                                                </span>
                                                            </td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </div>
                                    )}
                                </div>

                                <div className="px-6 py-3 bg-slate-50 border-t border-slate-100 flex justify-end">
                                    <button
                                        type="button"
                                        onClick={() => setIsImportHistoryOpen(false)}
                                        className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-xl text-xs font-bold border-none cursor-pointer transition"
                                    >
                                        Close
                                    </button>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* Custom Confirm Dialog — replaces native window.confirm() */}
                    {confirmDialog && (
                        <div className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
                            <div className="w-full max-w-sm rounded-2xl bg-white border border-slate-200 shadow-2xl p-6 text-center animate-[ccPop_0.18s_ease-out]">
                                <div className={`w-13 h-13 mx-auto mb-4 rounded-full flex items-center justify-center ${confirmDialog.danger ? "bg-red-50 text-red-600" : "bg-sky-50 text-sky-600"}`} style={{ width: 52, height: 52 }}>
                                    <span className="material-symbols-outlined text-[26px]">
                                        {confirmDialog.danger ? "warning" : "help"}
                                    </span>
                                </div>
                                <h3 className="text-sm font-bold text-slate-900 mb-2">{confirmDialog.title}</h3>
                                <p className="text-xs text-slate-500 leading-relaxed mb-6 whitespace-pre-line">{confirmDialog.message}</p>
                                <div className="flex items-center gap-2.5">
                                    <button
                                        onClick={() => { confirmDialog.resolve(false); setConfirmDialog(null); }}
                                        className="flex-1 h-10 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold border-none cursor-pointer transition"
                                    >
                                        {confirmDialog.cancelText}
                                    </button>
                                    <button
                                        onClick={() => { confirmDialog.resolve(true); setConfirmDialog(null); }}
                                        className={`flex-1 h-10 rounded-xl text-white text-xs font-bold border-none cursor-pointer transition shadow-sm ${confirmDialog.danger ? "bg-red-600 hover:bg-red-700" : "bg-[#0284c7] hover:bg-[#0369a1]"}`}
                                    >
                                        {confirmDialog.confirmText}
                                    </button>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* Toast Alert popup notifications */}
                    {toast && (
                        <div className="fixed bottom-5 right-5 z-50 flex items-center gap-2.5 px-4.5 py-3.5 rounded-2xl border shadow-xl bg-white border-slate-200 text-slate-800 animate-bounce">
                            {toast.type === "success" ? (
                                <span className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse" />
                            ) : (
                                <span className="material-symbols-outlined text-[16px] text-red-500">warning</span>
                            )}
                            <span className="text-xs font-bold">{toast.message}</span>
                            {toast.action && (
                                <button
                                    onClick={() => { toast.action.onClick(); setToast(null); }}
                                    className="ml-1 px-2.5 py-1 bg-slate-900 hover:bg-slate-700 text-white rounded-lg text-[10px] font-bold cursor-pointer border-none transition"
                                >
                                    {toast.action.label}
                                </button>
                            )}
                            <button onClick={() => setToast(null)} className="ml-1 text-slate-400 hover:text-slate-600 bg-transparent border-none cursor-pointer">
                                <span className="material-symbols-outlined text-[14px]">close</span>
                            </button>
                        </div>
                    )}
                </div>
                <IdleScreensaver logoSrc={PG_LOGO_BASE_64} idleMinutes={idleMinutes} />
                </>
            );
        }

        // Render App
        
    
export default App;
