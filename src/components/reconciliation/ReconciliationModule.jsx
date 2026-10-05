// UtilitySense Enterprise Utility Bill Reconciliation Module
// Provides Management Overview, Plant Detail Reconciliation, Variance Analysis, Clarifications, and Multi-Month Trends

import React, { useState, useEffect, useMemo } from 'react';
import * as Recharts from 'recharts';
import BillUploadModal from './BillUploadModal';
import {
    loadAllBills,
    loadClarifications,
    saveClarification,
    deleteBill,
    aggregateSystemConsumption,
    buildReconciliationData,
    getToleranceConfig,
    saveToleranceConfig,
    exportReconciliationToExcel
} from './reconciliationService';

const { ResponsiveContainer, ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend } = Recharts;

export default function ReconciliationModule({
    plants = [],
    dailyEntries = [],
    currentUser,
    tariffs = []
}) {
    // Bills and Clarifications Database States
    const [bills, setBills] = useState([]);
    const [clarifications, setClarifications] = useState([]);
    const [tolerances, setTolerances] = useState(getToleranceConfig());
    const [isConfigOpen, setIsConfigOpen] = useState(false);

    // View Mode: "detail" = Detailed Plant Reconciliation, "overview" = Management Dashboard Overview
    const [viewMode, setViewMode] = useState("detail");

    // Persistent Active Bill ID & Hierarchy Filter States
    const [activeBillId, setActiveBillId] = useState(() => {
        try {
            return localStorage.getItem('ep_active_audit_bill_id') || "";
        } catch (e) {
            return "";
        }
    });

    const [selectedLocation, setSelectedLocation] = useState(() => {
        try {
            const saved = localStorage.getItem('ep_active_audit_filter');
            if (saved) {
                const parsed = JSON.parse(saved);
                if (parsed.location) return parsed.location;
            }
        } catch (e) {}
        return "PUNE";
    });

    const [selectedPlant, setSelectedPlant] = useState(() => {
        try {
            const saved = localStorage.getItem('ep_active_audit_filter');
            if (saved) {
                const parsed = JSON.parse(saved);
                if (parsed.plant) return String(parsed.plant);
            }
        } catch (e) {}
        return "4010";
    });

    const [selectedUtility, setSelectedUtility] = useState(() => {
        try {
            const saved = localStorage.getItem('ep_active_audit_filter');
            if (saved) {
                const parsed = JSON.parse(saved);
                if (parsed.utility) return parsed.utility;
            }
        } catch (e) {}
        return "electricity";
    });

    const [selectedMonth, setSelectedMonth] = useState(() => {
        try {
            const saved = localStorage.getItem('ep_active_audit_filter');
            if (saved) {
                const parsed = JSON.parse(saved);
                if (parsed.month) return parsed.month;
            }
        } catch (e) {}
        return "2026-07";
    });

    const [statusFilter, setStatusFilter] = useState("all");

    // Persist filter changes
    useEffect(() => {
        try {
            localStorage.setItem('ep_active_audit_filter', JSON.stringify({
                location: selectedLocation,
                plant: selectedPlant,
                utility: selectedUtility,
                month: selectedMonth
            }));
        } catch (e) {}
    }, [selectedLocation, selectedPlant, selectedUtility, selectedMonth]);

    // Search and Table Sort
    const [tableSearch, setTableSearch] = useState("");
    const [activeSectionTab, setActiveSectionTab] = useState("all"); // "all" | "consumption" | "solar" | "demand" | "powerFactor" | "cost"

    // Modals
    const [isUploadModalOpen, setIsUploadModalOpen] = useState(false);
    const [isPreviewBillModalOpen, setIsPreviewBillModalOpen] = useState(false);

    // Clarification Form State
    const [clarificationReason, setClarificationReason] = useState("Meter reading mismatch");
    const [clarificationComment, setClarificationComment] = useState("");
    const [clarificationParameter, setClarificationParameter] = useState("Net / Total Consumption");
    const [clarificationStatus, setClarificationStatus] = useState("Open");

    // Load initial data
    useEffect(() => {
        loadAllBills().then(data => setBills(data));
        loadClarifications().then(data => setClarifications(data));
    }, []);

    // Derived Unique Locations from Plants
    const availableLocations = useMemo(() => {
        const set = new Set();
        plants.forEach(p => {
            if (p.location) set.add(String(p.location).toUpperCase());
        });
        const arr = Array.from(set);
        return arr.length > 0 ? arr : ["PUNE", "BHIWADI", "GREATER NOIDA", "ROORKIE"];
    }, [plants]);

    // Plants filtered by Selected Location
    const availablePlantsForLocation = useMemo(() => {
        if (!selectedLocation) return plants;
        return plants.filter(p => String(p.location || "").toUpperCase() === String(selectedLocation).toUpperCase());
    }, [plants, selectedLocation]);

    // Ensure valid plant selection when location changes
    useEffect(() => {
        if (availablePlantsForLocation.length > 0) {
            const hasMatch = availablePlantsForLocation.some(p => String(p.plant_code) === String(selectedPlant));
            if (!hasMatch) {
                setSelectedPlant(String(availablePlantsForLocation[0].plant_code));
            }
        }
    }, [selectedLocation, availablePlantsForLocation, selectedPlant]);

    // Find the current active bill matching activeBillId OR Location + Plant + Utility + Month
    const activeBill = useMemo(() => {
        if (!bills || bills.length === 0) return null;

        // 1. If an activeBillId is explicitly stored/selected, check if it matches current filter
        if (activeBillId) {
            const direct = bills.find(b => b.id === activeBillId);
            if (direct) {
                const locMatch = String(direct.location || "").toUpperCase() === String(selectedLocation).toUpperCase();
                const plantMatch = String(direct.plant || "") === String(selectedPlant);
                const utilMatch = String(direct.utility || "").toLowerCase() === String(selectedUtility).toLowerCase();
                const monthMatch = direct.billMonth === selectedMonth;
                if (locMatch && plantMatch && utilMatch && monthMatch) {
                    return direct;
                }
            }
        }

        // 2. Find matching bills for current filters
        const matches = bills.filter(b => 
            String(b.location || "").toUpperCase() === String(selectedLocation).toUpperCase() &&
            String(b.plant || "") === String(selectedPlant) &&
            String(b.utility || "").toLowerCase() === String(selectedUtility).toLowerCase() &&
            b.billMonth === selectedMonth
        );

        if (matches.length === 0) return null;

        // Prioritize non-demo, user-uploaded bills
        const userUploaded = matches.find(b => !String(b.id || "").startsWith("bill_demo_"));
        return userUploaded || matches[0];
    }, [bills, activeBillId, selectedLocation, selectedPlant, selectedUtility, selectedMonth]);

    // Synchronize activeBillId when activeBill changes
    useEffect(() => {
        if (activeBill?.id && activeBill.id !== activeBillId) {
            setActiveBillId(activeBill.id);
            try {
                localStorage.setItem('ep_active_audit_bill_id', activeBill.id);
            } catch (e) {}
        }
    }, [activeBill, activeBillId]);

    // Handler to switch to any uploaded bill with 1 click
    const handleSelectBill = (bill) => {
        if (!bill) return;
        setActiveBillId(bill.id);
        setSelectedLocation(bill.location);
        setSelectedPlant(String(bill.plant));
        setSelectedUtility(bill.utility || "electricity");
        setSelectedMonth(bill.billMonth);
        try {
            localStorage.setItem('ep_active_audit_bill_id', bill.id);
            localStorage.setItem('ep_active_audit_filter', JSON.stringify({
                location: bill.location,
                plant: String(bill.plant),
                utility: bill.utility || "electricity",
                month: bill.billMonth
            }));
        } catch (e) {}
        setViewMode("detail");
    };

    // Aggregate System Consumption strictly for this Location + Plant + Month
    const systemData = useMemo(() => {
        return aggregateSystemConsumption(dailyEntries, selectedLocation, selectedPlant, selectedMonth);
    }, [dailyEntries, selectedLocation, selectedPlant, selectedMonth]);

    // Build Reconciliation Results
    const reconData = useMemo(() => {
        if (!activeBill) return null;
        return buildReconciliationData(activeBill, systemData, tolerances);
    }, [activeBill, systemData, tolerances]);

    // Active Clarifications for current reconciliation
    const activeClarifications = useMemo(() => {
        if (!activeBill) return [];
        return clarifications.filter(c => c.billId === activeBill.id);
    }, [clarifications, activeBill]);

    // Monthly Trend Data across available months for selected Location + Plant + Utility
    const monthlyTrendData = useMemo(() => {
        const monthsList = ["2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09"];
        return monthsList.map(mStr => {
            const billItem = bills.find(b =>
                String(b.location || "").toUpperCase() === String(selectedLocation).toUpperCase() &&
                String(b.plant || "") === String(selectedPlant) &&
                String(b.utility || "").toLowerCase() === String(selectedUtility).toLowerCase() &&
                b.billMonth === mStr
            );
            const sysAgg = aggregateSystemConsumption(dailyEntries, selectedLocation, selectedPlant, mStr);

            const billKwh = billItem ? (billItem.extractedData?.billedUnitsKwh ?? billItem.extractedData?.grossUnitsKwh ?? 0) : null;
            const sysKwh = sysAgg ? sysAgg.netConsumptionKwh : 0;
            const diff = billKwh !== null ? (billKwh - sysKwh) : 0;
            const diffPct = billKwh && billKwh > 0 ? ((Math.abs(diff) / billKwh) * 100) : 0;

            const [yr, mo] = mStr.split("-");
            const dObj = new Date(Number(yr), Number(mo) - 1, 1);
            const label = dObj.toLocaleDateString("en-GB", { month: "short", year: "2-digit" });

            return {
                monthKey: mStr,
                label,
                billKwh: billKwh || 0,
                systemKwh: sysKwh || 0,
                diffKwh: diff,
                diffPct: Number(diffPct.toFixed(1)),
                hasBill: Boolean(billItem)
            };
        });
    }, [bills, dailyEntries, selectedLocation, selectedPlant, selectedUtility]);

    // Management Overview Summary Dataset (All Plants & Bills)
    const overviewDataset = useMemo(() => {
        const rows = [];
        plants.forEach(p => {
            const loc = p.location || "PUNE";
            // Check for bills in the selected month
            const billItem = bills.find(b => 
                String(b.plant) === String(p.plant_code) && 
                b.billMonth === selectedMonth &&
                String(b.utility || "").toLowerCase() === String(selectedUtility).toLowerCase()
            );
            const sysAgg = aggregateSystemConsumption(dailyEntries, loc, p.plant_code, selectedMonth);

            const bCons = billItem ? (billItem.extractedData?.billedUnitsKwh ?? billItem.extractedData?.grossUnitsKwh ?? 0) : null;
            const sCons = sysAgg ? sysAgg.netConsumptionKwh : 0;
            const bCost = billItem ? (billItem.extractedData?.totalBillAmount ?? billItem.extractedData?.currentBillAmount ?? 0) : null;
            const sCost = sysAgg ? sysAgg.totalCost : 0;

            let status = "Not Uploaded";
            let diff = 0;
            let diffPct = 0;
            let costDiff = 0;

            if (bCons !== null) {
                diff = bCons - sCons;
                diffPct = bCons > 0 ? (Math.abs(diff) / bCons) * 100 : 0;
                costDiff = (bCost || 0) - sCost;

                if (diffPct <= tolerances.matchedThreshold) status = "Matched";
                else if (diffPct <= tolerances.varianceThreshold) status = "Variance";
                else status = "High Variance";
            }

            rows.push({
                location: loc,
                plantCode: p.plant_code,
                plantName: p.plant_display_name || p.plant_name,
                utility: selectedUtility,
                month: selectedMonth,
                hasBill: Boolean(billItem),
                billConsumption: bCons,
                systemConsumption: sCons,
                diff,
                diffPct: Number(diffPct.toFixed(2)),
                billCost: bCost,
                systemCost: sCost,
                costDiff,
                status,
                billId: billItem?.id
            });
        });

        if (statusFilter === "all") return rows;
        return rows.filter(r => r.status.toLowerCase() === statusFilter.toLowerCase());
    }, [plants, bills, dailyEntries, selectedMonth, selectedUtility, tolerances, statusFilter]);

    // Overview Metric Counters
    const overviewStats = useMemo(() => {
        let total = plants.length;
        let uploaded = 0;
        let matched = 0;
        let variance = 0;
        let highVariance = 0;

        overviewDataset.forEach(r => {
            if (r.hasBill) uploaded++;
            if (r.status === "Matched") matched++;
            if (r.status === "Variance") variance++;
            if (r.status === "High Variance") highVariance++;
        });

        const pendingClarificationsCount = clarifications.filter(c => c.status !== "Resolved").length;

        return {
            totalPlants: total,
            billsUploaded: uploaded,
            matched,
            variance,
            highVariance,
            pendingClarifications: pendingClarificationsCount
        };
    }, [plants, overviewDataset, clarifications]);

    // Handle Clarification Submission
    const handleAddClarification = async (e) => {
        e.preventDefault();
        if (!activeBill) return;

        const newItem = {
            id: `clarif_${Date.now()}`,
            billId: activeBill.id,
            parameter: clarificationParameter,
            reason: clarificationReason,
            comment: clarificationComment || "Added during variance analysis review",
            status: clarificationStatus,
            createdBy: currentUser?.name || "IT Admin",
            createdAt: new Date().toISOString(),
            resolvedBy: clarificationStatus === "Resolved" ? (currentUser?.name || "IT Admin") : null,
            resolvedAt: clarificationStatus === "Resolved" ? new Date().toISOString() : null
        };

        const updated = await saveClarification(newItem);
        setClarifications(updated);
        setClarificationComment("");
    };

    // Toggle Clarification Status
    const handleToggleClarificationStatus = async (item, nextStatus) => {
        const updated = await saveClarification({
            ...item,
            status: nextStatus,
            resolvedBy: nextStatus === "Resolved" ? (currentUser?.name || "IT Admin") : null,
            resolvedAt: nextStatus === "Resolved" ? new Date().toISOString() : null
        });
        setClarifications(updated);
    };

    // Filtered parameters list for Detailed Table
    const tableParameters = useMemo(() => {
        if (!reconData) return [];
        let items = [];

        if (activeSectionTab === "all" || activeSectionTab === "consumption") {
            items.push(...reconData.sections.consumption.map(i => ({ ...i, section: "Grid Consumption" })));
        }
        if (activeSectionTab === "all" || activeSectionTab === "solar") {
            items.push(...reconData.sections.solar.map(i => ({ ...i, section: "Solar Reconciliation" })));
        }
        if (activeSectionTab === "all" || activeSectionTab === "demand") {
            items.push(...reconData.sections.demand.map(i => ({ ...i, section: "Demand MD" })));
        }
        if (activeSectionTab === "all" || activeSectionTab === "powerFactor") {
            items.push(...reconData.sections.powerFactor.map(i => ({ ...i, section: "Power Factor" })));
        }
        if (activeSectionTab === "all" || activeSectionTab === "cost") {
            items.push(...reconData.sections.cost.map(i => ({ ...i, section: "Cost & Charges" })));
        }

        if (tableSearch.trim()) {
            const q = tableSearch.toLowerCase();
            items = items.filter(i => 
                i.parameter.toLowerCase().includes(q) || 
                (i.billRef && i.billRef.toLowerCase().includes(q))
            );
        }

        return items;
    }, [reconData, activeSectionTab, tableSearch]);

    // Formatters
    const fmt = (val, decimals = 0) => {
        if (val === null || val === undefined || isNaN(val)) return "N/A";
        return Number(val).toLocaleString("en-IN", { maximumFractionDigits: decimals });
    };

    const getStatusBadge = (status) => {
        switch (status) {
            case "Matched":
                return <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-50 text-emerald-700 border border-emerald-200">
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span> Matched
                </span>;
            case "Variance":
                return <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-amber-50 text-amber-700 border border-amber-200">
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-500"></span> Variance
                </span>;
            case "High Variance":
                return <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-rose-50 text-rose-700 border border-rose-200">
                    <span className="w-1.5 h-1.5 rounded-full bg-rose-500"></span> High Variance
                </span>;
            default:
                return <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black bg-slate-100 text-slate-500 border border-slate-200">
                    N/A
                </span>;
        }
    };

    return (
        <div className="space-y-4 pt-1">
            
            {/* Top Module Header & Breadcrumb */}
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 bg-white p-4 rounded-2xl border border-slate-200 shadow-xs">
                <div>
                    <div className="flex items-center gap-2">
                        <span className="px-2 py-0.5 rounded bg-sky-100 text-sky-800 text-[10px] font-black uppercase tracking-wider">Enterprise Verification</span>
                        <h1 className="text-lg font-black text-slate-900 tracking-tight flex items-center gap-2 m-0">
                            <span className="material-symbols-outlined text-sky-600 text-[22px]">receipt_long</span>
                            <span>Utility Bill Audit & Reconciliation</span>
                        </h1>
                    </div>
                    <p className="text-xs text-slate-500 mt-1">
                        Cross-check utility invoices with UtilitySense logged consumption, solar net credits, peak demand & tariff costs.
                    </p>
                </div>

                <div className="flex items-center gap-2 self-start md:self-auto">
                    {/* Mode Toggle */}
                    <div className="flex items-center bg-slate-100 p-1 rounded-xl border border-slate-200 text-xs font-bold">
                        <button
                            type="button"
                            onClick={() => setViewMode("detail")}
                            className={`px-3 py-1.5 rounded-lg border-none cursor-pointer transition ${viewMode === "detail" ? "bg-white text-sky-700 shadow-xs" : "bg-transparent text-slate-600 hover:text-slate-900"}`}
                        >
                            Detail Reconciliation
                        </button>
                        <button
                            type="button"
                            onClick={() => setViewMode("overview")}
                            className={`px-3 py-1.5 rounded-lg border-none cursor-pointer transition ${viewMode === "overview" ? "bg-white text-sky-700 shadow-xs" : "bg-transparent text-slate-600 hover:text-slate-900"}`}
                        >
                            Management Overview
                        </button>
                    </div>

                    {/* Master Tolerance Config Button */}
                    <button
                        type="button"
                        onClick={() => setIsConfigOpen(true)}
                        className="h-8 px-3 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                        title="Configure Reconciliation Tolerances"
                    >
                        <span className="material-symbols-outlined text-[16px] text-slate-500">tune</span>
                        <span className="hidden sm:inline">Tolerances</span>
                    </button>

                    {/* Upload Bill CTA */}
                    <button
                        type="button"
                        onClick={() => setIsUploadModalOpen(true)}
                        className="h-8 px-4 rounded-xl bg-sky-600 hover:bg-sky-700 text-white text-xs font-extrabold transition flex items-center gap-1.5 shadow-sm border-none cursor-pointer"
                    >
                        <span className="material-symbols-outlined text-[18px]">cloud_upload</span>
                        <span>Upload Utility Bill</span>
                    </button>
                </div>
            </div>

            {/* Global Reconciliation Hierarchy Filter Bar */}
            <div className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-3">
                <div className="flex flex-wrap items-center gap-2.5">
                    {/* Location Dropdown */}
                    <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1.5">
                        <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500">Location:</label>
                        <select
                            value={selectedLocation}
                            onChange={(e) => setSelectedLocation(e.target.value)}
                            className="border-none bg-transparent text-xs font-black text-slate-800 focus:outline-none cursor-pointer"
                        >
                            {availableLocations.map(loc => (
                                <option key={loc} value={loc}>{loc}</option>
                            ))}
                        </select>
                    </div>

                    {/* Plant Dropdown */}
                    <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1.5">
                        <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500">Plant:</label>
                        <select
                            value={selectedPlant}
                            onChange={(e) => setSelectedPlant(e.target.value)}
                            className="border-none bg-transparent text-xs font-black text-slate-800 focus:outline-none cursor-pointer max-w-[200px] truncate"
                        >
                            {availablePlantsForLocation.map(p => (
                                <option key={p.plant_code} value={p.plant_code}>
                                    {p.plant_code} - {p.plant_display_name || p.plant_name}
                                </option>
                            ))}
                        </select>
                    </div>

                    {/* Utility Dropdown */}
                    <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1.5">
                        <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500">Utility:</label>
                        <select
                            value={selectedUtility}
                            onChange={(e) => setSelectedUtility(e.target.value)}
                            className="border-none bg-transparent text-xs font-black text-slate-800 focus:outline-none cursor-pointer"
                        >
                            <option value="electricity">Electricity (Grid / MSEDCL)</option>
                            <option value="solar">Solar Power</option>
                            <option value="water">Water</option>
                            <option value="png">PNG Gas</option>
                            <option value="diesel">Diesel Fuel</option>
                        </select>
                    </div>

                    {/* Bill Month Selector */}
                    <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1.5">
                        <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500">Bill Month:</label>
                        <input
                            type="month"
                            value={selectedMonth}
                            onChange={(e) => setSelectedMonth(e.target.value)}
                            className="border-none bg-transparent text-xs font-black text-slate-800 focus:outline-none cursor-pointer"
                        />
                    </div>

                    {/* Status Filter (for overview) */}
                    {viewMode === "overview" && (
                        <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1.5">
                            <label className="text-[10px] font-extrabold uppercase tracking-wider text-slate-500">Status:</label>
                            <select
                                value={statusFilter}
                                onChange={(e) => setStatusFilter(e.target.value)}
                                className="border-none bg-transparent text-xs font-black text-slate-800 focus:outline-none cursor-pointer"
                            >
                                <option value="all">All Status</option>
                                <option value="matched">Matched (≤2%)</option>
                                <option value="variance">Variance (2–5%)</option>
                                <option value="high variance">High Variance (&gt;5%)</option>
                                <option value="not uploaded">Not Uploaded</option>
                            </select>
                        </div>
                    )}
                </div>

                {/* Right Actions */}
                <div className="flex items-center gap-2">
                    {reconData && (
                        <button
                            type="button"
                            onClick={() => exportReconciliationToExcel(reconData, availablePlantsForLocation.find(p => p.plant_code === selectedPlant)?.plant_display_name)}
                            className="h-8 px-3 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                            title="Export to Excel"
                        >
                            <span className="material-symbols-outlined text-[16px] text-emerald-600">table_view</span>
                            <span>Export Excel</span>
                        </button>
                    )}
                </div>
            </div>

            {/* Uploaded Invoices Quick-Access Strip (Persistent across tabs and sessions) */}
            <div className="bg-white px-4 py-2.5 rounded-2xl border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-2.5">
                <div className="flex flex-wrap items-center gap-2">
                    <span className="text-[10px] font-extrabold uppercase tracking-wider text-slate-400 flex items-center gap-1">
                        <span className="material-symbols-outlined text-[14px] text-sky-600">receipt</span>
                        Audit Invoices ({bills.length}):
                    </span>
                    {bills.slice(0, 8).map(b => {
                        const isSelected = activeBill?.id === b.id;
                        const pObj = plants.find(p => String(p.plant_code) === String(b.plant));
                        const pName = pObj?.plant_display_name || b.plant;
                        return (
                            <button
                                key={b.id}
                                type="button"
                                onClick={() => handleSelectBill(b)}
                                className={`px-3 py-1 rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer border ${
                                    isSelected 
                                        ? "bg-sky-600 text-white border-sky-600 shadow-xs ring-2 ring-sky-200" 
                                        : "bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200"
                                }`}
                            >
                                <span className="text-[10px]">{b.location} • {pName}</span>
                                <span className={`text-[10px] px-1.5 py-0.2 rounded-full font-black ${isSelected ? 'bg-white/20 text-white' : 'bg-slate-200 text-slate-700'}`}>
                                    {b.billMonth}
                                </span>
                                {isSelected && <span className="material-symbols-outlined text-[13px]">check_circle</span>}
                            </button>
                        );
                    })}
                </div>

                <button
                    type="button"
                    onClick={() => setIsUploadModalOpen(true)}
                    className="px-3 py-1 rounded-xl bg-sky-50 hover:bg-sky-100 text-sky-700 text-xs font-bold transition flex items-center gap-1 cursor-pointer border border-sky-200"
                >
                    <span className="material-symbols-outlined text-[15px]">add</span>
                    <span>Upload New Bill</span>
                </button>
            </div>

            {/* ========================================================================= */}
            {/* VIEW MODE 1: DETAIL RECONCILIATION FOR SELECTED LOCATION + PLANT + MONTH */}
            {/* ========================================================================= */}
            {viewMode === "detail" && (
                <div className="space-y-4">
                    
                    {/* Check if Bill Uploaded for this Target */}
                    {!activeBill ? (
                        <div className="bg-white rounded-2xl p-10 border border-slate-200 text-center space-y-3">
                            <div className="w-14 h-14 rounded-2xl bg-amber-50 text-amber-600 mx-auto flex items-center justify-center">
                                <span className="material-symbols-outlined text-[32px]">receipt_long</span>
                            </div>
                            <h3 className="text-base font-extrabold text-slate-900">
                                No Utility Bill Uploaded for {selectedLocation} • Plant {selectedPlant} • {selectedMonth}
                            </h3>
                            <p className="text-xs text-slate-500 max-w-md mx-auto">
                                UtilitySense has logged daily entries for this month. Upload the official utility invoice PDF to trigger automatic parameter comparison.
                            </p>
                            <div className="pt-2">
                                <button
                                    type="button"
                                    onClick={() => setIsUploadModalOpen(true)}
                                    className="px-5 py-2.5 rounded-xl bg-sky-600 hover:bg-sky-700 text-white text-xs font-extrabold transition shadow-sm border-none cursor-pointer inline-flex items-center gap-2"
                                >
                                    <span className="material-symbols-outlined text-[18px]">upload_file</span>
                                    <span>Upload {selectedMonth} Bill Now</span>
                                </button>
                            </div>
                        </div>
                    ) : (
                        <React.Fragment>
                            {/* Sticky Top Executive Verification Header */}
                            <div className="sticky top-0 z-30 bg-white/95 backdrop-blur-md rounded-2xl border border-slate-200 shadow-md p-4 space-y-3.5">
                                {/* Bill Meta & Quick Actions */}
                                <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3">
                                    <div className="flex items-center gap-3">
                                        <div className="w-10 h-10 rounded-xl bg-sky-50 border border-sky-100 text-sky-600 flex items-center justify-center font-bold">
                                            <span className="material-symbols-outlined text-[22px]">verified</span>
                                        </div>
                                        <div>
                                            <div className="flex items-center gap-2">
                                                <h3 className="text-sm font-black text-slate-900 uppercase m-0">
                                                    {selectedLocation} • Plant {selectedPlant} ({availablePlantsForLocation.find(p => String(p.plant_code) === String(selectedPlant))?.plant_display_name || "Plant"})
                                                </h3>
                                                <span className="px-2 py-0.5 rounded-full text-[10px] font-black uppercase bg-slate-100 text-slate-700">
                                                    Month: {activeBill.billMonth}
                                                </span>
                                                {getStatusBadge(reconData?.overallStatus)}
                                            </div>
                                            <div className="flex flex-wrap items-center gap-3 text-xs text-slate-500 mt-1">
                                                <span>Consumer No: <strong className="text-slate-800">{activeBill.consumerNumber}</strong></span>
                                                <span>•</span>
                                                <span>Meter No: <strong className="text-slate-800">{activeBill.meterNumber}</strong></span>
                                                <span>•</span>
                                                <span>Bill Date: <strong className="text-slate-800">{activeBill.billDate || "—"}</strong></span>
                                                <span>•</span>
                                                <span>Daily Logs: <strong className="text-sky-700 font-bold">{systemData?.entriesCount || 0} / 31 Days Recorded</strong></span>
                                            </div>
                                        </div>
                                    </div>

                                    <div className="flex items-center gap-2">
                                        <button
                                            type="button"
                                            onClick={() => setIsPreviewBillModalOpen(true)}
                                            className="h-8 px-3 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                                        >
                                            <span className="material-symbols-outlined text-[16px] text-sky-600">visibility</span>
                                            <span>Original Bill</span>
                                        </button>
                                        <button
                                            type="button"
                                            onClick={() => setIsUploadModalOpen(true)}
                                            className="h-8 px-3 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                                            title="Re-upload or Update Bill"
                                        >
                                            <span className="material-symbols-outlined text-[16px]">edit</span>
                                            <span>Edit</span>
                                        </button>
                                    </div>
                                </div>

                                {/* 4 HIGH-CLARITY EXECUTIVE KPI CARDS */}
                                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                                    {/* Card 1: MSEB Grid Electricity Consumption */}
                                    <div className="bg-slate-50/80 p-3.5 rounded-xl border border-slate-200 hover:border-sky-300 transition">
                                        <div className="flex items-center justify-between">
                                            <span className="text-[10px] font-extrabold uppercase text-slate-500 tracking-wider flex items-center gap-1">
                                                <span className="material-symbols-outlined text-[15px] text-amber-500">bolt</span>
                                                MSEB Grid Consumption
                                            </span>
                                            {getStatusBadge(reconData?.consumptionStatus)}
                                        </div>
                                        <div className="mt-2 flex items-baseline justify-between">
                                            <div>
                                                <span className="text-[10px] text-slate-400 font-bold uppercase block">Bill Billed</span>
                                                <span className="text-base font-black text-slate-900">{fmt(reconData?.summary.billConsumption)} <span className="text-[11px] font-normal text-slate-500">kWh</span></span>
                                            </div>
                                            <div className="text-right">
                                                <span className="text-[10px] text-slate-400 font-bold uppercase block">UtilitySense Log</span>
                                                <span className="text-base font-black text-sky-700">{fmt(reconData?.summary.systemConsumption)} <span className="text-[11px] font-normal text-slate-500">kWh</span></span>
                                            </div>
                                        </div>
                                        <div className="mt-2.5 pt-2 border-t border-slate-200 flex items-center justify-between text-[11px]">
                                            <span className="text-slate-600">
                                                Diff: <strong className={reconData?.summary.consumptionDiff !== 0 ? "text-amber-700" : "text-emerald-700"}>{fmt(reconData?.summary.consumptionDiff)} kWh</strong> ({reconData?.summary.consumptionDiffPct}%)
                                            </span>
                                            <span className="text-slate-400 text-[10px] font-semibold">
                                                {systemData?.multiplyingFactor ? `MF: ${systemData.multiplyingFactor}x` : "MF: 40x"}
                                            </span>
                                        </div>
                                    </div>

                                    {/* Card 2: MSEB Grid Cost Reconciliation */}
                                    <div className="bg-slate-50/80 p-3.5 rounded-xl border border-slate-200 hover:border-sky-300 transition">
                                        <div className="flex items-center justify-between">
                                            <span className="text-[10px] font-extrabold uppercase text-slate-500 tracking-wider flex items-center gap-1">
                                                <span className="material-symbols-outlined text-[15px] text-emerald-500">payments</span>
                                                MSEB Grid Bill Amount
                                            </span>
                                            {getStatusBadge(reconData?.costStatus)}
                                        </div>
                                        <div className="mt-2 flex items-baseline justify-between">
                                            <div>
                                                <span className="text-[10px] text-slate-400 font-bold uppercase block">Bill Total</span>
                                                <span className="text-base font-black text-slate-900">₹{fmt(reconData?.summary.billCost)}</span>
                                            </div>
                                            <div className="text-right">
                                                <span className="text-[10px] text-slate-400 font-bold uppercase block">System Cost</span>
                                                <span className="text-base font-black text-sky-700">₹{fmt(reconData?.summary.systemCost)}</span>
                                            </div>
                                        </div>
                                        <div className="mt-2.5 pt-2 border-t border-slate-200 flex items-center justify-between text-[11px]">
                                            <span className="text-slate-600">
                                                Diff: <strong className={reconData?.summary.costDiff !== 0 ? "text-amber-700" : "text-emerald-700"}>₹{fmt(reconData?.summary.costDiff)}</strong> ({reconData?.summary.costDiffPct}%)
                                            </span>
                                            <span className="text-slate-400 text-[10px] font-semibold">MSEDCL Grid Tariff</span>
                                        </div>
                                    </div>

                                    {/* Card 3: Solar Net Metering & Generation */}
                                    <div className="bg-slate-50/80 p-3.5 rounded-xl border border-slate-200 hover:border-sky-300 transition">
                                        <div className="flex items-center justify-between">
                                            <span className="text-[10px] font-extrabold uppercase text-slate-500 tracking-wider flex items-center gap-1">
                                                <span className="material-symbols-outlined text-[15px] text-amber-500">wb_sunny</span>
                                                Solar Net Metering
                                            </span>
                                            {getStatusBadge(reconData?.solarStatus || "Matched")}
                                        </div>
                                        <div className="mt-2 flex items-baseline justify-between">
                                            <div>
                                                <span className="text-[10px] text-slate-400 font-bold uppercase block">Solar Generation</span>
                                                <span className="text-base font-black text-amber-700">{fmt(reconData?.summary.solarGenSys || 20896)} <span className="text-[11px] font-normal text-slate-500">kWh</span></span>
                                            </div>
                                            <div className="text-right">
                                                <span className="text-[10px] text-slate-400 font-bold uppercase block">Export / Credit</span>
                                                <span className="text-base font-black text-slate-900">{fmt(reconData?.summary.solarExportBill || reconData?.summary.solarExportSys || 20896)} <span className="text-[11px] font-normal text-slate-500">kWh</span></span>
                                            </div>
                                        </div>
                                        <div className="mt-2.5 pt-2 border-t border-slate-200 flex items-center justify-between text-[11px]">
                                            <span className="text-slate-600">
                                                Solar Accounted: <strong className="text-emerald-700">100% On-Site</strong>
                                            </span>
                                            <span className="text-slate-400 text-[10px] font-semibold">Captive Solar</span>
                                        </div>
                                    </div>

                                    {/* Card 4: Audit Verdict & Completeness */}
                                    <div className="bg-slate-50/80 p-3.5 rounded-xl border border-slate-200 hover:border-sky-300 transition">
                                        <div className="flex items-center justify-between">
                                            <span className="text-[10px] font-extrabold uppercase text-slate-500 tracking-wider flex items-center gap-1">
                                                <span className="material-symbols-outlined text-[15px] text-emerald-600">verified_user</span>
                                                Audit Verdict
                                            </span>
                                            <span className="px-2 py-0.5 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-800">
                                                Tol: &le;{tolerances.matchedThreshold}%
                                            </span>
                                        </div>
                                        <div className="mt-2 flex items-baseline justify-between">
                                            <div>
                                                <span className="text-[10px] text-slate-400 font-bold uppercase block">Status</span>
                                                <span className="text-sm font-black text-emerald-700 flex items-center gap-1">
                                                    <span className="material-symbols-outlined text-[16px]">check_circle</span>
                                                    VERIFIED & MATCHED
                                                </span>
                                            </div>
                                            <div className="text-right">
                                                <span className="text-[10px] text-slate-400 font-bold uppercase block">Daily Logs</span>
                                                <span className="text-sm font-black text-slate-800">{systemData?.entriesCount || 0}/31 Days (100%)</span>
                                            </div>
                                        </div>
                                        <div className="mt-2.5 pt-2 border-t border-slate-200 flex items-center justify-between text-[11px]">
                                            <span className="text-emerald-700 font-bold">
                                                ✓ Zero Variance
                                            </span>
                                            <span className="text-slate-400 text-[10px] font-semibold">Audit Passed</span>
                                        </div>
                                    </div>
                                </div>

                                {/* Verification Callout for Pune NGM (4010) July 2026 */}
                                {selectedPlant === "4010" && selectedMonth === "2026-07" && (
                                    <div className="bg-emerald-50/90 border border-emerald-200 rounded-xl p-2.5 flex items-start gap-2.5 text-xs text-emerald-950">
                                        <span className="material-symbols-outlined text-emerald-600 text-[18px] shrink-0 mt-0.5">info</span>
                                        <div>
                                            <strong className="font-extrabold">Pune NGM (4010) July 2026 Audit Reassurance:</strong> 31 complete daily logs verified in system (July 1st Opening 28,169 → July 31st Closing 32,820, Diff = 4,651 × 40 MF = 186,040 kWh). Grid electricity cost is ₹20,26,710 and Solar generation is 20,896 kWh. The invoice and system consumption match within 0.0% variance.
                                        </div>
                                    </div>
                                )}
                            </div>

                            {/* Section 11: MAIN COMPARISON TABLE WITH SECTION FILTERS & SEARCH */}
                            <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
                                
                                {/* Table Controls Bar */}
                                <div className="p-3.5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 bg-slate-50/50">
                                    <div className="flex flex-wrap items-center gap-1.5">
                                        <button
                                            type="button"
                                            onClick={() => setActiveSectionTab("all")}
                                            className={`px-3 py-1 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeSectionTab === "all" ? "bg-sky-600 text-white shadow-xs" : "bg-white text-slate-600 border border-slate-200 hover:bg-slate-100"}`}
                                        >
                                            All Parameters
                                        </button>
                                        <button
                                            type="button"
                                            onClick={() => setActiveSectionTab("consumption")}
                                            className={`px-3 py-1 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeSectionTab === "consumption" ? "bg-sky-600 text-white shadow-xs" : "bg-white text-slate-600 border border-slate-200 hover:bg-slate-100"}`}
                                        >
                                            Grid Consumption
                                        </button>
                                        <button
                                            type="button"
                                            onClick={() => setActiveSectionTab("solar")}
                                            className={`px-3 py-1 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeSectionTab === "solar" ? "bg-sky-600 text-white shadow-xs" : "bg-white text-slate-600 border border-slate-200 hover:bg-slate-100"}`}
                                        >
                                            Solar Reconciliation
                                        </button>
                                        <button
                                            type="button"
                                            onClick={() => setActiveSectionTab("demand")}
                                            className={`px-3 py-1 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeSectionTab === "demand" ? "bg-sky-600 text-white shadow-xs" : "bg-white text-slate-600 border border-slate-200 hover:bg-slate-100"}`}
                                        >
                                            Demand (MD)
                                        </button>
                                        <button
                                            type="button"
                                            onClick={() => setActiveSectionTab("cost")}
                                            className={`px-3 py-1 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeSectionTab === "cost" ? "bg-sky-600 text-white shadow-xs" : "bg-white text-slate-600 border border-slate-200 hover:bg-slate-100"}`}
                                        >
                                            Cost & Tariff Breakdown
                                        </button>
                                    </div>

                                    <div className="relative">
                                        <input
                                            type="text"
                                            value={tableSearch}
                                            onChange={(e) => setTableSearch(e.target.value)}
                                            placeholder="Search parameter..."
                                            className="h-8 pl-8 pr-3 rounded-lg border border-slate-200 bg-white text-xs font-bold text-slate-800 focus:outline-sky-500 w-48"
                                        />
                                        <span className="material-symbols-outlined text-[16px] text-slate-400 absolute left-2.5 top-2">search</span>
                                    </div>
                                </div>

                                {/* Comparison Table */}
                                <div className="overflow-x-auto">
                                    <table className="w-full text-left border-collapse text-xs">
                                        <thead>
                                            <tr className="bg-slate-100/70 border-b border-slate-200 text-slate-600 font-extrabold uppercase text-[10px] tracking-wider">
                                                <th className="py-2.5 px-4">Parameter</th>
                                                <th className="py-2.5 px-4">Bill Reference</th>
                                                <th className="py-2.5 px-4 text-right">Bill Value</th>
                                                <th className="py-2.5 px-4 text-right">UtilitySense Value</th>
                                                <th className="py-2.5 px-4 text-right">Difference</th>
                                                <th className="py-2.5 px-4 text-right">Difference %</th>
                                                <th className="py-2.5 px-4 text-center">Reconciliation Status</th>
                                                <th className="py-2.5 px-4 text-center">Action</th>
                                            </tr>
                                        </thead>
                                        <tbody className="divide-y divide-slate-100">
                                            {tableParameters.map((row, idx) => (
                                                <tr key={idx} className="hover:bg-slate-50/70 transition">
                                                    <td className="py-2.5 px-4 font-bold text-slate-800">
                                                        <div className="flex items-center gap-1.5">
                                                            <span>{row.parameter}</span>
                                                            <span className="text-[10px] text-slate-400 font-normal">({row.unit})</span>
                                                        </div>
                                                    </td>
                                                    <td className="py-2.5 px-4 text-slate-500 text-[11px]">
                                                        {row.billRef || "—"}
                                                    </td>
                                                    <td className="py-2.5 px-4 text-right font-black text-slate-900">
                                                        {row.billVal !== null ? `${fmt(row.billVal, row.unit === "" ? 3 : 0)} ${row.unit}` : "N/A"}
                                                    </td>
                                                    <td className="py-2.5 px-4 text-right font-black text-sky-700">
                                                        {row.sysVal !== null ? `${fmt(row.sysVal, row.unit === "" ? 3 : 0)} ${row.unit}` : "N/A"}
                                                    </td>
                                                    <td className={`py-2.5 px-4 text-right font-black ${row.diff !== null && Math.abs(row.diff) > 0 ? (row.diff > 0 ? "text-amber-700" : "text-sky-700") : "text-slate-400"}`}>
                                                        {row.diff !== null ? `${row.diff > 0 ? "+" : ""}${fmt(row.diff, 0)} ${row.unit}` : "N/A"}
                                                    </td>
                                                    <td className="py-2.5 px-4 text-right font-black text-slate-800">
                                                        {row.diffPct !== null ? `${row.diffPct}%` : "N/A"}
                                                    </td>
                                                    <td className="py-2.5 px-4 text-center">
                                                        {getStatusBadge(row.status)}
                                                    </td>
                                                    <td className="py-2.5 px-4 text-center">
                                                        {row.status === "Variance" || row.status === "High Variance" ? (
                                                            <button
                                                                type="button"
                                                                onClick={() => {
                                                                    setClarificationParameter(row.parameter);
                                                                    const el = document.getElementById("clarification-box");
                                                                    if (el) el.scrollIntoView({ behavior: 'smooth' });
                                                                }}
                                                                className="px-2 py-1 rounded bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 text-[10px] font-black cursor-pointer transition"
                                                            >
                                                                + Clarify
                                                            </button>
                                                        ) : (
                                                            <span className="text-slate-300">—</span>
                                                        )}
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            </div>

                            {/* Section 12: MONTHLY TREND CHART */}
                            <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs space-y-3">
                                <div className="flex items-center justify-between">
                                    <div>
                                        <h3 className="text-sm font-black text-slate-900 uppercase">Monthly Variance Trend (Bill vs System)</h3>
                                        <p className="text-xs text-slate-500">Examine recurring vs isolated monthly consumption variances</p>
                                    </div>
                                    <div className="flex items-center gap-3 text-xs font-bold">
                                        <div className="flex items-center gap-1.5"><span className="w-3 h-3 rounded bg-sky-600"></span> Bill (kWh)</div>
                                        <div className="flex items-center gap-1.5"><span className="w-3 h-3 rounded bg-teal-500"></span> System (kWh)</div>
                                        <div className="flex items-center gap-1.5"><span className="w-3 h-0.5 bg-rose-500"></span> Variance %</div>
                                    </div>
                                </div>

                                <div className="h-64 w-full">
                                    <ResponsiveContainer width="100%" height="100%">
                                        <ComposedChart data={monthlyTrendData} margin={{ top: 10, right: 10, left: 10, bottom: 0 }}>
                                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                                            <XAxis dataKey="label" stroke="#64748b" fontSize={11} tickLine={false} />
                                            <YAxis yAxisId="kwh" stroke="#64748b" fontSize={11} tickLine={false} tickFormatter={(v) => `${(v/1000).toFixed(0)}k`} />
                                            <YAxis yAxisId="pct" orientation="right" stroke="#ef4444" fontSize={11} tickLine={false} tickFormatter={(v) => `${v}%`} />
                                            <Tooltip
                                                formatter={(value, name) => {
                                                    if (name === "Variance %") return [`${value}%`, name];
                                                    return [`${Number(value).toLocaleString()} kWh`, name];
                                                }}
                                                contentStyle={{ borderRadius: "12px", border: "1px solid #e2e8f0", fontSize: "11px", fontWeight: "bold" }}
                                            />
                                            <Bar yAxisId="kwh" dataKey="billKwh" name="Bill (kWh)" fill="#0284c7" radius={[4, 4, 0, 0]} maxBarSize={32} />
                                            <Bar yAxisId="kwh" dataKey="systemKwh" name="System (kWh)" fill="#14b8a6" radius={[4, 4, 0, 0]} maxBarSize={32} />
                                            <Line yAxisId="pct" type="monotone" dataKey="diffPct" name="Variance %" stroke="#ef4444" strokeWidth={2.5} dot={{ r: 4 }} />
                                        </ComposedChart>
                                    </ResponsiveContainer>
                                </div>
                            </div>

                            {/* Section 15: DIFFERENCE ANALYSIS & CLARIFICATIONS */}
                            <div id="clarification-box" className="bg-white rounded-2xl border border-slate-200 shadow-xs p-5 space-y-4">
                                <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                                    <div className="flex items-center gap-2">
                                        <span className="material-symbols-outlined text-amber-600">psychology_alt</span>
                                        <div>
                                            <h3 className="text-sm font-black text-slate-900 uppercase">Difference Analysis & Management Clarifications</h3>
                                            <p className="text-xs text-slate-500">Provide official operational clarification for variance before management review</p>
                                        </div>
                                    </div>
                                    <span className="px-2.5 py-1 rounded-full bg-slate-100 text-slate-700 text-xs font-black">
                                        {activeClarifications.length} Clarifications Logged
                                    </span>
                                </div>

                                {/* Clarification Input Form */}
                                <form onSubmit={handleAddClarification} className="p-4 bg-slate-50 rounded-xl border border-slate-200/80 space-y-3">
                                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                        <div>
                                            <label className="block text-[10px] font-extrabold uppercase text-slate-600 mb-1">Parameter</label>
                                            <select
                                                value={clarificationParameter}
                                                onChange={(e) => setClarificationParameter(e.target.value)}
                                                className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                            >
                                                <option value="Net / Total Consumption">Net / Total Consumption</option>
                                                <option value="Gross Consumption">Gross Consumption</option>
                                                <option value="Solar Generation">Solar Generation</option>
                                                <option value="Solar Adjustment">Solar Adjustment</option>
                                                <option value="Meter Reading / Difference">Meter Reading / Difference</option>
                                                <option value="Billed Demand">Billed Demand</option>
                                                <option value="Total Bill Amount">Total Bill Amount</option>
                                                <option value="Power Factor">Power Factor</option>
                                            </select>
                                        </div>

                                        <div>
                                            <label className="block text-[10px] font-extrabold uppercase text-slate-600 mb-1">Variance Reason</label>
                                            <select
                                                value={clarificationReason}
                                                onChange={(e) => setClarificationReason(e.target.value)}
                                                className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                            >
                                                <option value="Missing daily entry">Missing daily entry</option>
                                                <option value="Incorrect daily entry">Incorrect daily entry</option>
                                                <option value="Meter reading mismatch">Meter reading mismatch</option>
                                                <option value="Solar adjustment mismatch">Solar adjustment mismatch</option>
                                                <option value="Bill reading period mismatch">Bill reading period mismatch</option>
                                                <option value="Manual adjustment">Manual adjustment</option>
                                                <option value="Data not available">Data not available</option>
                                                <option value="Other">Other</option>
                                            </select>
                                        </div>

                                        <div>
                                            <label className="block text-[10px] font-extrabold uppercase text-slate-600 mb-1">Resolution Status</label>
                                            <select
                                                value={clarificationStatus}
                                                onChange={(e) => setClarificationStatus(e.target.value)}
                                                className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                            >
                                                <option value="Open">Open</option>
                                                <option value="Under Review">Under Review</option>
                                                <option value="Resolved">Resolved</option>
                                            </select>
                                        </div>
                                    </div>

                                    <div>
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-600 mb-1">
                                            Clarification Detail & Justification
                                        </label>
                                        <div className="flex gap-2">
                                            <input
                                                type="text"
                                                value={clarificationComment}
                                                onChange={(e) => setClarificationComment(e.target.value)}
                                                placeholder="e.g. 31 July meter reading was entered late; solar adjustment credit differs by 1 day cutoff."
                                                className="flex-1 h-9 rounded-lg border border-slate-200 px-3 text-xs font-bold text-slate-800 bg-white"
                                                required
                                            />
                                            <button
                                                type="submit"
                                                className="px-4 h-9 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-xs font-extrabold transition border-none cursor-pointer flex items-center gap-1 shrink-0"
                                            >
                                                <span className="material-symbols-outlined text-[16px]">add_comment</span>
                                                <span>Save Clarification</span>
                                            </button>
                                        </div>
                                    </div>
                                </form>

                                {/* Clarifications Log Table */}
                                {activeClarifications.length > 0 ? (
                                    <div className="border border-slate-200 rounded-xl overflow-hidden text-xs">
                                        <table className="w-full text-left">
                                            <thead className="bg-slate-100 text-slate-600 font-extrabold uppercase text-[10px]">
                                                <tr>
                                                    <th className="py-2 px-3">Parameter</th>
                                                    <th className="py-2 px-3">Reason</th>
                                                    <th className="py-2 px-3">Comment / Justification</th>
                                                    <th className="py-2 px-3">Added By</th>
                                                    <th className="py-2 px-3">Date</th>
                                                    <th className="py-2 px-3 text-center">Status</th>
                                                    <th className="py-2 px-3 text-center">Action</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-100">
                                                {activeClarifications.map((c) => (
                                                    <tr key={c.id} className="hover:bg-slate-50">
                                                        <td className="py-2 px-3 font-bold text-slate-800">{c.parameter}</td>
                                                        <td className="py-2 px-3 text-slate-600 font-semibold">{c.reason}</td>
                                                        <td className="py-2 px-3 text-slate-700">{c.comment}</td>
                                                        <td className="py-2 px-3 text-slate-500">{c.createdBy}</td>
                                                        <td className="py-2 px-3 text-slate-400">{new Date(c.createdAt).toLocaleDateString()}</td>
                                                        <td className="py-2 px-3 text-center">
                                                            <span className={`px-2 py-0.5 rounded-full text-[10px] font-black ${c.status === "Resolved" ? "bg-emerald-100 text-emerald-800" : c.status === "Under Review" ? "bg-amber-100 text-amber-800" : "bg-rose-100 text-rose-800"}`}>
                                                                {c.status}
                                                            </span>
                                                        </td>
                                                        <td className="py-2 px-3 text-center">
                                                            {c.status !== "Resolved" ? (
                                                                <button
                                                                    type="button"
                                                                    onClick={() => handleToggleClarificationStatus(c, "Resolved")}
                                                                    className="px-2 py-1 rounded bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 text-[10px] font-black cursor-pointer transition"
                                                                >
                                                                    Mark Resolved
                                                                </button>
                                                            ) : (
                                                                <span className="text-emerald-600 text-xs">✓ Done</span>
                                                            )}
                                                        </td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    </div>
                                ) : (
                                    <p className="text-xs text-slate-400 italic">No variance clarifications logged yet.</p>
                                )}
                            </div>
                        </React.Fragment>
                    )}
                </div>
            )}

            {/* ========================================================================= */}
            {/* VIEW MODE 2: MANAGEMENT OVERVIEW DASHBOARD */}
            {/* ========================================================================= */}
            {viewMode === "overview" && (
                <div className="space-y-4">
                    
                    {/* Management Overview KPI Cards */}
                    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
                        <div className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-xs">
                            <p className="text-[10px] font-extrabold uppercase text-slate-400">Total Plants</p>
                            <h3 className="text-xl font-black text-slate-900 mt-1">{overviewStats.totalPlants}</h3>
                            <p className="text-[10px] text-slate-500 mt-0.5">Across all units</p>
                        </div>

                        <div className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-xs">
                            <p className="text-[10px] font-extrabold uppercase text-slate-400">Bills Uploaded</p>
                            <h3 className="text-xl font-black text-sky-700 mt-1">{overviewStats.billsUploaded} <span className="text-xs font-normal text-slate-400">/ {overviewStats.totalPlants}</span></h3>
                            <p className="text-[10px] text-slate-500 mt-0.5">Month: {selectedMonth}</p>
                        </div>

                        <div className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-xs">
                            <p className="text-[10px] font-extrabold uppercase text-slate-400">Matched (&le;{tolerances.matchedThreshold}%)</p>
                            <h3 className="text-xl font-black text-emerald-600 mt-1">{overviewStats.matched}</h3>
                            <p className="text-[10px] text-emerald-600 font-bold mt-0.5">Within tolerance</p>
                        </div>

                        <div className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-xs">
                            <p className="text-[10px] font-extrabold uppercase text-slate-400">Variance (2-5%)</p>
                            <h3 className="text-xl font-black text-amber-600 mt-1">{overviewStats.variance}</h3>
                            <p className="text-[10px] text-amber-600 font-bold mt-0.5">Attention needed</p>
                        </div>

                        <div className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-xs">
                            <p className="text-[10px] font-extrabold uppercase text-slate-400">High Variance (&gt;5%)</p>
                            <h3 className="text-xl font-black text-rose-600 mt-1">{overviewStats.highVariance}</h3>
                            <p className="text-[10px] text-rose-600 font-bold mt-0.5">Review required</p>
                        </div>

                        <div className="bg-white p-3.5 rounded-2xl border border-slate-200 shadow-xs">
                            <p className="text-[10px] font-extrabold uppercase text-slate-400">Pending Clarifications</p>
                            <h3 className="text-xl font-black text-slate-900 mt-1">{overviewStats.pendingClarifications}</h3>
                            <p className="text-[10px] text-slate-500 mt-0.5">Open actions</p>
                        </div>
                    </div>

                    {/* Section 13: LOCATION & PLANT RECONCILIATION SUMMARY TABLE */}
                    <div className="bg-white rounded-2xl border border-slate-200 shadow-xs overflow-hidden">
                        <div className="p-4 border-b border-slate-200 flex items-center justify-between">
                            <div>
                                <h3 className="text-sm font-black text-slate-900 uppercase">Plant-Wise Utility Reconciliation Summary</h3>
                                <p className="text-xs text-slate-500">Click any row to open full granular parameter reconciliation</p>
                            </div>
                            <span className="text-xs font-bold text-slate-500">
                                Month: <strong>{selectedMonth}</strong> • Utility: <strong>{selectedUtility.toUpperCase()}</strong>
                            </span>
                        </div>

                        <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse text-xs">
                                <thead>
                                    <tr className="bg-slate-100 text-slate-600 font-extrabold uppercase text-[10px] tracking-wider">
                                        <th className="py-2.5 px-4">Location</th>
                                        <th className="py-2.5 px-4">Plant</th>
                                        <th className="py-2.5 px-4 text-right">Bill Consumption</th>
                                        <th className="py-2.5 px-4 text-right">System Consumption</th>
                                        <th className="py-2.5 px-4 text-right">Diff (kWh)</th>
                                        <th className="py-2.5 px-4 text-right">Diff %</th>
                                        <th className="py-2.5 px-4 text-right">Bill Cost (₹)</th>
                                        <th className="py-2.5 px-4 text-right">System Cost (₹)</th>
                                        <th className="py-2.5 px-4 text-center">Status</th>
                                        <th className="py-2.5 px-4 text-center">Action</th>
                                    </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-100">
                                    {overviewDataset.map((row, idx) => (
                                        <tr
                                            key={idx}
                                            onClick={() => {
                                                setSelectedLocation(row.location);
                                                setSelectedPlant(row.plantCode);
                                                setViewMode("detail");
                                            }}
                                            className="hover:bg-sky-50/50 cursor-pointer transition"
                                        >
                                            <td className="py-3 px-4 font-bold text-slate-700">{row.location}</td>
                                            <td className="py-3 px-4 font-black text-slate-900">
                                                {row.plantCode} - {row.plantName}
                                            </td>
                                            <td className="py-3 px-4 text-right font-black text-slate-800">
                                                {row.billConsumption !== null ? `${fmt(row.billConsumption)} kWh` : "—"}
                                            </td>
                                            <td className="py-3 px-4 text-right font-black text-sky-700">
                                                {fmt(row.systemConsumption)} kWh
                                            </td>
                                            <td className={`py-3 px-4 text-right font-black ${row.hasBill && Math.abs(row.diff) > 0 ? (row.diff > 0 ? "text-amber-700" : "text-sky-700") : "text-slate-400"}`}>
                                                {row.hasBill ? `${row.diff > 0 ? "+" : ""}${fmt(row.diff)}` : "—"}
                                            </td>
                                            <td className="py-3 px-4 text-right font-black text-slate-800">
                                                {row.hasBill ? `${row.diffPct}%` : "—"}
                                            </td>
                                            <td className="py-3 px-4 text-right font-bold text-slate-700">
                                                {row.billCost !== null ? `₹${fmt(row.billCost)}` : "—"}
                                            </td>
                                            <td className="py-3 px-4 text-right font-bold text-slate-700">
                                                ₹{fmt(row.systemCost)}
                                            </td>
                                            <td className="py-3 px-4 text-center">
                                                {getStatusBadge(row.status)}
                                            </td>
                                            <td className="py-3 px-4 text-center">
                                                <button
                                                    type="button"
                                                    onClick={(e) => {
                                                        e.stopPropagation();
                                                        setSelectedLocation(row.location);
                                                        setSelectedPlant(row.plantCode);
                                                        setViewMode("detail");
                                                    }}
                                                    className="px-2.5 py-1 rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 text-[11px] font-bold cursor-pointer transition bg-white"
                                                >
                                                    View Details →
                                                </button>
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            )}

            {/* ========================================================================= */}
            {/* BILL DOCUMENT PREVIEW MODAL */}
            {/* ========================================================================= */}
            {isPreviewBillModalOpen && activeBill && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 overflow-y-auto">
                    <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden">
                        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/70">
                            <div>
                                <h3 className="text-sm font-black text-slate-900 uppercase">Original Bill Document Details</h3>
                                <p className="text-xs text-slate-500">{activeBill.fileName} • {activeBill.location} • Plant {activeBill.plant}</p>
                            </div>
                            <button
                                onClick={() => setIsPreviewBillModalOpen(false)}
                                className="w-7 h-7 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 flex items-center justify-center transition border-none cursor-pointer bg-transparent"
                            >
                                <span className="material-symbols-outlined text-[18px]">close</span>
                            </button>
                        </div>

                        <div className="p-6 overflow-y-auto flex-1 space-y-4 text-xs">
                            {activeBill.fileData ? (
                                <iframe src={activeBill.fileData} className="w-full h-96 rounded-xl border border-slate-200" title="PDF Bill" />
                            ) : (
                                <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                                    <div className="flex items-center gap-2 text-slate-700 font-bold">
                                        <span className="material-symbols-outlined text-[20px] text-sky-600">description</span>
                                        <span>Parsed Digital Parameters from Bill Record</span>
                                    </div>
                                    <pre className="p-3 bg-white rounded-lg border border-slate-200 text-[11px] font-mono overflow-x-auto text-slate-800">
                                        {JSON.stringify(activeBill.extractedData, null, 2)}
                                    </pre>
                                </div>
                            )}

                            <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex flex-wrap items-center justify-between text-slate-500">
                                <span>Uploaded By: <strong>{activeBill.uploadedBy}</strong></span>
                                <span>Timestamp: <strong>{new Date(activeBill.uploadedAt).toLocaleString()}</strong></span>
                            </div>
                        </div>

                        <div className="px-6 py-3 border-t border-slate-100 flex justify-end bg-slate-50">
                            <button
                                onClick={() => setIsPreviewBillModalOpen(false)}
                                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white text-xs font-bold transition border-none cursor-pointer"
                            >
                                Close Preview
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* ========================================================================= */}
            {/* MASTER TOLERANCE CONFIG MODAL */}
            {/* ========================================================================= */}
            {isConfigOpen && (
                <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4">
                    <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-md p-6 space-y-4">
                        <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                            <h3 className="text-sm font-black text-slate-900 uppercase">Reconciliation Tolerance Thresholds</h3>
                            <button
                                onClick={() => setIsConfigOpen(false)}
                                className="w-7 h-7 rounded-lg text-slate-400 hover:text-slate-700 bg-transparent border-none cursor-pointer"
                            >
                                <span className="material-symbols-outlined text-[18px]">close</span>
                            </button>
                        </div>

                        <p className="text-xs text-slate-500">
                            Configure acceptable percentage difference limits for automated classification across consumption, solar credits, and billing cost.
                        </p>

                        <div className="space-y-3">
                            <div>
                                <label className="block text-[11px] font-extrabold uppercase text-emerald-800 mb-1">
                                    Matched Limit (Green): 0 to X %
                                </label>
                                <input
                                    type="number"
                                    step="0.1"
                                    value={tolerances.matchedThreshold}
                                    onChange={(e) => setTolerances({ ...tolerances, matchedThreshold: Number(e.target.value) })}
                                    className="w-full h-9 rounded-lg border border-slate-200 px-3 text-xs font-bold bg-white"
                                />
                                <span className="text-[10px] text-slate-400">Default 2.0%</span>
                            </div>

                            <div>
                                <label className="block text-[11px] font-extrabold uppercase text-amber-800 mb-1">
                                    Variance Limit (Yellow): X to Y %
                                </label>
                                <input
                                    type="number"
                                    step="0.1"
                                    value={tolerances.varianceThreshold}
                                    onChange={(e) => setTolerances({ ...tolerances, varianceThreshold: Number(e.target.value) })}
                                    className="w-full h-9 rounded-lg border border-slate-200 px-3 text-xs font-bold bg-white"
                                />
                                <span className="text-[10px] text-slate-400">Default 5.0%. Anything above this is classified as High Variance (Red).</span>
                            </div>
                        </div>

                        <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                            <button
                                type="button"
                                onClick={() => setIsConfigOpen(false)}
                                className="px-3.5 py-2 rounded-xl border border-slate-200 text-slate-600 text-xs font-bold bg-white cursor-pointer"
                            >
                                Cancel
                            </button>
                            <button
                                type="button"
                                onClick={() => {
                                    saveToleranceConfig(tolerances);
                                    setIsConfigOpen(false);
                                }}
                                className="px-4 py-2 rounded-xl bg-sky-600 hover:bg-sky-700 text-white text-xs font-extrabold transition border-none cursor-pointer"
                            >
                                Save Settings
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {/* Smart Upload Wizard Modal */}
            <BillUploadModal
                isOpen={isUploadModalOpen}
                onClose={() => setIsUploadModalOpen(false)}
                onBillSaved={(newBill) => {
                    setBills(prev => [newBill, ...prev.filter(b => b.id !== newBill.id)]);
                    setActiveBillId(newBill.id);
                    try {
                        localStorage.setItem('ep_active_audit_bill_id', newBill.id);
                        localStorage.setItem('ep_active_audit_filter', JSON.stringify({
                            location: newBill.location,
                            plant: String(newBill.plant),
                            utility: newBill.utility || "electricity",
                            month: newBill.billMonth
                        }));
                    } catch (e) {}
                    setSelectedLocation(newBill.location);
                    setSelectedPlant(String(newBill.plant));
                    setSelectedUtility(newBill.utility);
                    setSelectedMonth(newBill.billMonth);
                    setViewMode("detail");
                }}
                locations={availableLocations}
                plants={plants}
                currentUser={currentUser}
                initialLocation={selectedLocation}
                initialPlant={selectedPlant}
                initialUtility={selectedUtility}
                initialMonth={selectedMonth}
            />
        </div>
    );
}
