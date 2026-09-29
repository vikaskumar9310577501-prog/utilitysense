// UtilitySense Utility Bill Reconciliation Service
// Manages bill storage, system aggregation, tolerance checks, variance analysis, and audit trails

import { supabase } from '../../supabaseClient';
import * as XLSX from 'xlsx';

const STORAGE_KEY_BILLS = 'ep_utility_bills_v1';
const STORAGE_KEY_CLARIFICATIONS = 'ep_reconciliation_clarifications_v1';
const STORAGE_KEY_TOLERANCES = 'ep_reconciliation_tolerances_v1';

// Default Configurable Tolerances (Master Config)
export const DEFAULT_TOLERANCE_CONFIG = {
    matchedThreshold: 2.0,     // 0 - 2% = Matched (Green)
    varianceThreshold: 5.0,    // 2 - 5% = Variance (Yellow)
    // Above 5% = High Variance (Red)
};

export function getToleranceConfig() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY_TOLERANCES);
        if (saved) return JSON.parse(saved);
    } catch (e) {
        console.error("Failed to load tolerances", e);
    }
    return { ...DEFAULT_TOLERANCE_CONFIG };
}

export function saveToleranceConfig(config) {
    try {
        localStorage.setItem(STORAGE_KEY_TOLERANCES, JSON.stringify(config));
    } catch (e) {
        console.error("Failed to save tolerances", e);
    }
}

/**
 * Calculates Difference and Percentage safely
 */
export function calculateVariance(billVal, sysVal, toleranceConfig = getToleranceConfig()) {
    if (billVal === null || billVal === undefined || isNaN(billVal)) {
        return { diff: null, diffPct: null, status: "N/A" };
    }
    if (sysVal === null || sysVal === undefined || isNaN(sysVal)) {
        return { diff: null, diffPct: null, status: "N/A" };
    }

    const b = Number(billVal);
    const s = Number(sysVal);
    const diff = b - s;
    const diffPct = Math.abs(b) > 0 ? (Math.abs(diff) / Math.abs(b)) * 100 : 0;

    let status = "Matched";
    if (diffPct > toleranceConfig.varianceThreshold) {
        status = "High Variance";
    } else if (diffPct > toleranceConfig.matchedThreshold) {
        status = "Variance";
    }

    return {
        diff: Number(diff.toFixed(2)),
        diffPct: Number(diffPct.toFixed(2)),
        status
    };
}

/**
 * Aggregates UtilitySense daily entries strictly by Location + Plant + Month
 * Zero mixing between locations, plants, or months!
 */
export function aggregateSystemConsumption(dailyEntries, location, plantCode, monthStr) {
    if (!dailyEntries || !Array.isArray(dailyEntries) || !monthStr) {
        return null;
    }

    const locClean = String(location || "").trim().toUpperCase();
    const plantClean = String(plantCode || "").trim().toUpperCase();

    // Filter strictly by Location, Plant and Month (YYYY-MM)
    const matchedEntries = dailyEntries.filter(e => {
        if (!e.date || !e.date.startsWith(monthStr)) return false;
        
        // Match plant code or plant name
        const ePlant = String(e.plant || "").trim().toUpperCase();
        const matchesPlant = ePlant === plantClean;
        if (!matchesPlant) return false;

        // Match location if entry has location specified
        if (locClean && e.location) {
            const eLoc = String(e.location || "").trim().toUpperCase();
            if (eLoc !== locClean) return false;
        }

        return true;
    });

    let totalElectricityKwh = 0;
    let totalElectricityCost = 0;
    let totalSolarGenKwh = 0;
    let totalSolarUtilKwh = 0;
    let totalSolarCost = 0;
    let totalCost = 0;
    let maxRecordedDemand = 0;
    let minOpening = null;
    let maxClosing = null;

    matchedEntries.forEach(e => {
        const kwh = Number(e.electricity_consumption) || 0;
        const eCost = Number(e.electricity_cost) || 0;
        const sGen = Number(e.solar_generated) || 0;
        const sUtil = Number(e.solar_utilized) || 0;
        const sCost = Number(e.solar_cost) || 0;
        const tCost = Number(e.total_cost) || (eCost + sCost);

        totalElectricityKwh += kwh;
        totalElectricityCost += eCost;
        totalSolarGenKwh += sGen;
        totalSolarUtilKwh += sUtil;
        totalSolarCost += sCost;
        totalCost += tCost;

        if (e.electricity_opening !== null && e.electricity_opening !== undefined) {
            const op = Number(e.electricity_opening);
            if (minOpening === null || op < minOpening) minOpening = op;
        }
        if (e.electricity_closing !== null && e.electricity_closing !== undefined) {
            const cl = Number(e.electricity_closing);
            if (maxClosing === null || cl > maxClosing) maxClosing = cl;
        }

        // Daily diff estimation for peak demand
        if (kwh > maxRecordedDemand) {
            maxRecordedDemand = kwh;
        }
    });

    return {
        entriesCount: matchedEntries.length,
        grossConsumptionKwh: totalElectricityKwh,
        netConsumptionKwh: totalElectricityKwh, // In UtilitySense, electricity_consumption is net grid import
        electricityCost: totalElectricityCost,
        solarGenKwh: totalSolarGenKwh,
        solarUtilKwh: totalSolarUtilKwh,
        solarExportKwh: totalSolarGenKwh > totalSolarUtilKwh ? (totalSolarGenKwh - totalSolarUtilKwh) : 0,
        solarAdjustmentKwh: totalSolarUtilKwh,
        solarCost: totalSolarCost,
        totalCost: totalCost,
        openingMeter: minOpening,
        closingMeter: maxClosing,
        meterDifference: maxClosing !== null && minOpening !== null ? Math.max(0, maxClosing - minOpening) : null,
        // Demand & PF are not tracked daily as single scalars, so return null rather than fake values
        recordedDemandKva: null,
        billedDemandKva: null,
        contractDemandKva: null,
        powerFactor: null
    };
}

/**
 * Builds Full Reconciliation Comparison between a Bill and UtilitySense System Data
 */
export function buildReconciliationData(bill, systemData, toleranceConfig = getToleranceConfig()) {
    if (!bill) return null;

    const b = bill.extractedData || {};
    const s = systemData || {};

    // 1. Grid Electricity Consumption
    const grossComp = calculateVariance(b.grossUnitsKwh, s.grossConsumptionKwh, toleranceConfig);
    const netComp = calculateVariance(b.billedUnitsKwh, s.netConsumptionKwh, toleranceConfig);
    const kvahComp = calculateVariance(b.billedUnitsKvah, null, toleranceConfig); // System doesn't track KVAh directly -> N/A
    const openMeterComp = calculateVariance(b.openingMeter, s.openingMeter, toleranceConfig);
    const closeMeterComp = calculateVariance(b.closingMeter, s.closingMeter, toleranceConfig);
    const meterDiffComp = calculateVariance(b.meterDifference, s.meterDifference, toleranceConfig);
    const mfComp = calculateVariance(b.multiplyingFactor, null, toleranceConfig);
    const rkLagComp = calculateVariance(b.rkvahLag, null, toleranceConfig);
    const rkLeadComp = calculateVariance(b.rkvahLead, null, toleranceConfig);

    const consumptionSection = [
        { parameter: "Gross Consumption", unit: "kWh", billVal: b.grossUnitsKwh, sysVal: s.grossConsumptionKwh, ...grossComp, billRef: "MSEDCL Gross Units" },
        { parameter: "Net / Total Consumption", unit: "kWh", billVal: b.billedUnitsKwh, sysVal: s.netConsumptionKwh, ...netComp, billRef: "MSEDCL Billed kWh" },
        { parameter: "Billed Units (kVAh)", unit: "kVAh", billVal: b.billedUnitsKvah, sysVal: null, ...kvahComp, billRef: "KVAH Reading" },
        { parameter: "Previous Meter Reading", unit: "kWh", billVal: b.openingMeter, sysVal: s.openingMeter, ...openMeterComp, billRef: "Prev Reading" },
        { parameter: "Current Meter Reading", unit: "kWh", billVal: b.closingMeter, sysVal: s.closingMeter, ...closeMeterComp, billRef: "Curr Reading" },
        { parameter: "Meter Difference", unit: "kWh", billVal: b.meterDifference, sysVal: s.meterDifference, ...meterDiffComp, billRef: "Closing - Opening" },
        { parameter: "Multiplying Factor", unit: "x", billVal: b.multiplyingFactor, sysVal: null, ...mfComp, billRef: "Meter Multiplying Factor" },
        { parameter: "RKVAH Lag", unit: "RKVAh", billVal: b.rkvahLag, sysVal: null, ...rkLagComp, billRef: "Lag Reactive Energy" },
        { parameter: "RKVAH Lead", unit: "RKVAh", billVal: b.rkvahLead, sysVal: null, ...rkLeadComp, billRef: "Lead Reactive Energy" }
    ];

    // 2. Solar Reconciliation Section
    const solGenComp = calculateVariance(b.solarGenUnits, s.solarGenKwh, toleranceConfig);
    const solExpComp = calculateVariance(b.solarExportUnits, s.solarExportKwh, toleranceConfig);
    const solAdjComp = calculateVariance(b.solarAdjUnits, s.solarAdjustmentKwh, toleranceConfig);
    const solCapComp = calculateVariance(b.solarCapacity, null, toleranceConfig);

    const solarSection = [
        { parameter: "Solar Generation", unit: "kWh", billVal: b.solarGenUnits, sysVal: s.solarGenKwh, ...solGenComp, billRef: "Solar Gen Meter Total" },
        { parameter: "Solar Export", unit: "kWh", billVal: b.solarExportUnits, sysVal: s.solarExportKwh, ...solExpComp, billRef: "Solar Net Export Units" },
        { parameter: "Solar Adjustment", unit: "kWh", billVal: b.solarAdjUnits, sysVal: s.solarAdjustmentKwh, ...solAdjComp, billRef: "Adjustment-Solar (Credit)" },
        { parameter: "Solar Sanctioned Capacity", unit: "kWp", billVal: b.solarCapacity, sysVal: null, ...solCapComp, billRef: "Sanctioned Solar Capacity" }
    ];

    // 3. Demand Reconciliation Section
    const recMdComp = calculateVariance(b.recordedDemandKva, null, toleranceConfig);
    const billMdComp = calculateVariance(b.billedDemandKva, null, toleranceConfig);
    const cdComp = calculateVariance(b.contractDemandKva, null, toleranceConfig);
    const kwMdComp = calculateVariance(b.kwMaxDemand, null, toleranceConfig);

    const demandSection = [
        { parameter: "Recorded Maximum Demand", unit: "kVA", billVal: b.recordedDemandKva, sysVal: null, ...recMdComp, billRef: "Recorded MD" },
        { parameter: "Billed Demand", unit: "kVA", billVal: b.billedDemandKva, sysVal: null, ...billMdComp, billRef: "Billed Demand (KVA)" },
        { parameter: "Contract Demand", unit: "kVA", billVal: b.contractDemandKva, sysVal: null, ...cdComp, billRef: "Contract Demand (KVA)" },
        { parameter: "KW Maximum Demand", unit: "kW", billVal: b.kwMaxDemand, sysVal: null, ...kwMdComp, billRef: "KW MD" }
    ];

    // 4. Power Factor Reconciliation Section
    const pfComp = calculateVariance(b.powerFactor, null, toleranceConfig);
    const powerFactorSection = [
        { parameter: "Power Factor (PF)", unit: "", billVal: b.powerFactor, sysVal: null, ...pfComp, billRef: "Average Monthly Power Factor" }
    ];

    // 5. Cost Reconciliation Section
    const nrgCostComp = calculateVariance(b.energyCharges, s.electricityCost, toleranceConfig);
    const totCostComp = calculateVariance(b.totalBillAmount, s.totalCost, toleranceConfig);
    const curBillComp = calculateVariance(b.currentBillAmount, s.totalCost, toleranceConfig);
    const demChgComp = calculateVariance(b.demandCharges, null, toleranceConfig);
    const excDemChgComp = calculateVariance(b.excessDemandCharges, null, toleranceConfig);
    const whlChgComp = calculateVariance(b.wheelingCharges, null, toleranceConfig);
    const facChgComp = calculateVariance(b.facCharges, null, toleranceConfig);
    const edChgComp = calculateVariance(b.electricityDuty, null, toleranceConfig);
    const todChgComp = calculateVariance(b.todCharges, null, toleranceConfig);
    const gsChgComp = calculateVariance(b.gridSupportCharges, null, toleranceConfig);
    const ppdComp = calculateVariance(b.promptPaymentDiscount, null, toleranceConfig);
    const subComp = calculateVariance(b.subsidiesTotal, null, toleranceConfig);

    const costSection = [
        { parameter: "Energy Charges", unit: "₹", billVal: b.energyCharges, sysVal: s.electricityCost, ...nrgCostComp, billRef: "Base Energy Charges" },
        { parameter: "Demand Charges", unit: "₹", billVal: b.demandCharges, sysVal: null, ...demChgComp, billRef: "Fixed Demand Charges" },
        { parameter: "Excess Demand Charges", unit: "₹", billVal: b.excessDemandCharges, sysVal: null, ...excDemChgComp, billRef: "Penalty for Demand Breach" },
        { parameter: "Wheeling Charges", unit: "₹", billVal: b.wheelingCharges, sysVal: null, ...whlChgComp, billRef: "Wheeling Charge @ Rs/U" },
        { parameter: "FAC (Fuel Adjustment Charge)", unit: "₹", billVal: b.facCharges, sysVal: null, ...facChgComp, billRef: "Fuel Adjustment Surcharge" },
        { parameter: "Electricity Duty", unit: "₹", billVal: b.electricityDuty, sysVal: null, ...edChgComp, billRef: "Govt Electricity Duty" },
        { parameter: "TOD Tariff Charges", unit: "₹", billVal: b.todCharges, sysVal: null, ...todChgComp, billRef: "Time of Day Surcharge/Rebate" },
        { parameter: "Grid Support Charges", unit: "₹", billVal: b.gridSupportCharges, sysVal: null, ...gsChgComp, billRef: "Grid Support Charge (Solar)" },
        { parameter: "Prompt Payment Discount (PPD)", unit: "₹", billVal: b.promptPaymentDiscount, sysVal: null, ...ppdComp, billRef: "PPD Rebate" },
        { parameter: "Govt Subsidy / Rebates", unit: "₹", billVal: b.subsidiesTotal, sysVal: null, ...subComp, billRef: "State Govt Subsidies" },
        { parameter: "Current Bill Amount", unit: "₹", billVal: b.currentBillAmount, sysVal: s.totalCost, ...curBillComp, billRef: "Total Current Bill" },
        { parameter: "Net Payable Bill Amount", unit: "₹", billVal: b.totalBillAmount, sysVal: s.totalCost, ...totCostComp, billRef: "Total Payable upto Due Date" }
    ];

    // Compute Overall Status
    const primaryComps = [netComp, grossComp, solGenComp, totCostComp].filter(c => c.status !== "N/A");
    let overallStatus = "Matched";
    if (primaryComps.some(c => c.status === "High Variance")) {
        overallStatus = "High Variance";
    } else if (primaryComps.some(c => c.status === "Variance")) {
        overallStatus = "Variance";
    }

    return {
        billId: bill.id,
        location: bill.location,
        plant: bill.plant,
        utility: bill.utility,
        billMonth: bill.billMonth,
        overallStatus,
        consumptionStatus: netComp.status !== "N/A" ? netComp.status : grossComp.status,
        solarStatus: solGenComp.status !== "N/A" ? solGenComp.status : "N/A",
        demandStatus: billMdComp.status,
        powerFactorStatus: pfComp.status,
        costStatus: totCostComp.status !== "N/A" ? totCostComp.status : curBillComp.status,
        summary: {
            billConsumption: b.billedUnitsKwh ?? b.grossUnitsKwh ?? 0,
            systemConsumption: s.netConsumptionKwh ?? 0,
            consumptionDiff: netComp.diff ?? 0,
            consumptionDiffPct: netComp.diffPct ?? 0,
            
            billCost: b.totalBillAmount ?? b.currentBillAmount ?? 0,
            systemCost: s.totalCost ?? 0,
            costDiff: totCostComp.diff ?? 0,
            costDiffPct: totCostComp.diffPct ?? 0,

            solarGenBill: b.solarGenUnits ?? 0,
            solarGenSys: s.solarGenKwh ?? 0,
            solarGenDiff: solGenComp.diff ?? 0,
            solarExportBill: b.solarExportUnits ?? b.solarAdjUnits ?? 0,
            solarExportSys: s.solarExportKwh ?? 0
        },
        sections: {
            consumption: consumptionSection,
            solar: solarSection,
            demand: demandSection,
            powerFactor: powerFactorSection,
            cost: costSection
        }
    };
}

/**
 * Storage & Database APIs
 */
export async function loadAllBills() {
    let localBills = [];
    try {
        const raw = localStorage.getItem(STORAGE_KEY_BILLS);
        if (raw) localBills = JSON.parse(raw);
    } catch (e) {
        console.error("Failed to read local bills", e);
    }

    // Try Supabase
    try {
        const { data, error } = await supabase.from('utility_bills').select('*');
        if (!error && data && data.length > 0) {
            // Merge with local storage
            const map = new Map();
            localBills.forEach(b => map.set(b.id, b));
            data.forEach(dbItem => {
                map.set(dbItem.id, {
                    id: dbItem.id,
                    location: dbItem.location_id,
                    plant: dbItem.plant_id,
                    utility: dbItem.utility_id || 'electricity',
                    billMonth: dbItem.bill_month,
                    billDate: dbItem.bill_date,
                    consumerNumber: dbItem.consumer_number,
                    meterNumber: dbItem.meter_number,
                    fileName: dbItem.file_name,
                    fileData: dbItem.file_data,
                    status: dbItem.status,
                    notes: dbItem.notes,
                    uploadedBy: dbItem.uploaded_by,
                    uploadedAt: dbItem.uploaded_at,
                    extractedData: dbItem.notes ? JSON.parse(dbItem.notes) : {}
                });
            });
            const merged = Array.from(map.values());
            localStorage.setItem(STORAGE_KEY_BILLS, JSON.stringify(merged));
            return merged;
        }
    } catch (err) {
        // Fallback to local
    }

    // If empty, seed initial sample bills for demo so management sees live data immediately
    if (localBills.length === 0) {
        const seedBills = getInitialSeedBills();
        localStorage.setItem(STORAGE_KEY_BILLS, JSON.stringify(seedBills));
        return seedBills;
    }

    return localBills;
}

export async function saveBill(billRecord) {
    let bills = [];
    try {
        const raw = localStorage.getItem(STORAGE_KEY_BILLS);
        if (raw) bills = JSON.parse(raw);
    } catch (e) {}

    const existingIdx = bills.findIndex(b => b.id === billRecord.id || 
        (b.location === billRecord.location && b.plant === billRecord.plant && b.utility === billRecord.utility && b.billMonth === billRecord.billMonth));
    
    if (existingIdx >= 0) {
        bills[existingIdx] = { ...bills[existingIdx], ...billRecord };
    } else {
        bills.unshift(billRecord);
    }

    localStorage.setItem(STORAGE_KEY_BILLS, JSON.stringify(bills));

    // Supabase upsert attempt
    try {
        const payload = {
            id: billRecord.id,
            location_id: billRecord.location,
            plant_id: billRecord.plant,
            utility_id: billRecord.utility,
            bill_month: billRecord.billMonth,
            bill_date: billRecord.billDate || null,
            consumer_number: billRecord.consumerNumber || null,
            meter_number: billRecord.meterNumber || null,
            file_name: billRecord.fileName || null,
            file_data: billRecord.fileData || null,
            status: billRecord.status || 'Matched',
            notes: JSON.stringify(billRecord.extractedData || {}),
            uploaded_by: billRecord.uploadedBy || 'IT Admin',
            uploaded_at: billRecord.uploadedAt || new Date().toISOString(),
            last_modified_by: billRecord.lastModifiedBy || billRecord.uploadedBy || 'IT Admin',
            last_modified_at: new Date().toISOString()
        };
        await supabase.from('utility_bills').upsert(payload);
    } catch (e) {
        console.warn("Supabase utility_bills upsert skipped / unmigrated:", e.message);
    }

    return bills;
}

export async function deleteBill(billId) {
    let bills = [];
    try {
        const raw = localStorage.getItem(STORAGE_KEY_BILLS);
        if (raw) bills = JSON.parse(raw);
    } catch (e) {}

    bills = bills.filter(b => b.id !== billId);
    localStorage.setItem(STORAGE_KEY_BILLS, JSON.stringify(bills));

    try {
        await supabase.from('utility_bills').delete().eq('id', billId);
    } catch (e) {}

    return bills;
}

/**
 * Clarifications Management
 */
export async function loadClarifications() {
    let local = [];
    try {
        const raw = localStorage.getItem(STORAGE_KEY_CLARIFICATIONS);
        if (raw) local = JSON.parse(raw);
    } catch (e) {}

    try {
        const { data, error } = await supabase.from('reconciliation_clarifications').select('*');
        if (!error && data && data.length > 0) {
            const map = new Map();
            local.forEach(c => map.set(c.id, c));
            data.forEach(d => {
                map.set(d.id, {
                    id: d.id,
                    billId: d.bill_id,
                    parameter: d.parameter,
                    reason: d.reason,
                    comment: d.comment,
                    status: d.status,
                    createdBy: d.created_by,
                    createdAt: d.created_at,
                    resolvedBy: d.resolved_by,
                    resolvedAt: d.resolved_at
                });
            });
            const merged = Array.from(map.values());
            localStorage.setItem(STORAGE_KEY_CLARIFICATIONS, JSON.stringify(merged));
            return merged;
        }
    } catch (e) {}

    return local;
}

export async function saveClarification(item) {
    let items = [];
    try {
        const raw = localStorage.getItem(STORAGE_KEY_CLARIFICATIONS);
        if (raw) items = JSON.parse(raw);
    } catch (e) {}

    const idx = items.findIndex(c => c.id === item.id);
    if (idx >= 0) {
        items[idx] = { ...items[idx], ...item };
    } else {
        items.unshift(item);
    }

    localStorage.setItem(STORAGE_KEY_CLARIFICATIONS, JSON.stringify(items));

    try {
        await supabase.from('reconciliation_clarifications').upsert({
            id: item.id,
            bill_id: item.billId,
            parameter: item.parameter,
            reason: item.reason,
            comment: item.comment,
            status: item.status || 'Open',
            created_by: item.createdBy,
            created_at: item.createdAt || new Date().toISOString(),
            resolved_by: item.resolvedBy || null,
            resolved_at: item.resolvedAt || null
        });
    } catch (e) {}

    return items;
}

/**
 * Excel Export for Reconciliation Report
 */
export function exportReconciliationToExcel(reconData, plantName = "") {
    if (!reconData) return;

    const wb = XLSX.utils.book_new();

    // 1. Summary Sheet
    const summaryData = [
        ["UTILITY BILL RECONCILIATION REPORT"],
        ["Location", reconData.location],
        ["Plant", `${reconData.plant} ${plantName ? '- ' + plantName : ''}`],
        ["Utility", reconData.utility],
        ["Bill Month", reconData.billMonth],
        ["Reconciliation Status", reconData.overallStatus],
        [],
        ["SUMMARY METRICS", "BILL VALUE", "UTILITYSENSE VALUE", "VARIANCE", "VARIANCE %", "STATUS"],
        [
            "Grid / Net Consumption (kWh)",
            reconData.summary.billConsumption,
            reconData.summary.systemConsumption,
            reconData.summary.consumptionDiff,
            `${reconData.summary.consumptionDiffPct}%`,
            reconData.consumptionStatus
        ],
        [
            "Total Cost (₹)",
            reconData.summary.billCost,
            reconData.summary.systemCost,
            reconData.summary.costDiff,
            `${reconData.summary.costDiffPct}%`,
            reconData.costStatus
        ],
        [
            "Solar Generation (kWh)",
            reconData.summary.solarGenBill,
            reconData.summary.solarGenSys,
            reconData.summary.solarGenDiff,
            "—",
            reconData.solarStatus
        ]
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, wsSummary, "Executive Summary");

    // 2. Detailed Parameters Sheet
    const allRows = [
        ["Category", "Parameter", "Unit", "Bill Value", "UtilitySense Value", "Difference", "Difference %", "Status", "Bill Reference"]
    ];

    Object.entries(reconData.sections).forEach(([secKey, secItems]) => {
        secItems.forEach(row => {
            allRows.push([
                secKey.toUpperCase(),
                row.parameter,
                row.unit,
                row.billVal !== null ? row.billVal : "N/A",
                row.sysVal !== null ? row.sysVal : "N/A",
                row.diff !== null ? row.diff : "N/A",
                row.diffPct !== null ? `${row.diffPct}%` : "N/A",
                row.status,
                row.billRef || ""
            ]);
        });
    });

    const wsDetails = XLSX.utils.aoa_to_sheet(allRows);
    XLSX.utils.book_append_sheet(wb, wsDetails, "Detailed Comparison");

    // Save File
    const fname = `Reconciliation_${reconData.location}_${reconData.plant}_${reconData.billMonth}.xlsx`;
    XLSX.writeFile(wb, fname);
}

/**
 * Seed initial sample bills for demo with realistic MSEDCL industrial bill data
 */
function getInitialSeedBills() {
    return [
        {
            id: "bill_demo_4010_2026_07",
            location: "PUNE",
            plant: "4010",
            utility: "electricity",
            billMonth: "2026-07",
            billDate: "2026-08-05",
            consumerNumber: "170019014520",
            meterNumber: "LT-8849120",
            status: "Matched",
            notes: "Verified against MSEDCL monthly bill copy",
            uploadedBy: "Vikas Sharma (Energy Mgr)",
            uploadedAt: "2026-08-06T10:30:00.000Z",
            fileName: "MSEDCL_Bill_4010_July_2026.pdf",
            fileData: null,
            extractedData: {
                billMonth: "2026-07",
                billMonthDisplay: "JUL 2026",
                startDate: "2026-07-01",
                endDate: "2026-07-31",
                billDate: "2026-08-05",
                consumerNo: "170019014520",
                consumerName: "NEXT GENERATION MANUFACTURERS",
                meterNo: "LT-8849120",
                tariffCategory: "HT-1 (A) Continuous",
                contractDemandKva: 1250,
                billedDemandKva: 980,
                recordedDemandKva: 940,
                kwMaxDemand: 890,
                powerFactor: 0.995,
                billedUnitsKvah: 128450,
                billedUnitsKwh: 124800,
                grossUnitsKwh: 145910,
                openingMeter: 48920,
                closingMeter: 51338,
                meterDifference: 2418,
                multiplyingFactor: 60,
                rkvahLag: 4200,
                rkvahLead: 1100,
                solarGenUnits: 21110,
                solarExportUnits: 2401,
                solarAdjUnits: 2401,
                solarCapacity: 250,
                energyCharges: 1048320,
                demandCharges: 480200,
                excessDemandCharges: 0,
                wheelingCharges: 147264,
                facCharges: 68640,
                electricityDuty: 115200,
                todCharges: -28400,
                gridSupportCharges: 18450,
                promptPaymentDiscount: 18420,
                subsidiesTotal: 0,
                currentBillAmount: 1831254,
                totalBillAmount: 1812834
            }
        },
        {
            id: "bill_demo_2020_2026_07",
            location: "PUNE",
            plant: "2020",
            utility: "electricity",
            billMonth: "2026-07",
            billDate: "2026-08-06",
            consumerNumber: "170019028811",
            meterNumber: "HT-992144",
            status: "Variance",
            notes: "Pending review for solar adjustment difference",
            uploadedBy: "IT Admin",
            uploadedAt: "2026-08-07T14:15:00.000Z",
            fileName: "MSEDCL_PGTL_2020_JUL_2026.pdf",
            fileData: null,
            extractedData: {
                billMonth: "2026-07",
                billMonthDisplay: "JUL 2026",
                startDate: "2026-07-01",
                endDate: "2026-07-31",
                billDate: "2026-08-06",
                consumerNo: "170019028811",
                consumerName: "PG TECHNOPLAST LIMITED",
                meterNo: "HT-992144",
                tariffCategory: "HT-1 (B)",
                contractDemandKva: 1800,
                billedDemandKva: 1420,
                recordedDemandKva: 1395,
                kwMaxDemand: 1320,
                powerFactor: 0.992,
                billedUnitsKvah: 198400,
                billedUnitsKwh: 192500,
                grossUnitsKwh: 226820,
                openingMeter: 75357,
                closingMeter: 81027,
                meterDifference: 5670,
                multiplyingFactor: 40,
                rkvahLag: 7200,
                rkvahLead: 1800,
                solarGenUnits: 34320,
                solarExportUnits: 3850,
                solarAdjUnits: 3850,
                solarCapacity: 400,
                energyCharges: 1617000,
                demandCharges: 695800,
                excessDemandCharges: 0,
                wheelingCharges: 227150,
                facCharges: 105875,
                electricityDuty: 177100,
                todCharges: -34000,
                gridSupportCharges: 24800,
                promptPaymentDiscount: 27800,
                subsidiesTotal: 0,
                currentBillAmount: 2790925,
                totalBillAmount: 2763125
            }
        },
        {
            id: "bill_demo_4010_2026_06",
            location: "PUNE",
            plant: "4010",
            utility: "electricity",
            billMonth: "2026-06",
            billDate: "2026-07-05",
            consumerNumber: "170019014520",
            meterNumber: "LT-8849120",
            status: "Matched",
            notes: "June verified and reconciled",
            uploadedBy: "Vikas Sharma (Energy Mgr)",
            uploadedAt: "2026-07-06T11:00:00.000Z",
            fileName: "MSEDCL_Bill_4010_June_2026.pdf",
            fileData: null,
            extractedData: {
                billMonth: "2026-06",
                billMonthDisplay: "JUN 2026",
                startDate: "2026-06-01",
                endDate: "2026-06-30",
                billDate: "2026-07-05",
                consumerNo: "170019014520",
                consumerName: "NEXT GENERATION MANUFACTURERS",
                meterNo: "LT-8849120",
                tariffCategory: "HT-1 (A) Continuous",
                contractDemandKva: 1250,
                billedDemandKva: 950,
                recordedDemandKva: 910,
                kwMaxDemand: 865,
                powerFactor: 0.996,
                billedUnitsKvah: 122100,
                billedUnitsKwh: 118500,
                grossUnitsKwh: 139800,
                openingMeter: 46612,
                closingMeter: 48920,
                meterDifference: 2308,
                multiplyingFactor: 60,
                rkvahLag: 3900,
                rkvahLead: 950,
                solarGenUnits: 21300,
                solarExportUnits: 2500,
                solarAdjUnits: 2500,
                solarCapacity: 250,
                energyCharges: 995400,
                demandCharges: 465500,
                excessDemandCharges: 0,
                wheelingCharges: 139830,
                facCharges: 65175,
                electricityDuty: 109350,
                todCharges: -26500,
                gridSupportCharges: 17500,
                promptPaymentDiscount: 17400,
                subsidiesTotal: 0,
                currentBillAmount: 1748855,
                totalBillAmount: 1731455
            }
        }
    ];
}
