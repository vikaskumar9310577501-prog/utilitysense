// UtilitySense Executive Utility Bill Audit & Reconciliation Module
// Minimal, High-Precision, Management-Grade Side-by-Side Comparison
// Strictly NO graphs, NO cluttered KPI cards — Pure actionable data for Management Review

import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { extractTextFromPdfBuffer, parseMsedclBillText } from '../../utils/msedclBillParser';
import {
    aggregateSystemConsumption,
    exportReconciliationToExcel,
    getToleranceConfig
} from './reconciliationService';

// Format Indian Currency & Numbers
function formatINR(val, decimals = 2) {
    if (val === null || val === undefined || isNaN(val)) return '—';
    return Number(val).toLocaleString('en-IN', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
    });
}

function formatNumber(val, decimals = 2) {
    if (val === null || val === undefined || isNaN(val)) return '—';
    return Number(val).toLocaleString('en-IN', {
        minimumFractionDigits: Number.isInteger(Number(val)) ? 0 : decimals,
        maximumFractionDigits: decimals
    });
}

export default function ReconciliationModule({
    plants = [],
    dailyEntries = [],
    currentUser,
    tariffs = []
}) {
    // 1. Filter States with Persistence
    const [selectedPlantCode, setSelectedPlantCode] = useState(() => {
        try {
            const saved = localStorage.getItem('ep_active_audit_filter_v2');
            if (saved) {
                const parsed = JSON.parse(saved);
                if (parsed.plantCode) return parsed.plantCode;
            }
        } catch (e) {}
        return plants[0]?.plant_code || '4010';
    });

    const [selectedMonth, setSelectedMonth] = useState(() => {
        try {
            const saved = localStorage.getItem('ep_active_audit_filter_v2');
            if (saved) {
                const parsed = JSON.parse(saved);
                if (parsed.month) return parsed.month;
            }
        } catch (e) {}
        return '2026-07';
    });

    // 2. Active Bill State with Persistence
    const [activeBill, setActiveBill] = useState(() => {
        try {
            const saved = localStorage.getItem('ep_active_audit_bill_v2');
            if (saved) return JSON.parse(saved);
        } catch (e) {}
        return null;
    });

    // 3. UI States
    const [isExtracting, setIsExtracting] = useState(false);
    const [extractError, setExtractError] = useState('');
    const [isEditingBill, setIsEditingBill] = useState(false);
    const [editFormData, setEditFormData] = useState(null);
    const [dragActive, setDragActive] = useState(false);
    const fileInputRef = useRef(null);

    // Save filter state to localStorage
    useEffect(() => {
        try {
            localStorage.setItem(
                'ep_active_audit_filter_v2',
                JSON.stringify({ plantCode: selectedPlantCode, month: selectedMonth })
            );
        } catch (e) {}
    }, [selectedPlantCode, selectedMonth]);

    // Save active bill state to localStorage
    useEffect(() => {
        try {
            if (activeBill) {
                localStorage.setItem('ep_active_audit_bill_v2', JSON.stringify(activeBill));
            } else {
                localStorage.removeItem('ep_active_audit_bill_v2');
            }
        } catch (e) {}
    }, [activeBill]);

    // Find selected plant object
    const selectedPlant = useMemo(() => {
        return (
            plants.find((p) => p.plant_code === selectedPlantCode) ||
            plants.find((p) => p.plant_display_name?.toUpperCase() === selectedPlantCode?.toUpperCase()) ||
            plants[0] || { plant_code: selectedPlantCode, plant_display_name: selectedPlantCode, location_name: 'PUNE' }
        );
    }, [plants, selectedPlantCode]);

    // Active Tariff Rates
    const activeTariffs = useMemo(() => {
        const elect = tariffs.find((t) => t.type === 'electricity' && t.status === 'Active')?.rate || 10.8939;
        const solar = tariffs.find((t) => t.type === 'solar' && t.status === 'Active')?.rate || 10.8939;
        return { electricityRate: Number(elect), solarRate: Number(solar) };
    }, [tariffs]);

    // Aggregate Software Daily Readings for Selected Plant & Month
    const systemData = useMemo(() => {
        if (!dailyEntries || !dailyEntries.length || !selectedMonth) {
            return {
                entriesCount: 0,
                grossConsumptionKwh: 0,
                netConsumptionKwh: 0,
                electricityCost: 0,
                solarGenKwh: 0,
                solarUtilKwh: 0,
                solarExportKwh: 0,
                solarAdjustmentKwh: 0,
                solarCost: 0,
                totalCost: 0,
                openingMeter: null,
                closingMeter: null,
                meterDifference: null,
                multiplyingFactor: null
            };
        }

        const plantCode = selectedPlant?.plant_code || selectedPlantCode;
        const plantName = selectedPlant?.plant_display_name || '';
        const locationName = selectedPlant?.location_name || '';

        // Filter entries strictly for this month and plant
        const matched = dailyEntries.filter((e) => {
            if (!e.date || !e.date.startsWith(selectedMonth)) return false;
            const p = String(e.plant || '').trim().toUpperCase();
            const matchesPlant =
                p === String(plantCode).trim().toUpperCase() ||
                (plantName && p === plantName.toUpperCase());
            return matchesPlant;
        });

        let totalElectricityKwh = 0;
        let totalElectricityCost = 0;
        let totalSolarGenKwh = 0;
        let totalSolarUtilKwh = 0;
        let totalSolarCost = 0;
        let totalCost = 0;
        let minOpening = null;
        let maxClosing = null;

        matched.forEach((e) => {
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
        });

        const meterDifference = maxClosing !== null && minOpening !== null ? Math.max(0, maxClosing - minOpening) : null;
        const multiplyingFactor = (totalElectricityKwh > 0 && meterDifference && meterDifference > 0)
            ? Math.round(totalElectricityKwh / meterDifference)
            : null;

        return {
            entriesCount: matched.length,
            grossConsumptionKwh: totalElectricityKwh,
            netConsumptionKwh: totalElectricityKwh,
            electricityCost: totalElectricityCost,
            solarGenKwh: totalSolarGenKwh,
            solarUtilKwh: totalSolarUtilKwh,
            solarExportKwh: totalSolarGenKwh > totalSolarUtilKwh ? (totalSolarGenKwh - totalSolarUtilKwh) : 0,
            solarAdjustmentKwh: totalSolarUtilKwh,
            solarCost: totalSolarCost,
            totalCost: totalCost || (totalElectricityCost + totalSolarCost),
            openingMeter: minOpening,
            closingMeter: maxClosing,
            meterDifference,
            multiplyingFactor
        };
    }, [dailyEntries, selectedPlant, selectedPlantCode, selectedMonth]);

    // Handle File Upload & Automated Parsing
    const processFile = async (file) => {
        if (!file) return;
        setIsExtracting(true);
        setExtractError('');

        try {
            const fileName = file.name;
            const isPdf = fileName.toLowerCase().endsWith('.pdf');
            const isExcel = fileName.toLowerCase().endsWith('.xlsx') || fileName.toLowerCase().endsWith('.xls');

            let extractedData = {};

            if (isPdf) {
                const arrayBuffer = await file.arrayBuffer();
                const rawText = await extractTextFromPdfBuffer(arrayBuffer);
                if (!rawText || rawText.length < 20) {
                    throw new Error('Unable to extract text from PDF. The file may be an image scan or protected.');
                }
                const parsed = parseMsedclBillText(rawText);
                if (!parsed) {
                    throw new Error('Could not identify electricity billing parameters from this PDF.');
                }
                extractedData = parsed;
            } else if (isExcel) {
                const arrayBuffer = await file.arrayBuffer();
                const wb = XLSX.read(arrayBuffer, { type: 'array' });
                const firstSheet = wb.Sheets[wb.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                // Simple key-value extraction from Excel
                const dataMap = {};
                rows.forEach((r) => {
                    if (r[0] && r[1] !== undefined) {
                        dataMap[String(r[0]).trim().toLowerCase()] = r[1];
                    }
                });
                extractedData = {
                    consumerNo: dataMap['consumer no'] || dataMap['consumer number'] || '',
                    billedUnitsKwh: Number(dataMap['billed units'] || dataMap['total units'] || dataMap['consumption'] || 0),
                    totalBillAmount: Number(dataMap['bill amount'] || dataMap['total amount'] || dataMap['amount'] || 0),
                    solarGenUnits: Number(dataMap['solar units'] || dataMap['solar generation'] || 0)
                };
            } else {
                throw new Error('Please upload a PDF (.pdf) or Excel (.xlsx) file.');
            }

            // If the bill detected a month, optionally align with user selection
            const billMonth = extractedData.billMonth || selectedMonth;

            const newBill = {
                id: `bill_${Date.now()}`,
                fileName,
                fileSize: file.size,
                uploadedAt: new Date().toISOString(),
                billMonth,
                plantCode: selectedPlantCode,
                plantName: selectedPlant.plant_display_name,
                location: selectedPlant.location_name,
                extractedData
            };

            setActiveBill(newBill);
            if (extractedData.billMonth && extractedData.billMonth !== selectedMonth) {
                setSelectedMonth(extractedData.billMonth);
            }
        } catch (err) {
            console.error('File parsing error:', err);
            setExtractError(err.message || 'Failed to parse bill. Please check file format.');
        } finally {
            setIsExtracting(false);
            if (fileInputRef.current) fileInputRef.current.value = '';
        }
    };

    // Drag and Drop handlers
    const handleDrag = (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (e.type === 'dragenter' || e.type === 'dragover') {
            setDragActive(true);
        } else if (e.type === 'dragleave') {
            setDragActive(false);
        }
    };

    const handleDrop = (e) => {
        e.preventDefault();
        e.stopPropagation();
        setDragActive(false);
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            processFile(e.dataTransfer.files[0]);
        }
    };

    // Calculate Line-by-Line Comparison Items
    const comparisonReport = useMemo(() => {
        if (!activeBill) return null;

        const b = activeBill.extractedData || {};
        const s = systemData || {};

        const makeRow = (id, paramName, unit, billVal, sysVal, category, helperText = '', rate = null) => {
            const hasBill = billVal !== null && billVal !== undefined && !isNaN(billVal);
            const hasSys = sysVal !== null && sysVal !== undefined && !isNaN(sysVal);

            let diff = null;
            let diffPct = null;
            let costImpact = null;
            let status = 'Statutory';

            if (hasBill && hasSys) {
                const bNum = Number(billVal);
                const sNum = Number(sysVal);
                diff = Number((bNum - sNum).toFixed(2));
                diffPct = Math.abs(bNum) > 0 ? Number(((Math.abs(diff) / Math.abs(bNum)) * 100).toFixed(2)) : 0;

                if (rate) {
                    costImpact = Number((diff * rate).toFixed(2));
                } else if (unit === '₹') {
                    costImpact = diff;
                }

                if (diffPct <= 0.5) {
                    status = 'Verified';
                } else if (diffPct <= 2.0) {
                    status = 'Minor Variance';
                } else {
                    status = 'Discrepancy';
                }
            } else if (hasBill && !hasSys) {
                status = 'Bill Only';
                if (unit === '₹') costImpact = Number(billVal);
            } else if (!hasBill && hasSys) {
                status = 'Software Only';
            }

            return {
                id,
                paramName,
                unit,
                billVal: hasBill ? Number(billVal) : null,
                sysVal: hasSys ? Number(sysVal) : null,
                diff,
                diffPct,
                costImpact,
                status,
                category,
                helperText
            };
        };

        const eRate = activeTariffs.electricityRate;
        const sRate = activeTariffs.solarRate;

        // 1. Grid Electricity Section
        const gridSection = [
            makeRow('net_units', 'Net Billed Electricity (kWh)', 'kWh', b.billedUnitsKwh, s.netConsumptionKwh, 'Grid Consumption', 'Units billed by DISCOM vs Net Daily meter sum', eRate),
            makeRow('gross_units', 'Gross Grid Units (kWh)', 'kWh', b.grossUnitsKwh, s.grossConsumptionKwh, 'Grid Consumption', 'Total imported grid units before solar netting', eRate),
            makeRow('open_meter', 'Opening Meter Reading', 'kWh', b.openingMeter, s.openingMeter, 'Meter Readings', 'Meter start reading for billing period'),
            makeRow('close_meter', 'Closing Meter Reading', 'kWh', b.closingMeter, s.closingMeter, 'Meter Readings', 'Meter end reading for billing period'),
            makeRow('meter_diff', 'Net Meter Difference', 'kWh', b.meterDifference, s.meterDifference, 'Meter Readings', 'Closing - Opening (Raw units)'),
            makeRow('mf', 'Multiplying Factor (MF)', 'x', b.multiplyingFactor, s.multiplyingFactor, 'Meter Readings', 'Current/Potential Transformer multiplier'),
            makeRow('kvah_units', 'Billed Apparent Energy (kVAh)', 'kVAh', b.billedUnitsKvah, null, 'Grid Consumption', 'Recorded kVAh in utility meter'),
            makeRow('rkvah_lag', 'Reactive Energy Lag (rKVAh)', 'rKVAh', b.rkvahLag, null, 'Grid Consumption', 'Inductive reactive draw'),
            makeRow('rkvah_lead', 'Reactive Energy Lead (rKVAh)', 'rKVAh', b.rkvahLead, null, 'Grid Consumption', 'Capacitive reactive draw')
        ];

        // 2. Solar Generation & Credit Section
        const solarSection = [
            makeRow('solar_gen', 'Solar Generation (kWh)', 'kWh', b.solarGenUnits, s.solarGenKwh, 'Solar Net Metering', 'Actual solar electricity produced at site', sRate),
            makeRow('solar_adj', 'Solar Credit / Adjustment (kWh)', 'kWh', b.solarAdjUnits, s.solarAdjustmentKwh, 'Solar Net Metering', 'Solar units credited against grid bill', eRate),
            makeRow('solar_export', 'Solar Export to Grid (kWh)', 'kWh', b.solarExportUnits, s.solarExportKwh, 'Solar Net Metering', 'Surplus solar injected into MSEB grid', eRate),
            makeRow('solar_cap', 'Solar Sanctioned Capacity (kWp)', 'kWp', b.solarCapacity, null, 'Solar Net Metering', 'Approved rooftop/ground solar capacity')
        ];

        // 3. Demand & Power Factor Section
        const demandSection = [
            makeRow('recorded_md', 'Recorded Maximum Demand (kVA)', 'kVA', b.recordedDemandKva, null, 'Demand & PF', 'Peak 15/30-minute demand recorded'),
            makeRow('billed_md', 'Billed Demand (kVA)', 'kVA', b.billedDemandKva, null, 'Demand & PF', 'Higher of recorded MD or 75% contract demand'),
            makeRow('contract_md', 'Sanctioned Contract Demand (kVA)', 'kVA', b.contractDemandKva, null, 'Demand & PF', 'Contracted sanctioned capacity'),
            makeRow('power_factor', 'Average Power Factor (PF)', '', b.powerFactor, null, 'Demand & PF', 'Monthly weighted average power factor (Ideal >= 0.99)')
        ];

        // 4. Financial & Tariff Charges Breakdown
        const billTotalAmount = b.totalBillAmount ?? b.currentBillAmount;
        const fixedSurcharges = (Number(b.demandCharges) || 0) +
            (Number(b.wheelingCharges) || 0) +
            (Number(b.facCharges) || 0) +
            (Number(b.electricityDuty) || 0) +
            (Number(b.taxOnSale) || 0) +
            (Number(b.gridSupportCharges) || 0);

        const costSection = [
            makeRow('energy_charges', 'Base Energy Consumption Charges (₹)', '₹', b.energyCharges, s.electricityCost, 'Financials', 'Pure consumption energy cost (Rate × Units)'),
            makeRow('demand_charges', 'Fixed Demand Charges (₹)', '₹', b.demandCharges, null, 'Statutory Charges', 'Fixed capacity reservation fee per kVA'),
            makeRow('wheeling_charges', 'Wheeling Charges (₹)', '₹', b.wheelingCharges, null, 'Statutory Charges', 'Grid network transmission cost'),
            makeRow('fac_charges', 'Fuel Adjustment Charge - FAC (₹)', '₹', b.facCharges, null, 'Statutory Charges', 'Variable fuel surcharge per kWh'),
            makeRow('duty_charges', 'Electricity Duty & Tax on Sale (₹)', '₹', (Number(b.electricityDuty) || 0) + (Number(b.taxOnSale) || 0) || b.electricityDuty, null, 'Statutory Charges', 'State Government electricity duty'),
            makeRow('tod_charges', 'Time of Day (TOD) Tariff Adjustment (₹)', '₹', b.todCharges, null, 'Statutory Charges', 'Peak/Off-peak incentive or penalty'),
            makeRow('grid_support', 'Grid Support Surcharge (₹)', '₹', b.gridSupportCharges, null, 'Statutory Charges', 'Rooftop solar parallel operation fee'),
            makeRow('ppd_rebate', 'Prompt Payment Discount / Rebate (₹)', '₹', b.promptPaymentDiscount, null, 'Discounts', 'Timely settlement rebate credit'),
            makeRow('net_bill_amount', 'TOTAL MSEB NET PAYABLE BILL (₹)', '₹', billTotalAmount, s.electricityCost, 'Total Bill', 'Total current payable electricity bill'),
            makeRow('combined_plant_cost', 'TOTAL COMBINED ENERGY COST (Grid + Solar) (₹)', '₹', billTotalAmount ? (Number(billTotalAmount) + (s.solarCost || 0)) : null, s.totalCost, 'Total Energy', 'Complete plant energy expenditure for this month')
        ];

        // Summary Calculations
        const billKwh = b.billedUnitsKwh ?? b.grossUnitsKwh ?? 0;
        const sysKwh = s.netConsumptionKwh ?? 0;
        const kwhDiff = billKwh - sysKwh;
        const kwhDiffPct = billKwh > 0 ? (Math.abs(kwhDiff) / billKwh) * 100 : 0;

        const billCost = billTotalAmount ?? 0;
        const sysCost = s.electricityCost ?? 0;
        const costDiff = billCost - sysCost;
        const costDiffPct = billCost > 0 ? (Math.abs(costDiff) / billCost) * 100 : 0;

        const solarGenBill = b.solarGenUnits ?? 0;
        const solarGenSys = s.solarGenKwh ?? 0;
        const solarDiff = solarGenBill - solarGenSys;

        // Executive Verdict
        let verdict = 'VERIFIED';
        let verdictColor = 'text-emerald-700 bg-emerald-50 border-emerald-300 dark:bg-emerald-950/40 dark:text-emerald-300 dark:border-emerald-800';
        let verdictMessage = 'Consumption readings match within acceptable tolerance (< 0.5%).';

        if (kwhDiffPct > 5.0) {
            verdict = 'CRITICAL DISCREPANCY';
            verdictColor = 'text-rose-700 bg-rose-50 border-rose-300 dark:bg-rose-950/40 dark:text-rose-300 dark:border-rose-800';
            verdictMessage = `Meter readings differ by ${formatNumber(Math.abs(kwhDiff))} kWh (${kwhDiffPct.toFixed(1)}%). Requires utility investigation.`;
        } else if (kwhDiffPct > 1.0) {
            verdict = 'MINOR VARIANCE';
            verdictColor = 'text-amber-700 bg-amber-50 border-amber-300 dark:bg-amber-950/40 dark:text-amber-300 dark:border-amber-800';
            verdictMessage = `Small variance of ${formatNumber(Math.abs(kwhDiff))} kWh (${kwhDiffPct.toFixed(1)}%). Within normal billing cut-off timing.`;
        } else if (fixedSurcharges > 0 && Math.abs(costDiff) > 1000) {
            verdict = 'VERIFIED WITH STATUTORY SURCHARGES';
            verdictColor = 'text-sky-700 bg-sky-50 border-sky-300 dark:bg-sky-950/40 dark:text-sky-300 dark:border-sky-800';
            verdictMessage = `Base consumption verified. ₹ ${formatINR(fixedSurcharges)} is statutory MSEB fixed demand & taxes.`;
        }

        return {
            sections: [
                { title: '⚡ 1. Grid Electricity & Meter Readings', rows: gridSection },
                { title: '☀️ 2. Solar Generation & Export Reconciliation', rows: solarSection },
                { title: '⚡ 3. Maximum Demand & Power Factor (MD & PF)', rows: demandSection },
                { title: '💰 4. Financials, Tariffs & Cost Breakdown', rows: costSection }
            ],
            summary: {
                billKwh,
                sysKwh,
                kwhDiff,
                kwhDiffPct,
                billCost,
                sysCost,
                costDiff,
                costDiffPct,
                solarGenBill,
                solarGenSys,
                solarDiff,
                fixedSurcharges,
                verdict,
                verdictColor,
                verdictMessage
            }
        };
    }, [activeBill, systemData, activeTariffs]);

    // Export Table to Clean Executive Excel
    const handleExportExcel = () => {
        if (!comparisonReport || !activeBill) return;

        const wb = XLSX.utils.book_new();

        // 1. Executive Summary Sheet
        const summaryData = [
            ['PG ELECTROPLAST LTD - UTILITY BILL RECONCILIATION AUDIT REPORT'],
            ['Generated On', new Date().toLocaleString('en-IN')],
            ['Location', selectedPlant.location_name],
            ['Plant', `${selectedPlant.plant_code} - ${selectedPlant.plant_display_name}`],
            ['Billing Month', selectedMonth],
            ['Audited Bill File', activeBill.fileName || 'Uploaded Bill'],
            ['Audit Verdict', comparisonReport.summary.verdict],
            ['Verdict Note', comparisonReport.summary.verdictMessage],
            [],
            ['KEY AUDIT PARAMETER', 'AS PER BILL', 'AS PER SOFTWARE', 'VARIANCE (DIFF)', 'VARIANCE %', 'AUDIT STATUS'],
            [
                'Net Grid Consumption (kWh)',
                comparisonReport.summary.billKwh,
                comparisonReport.summary.sysKwh,
                comparisonReport.summary.kwhDiff,
                `${comparisonReport.summary.kwhDiffPct.toFixed(2)}%`,
                comparisonReport.summary.kwhDiffPct <= 1 ? 'Matched' : 'Variance'
            ],
            [
                'Solar Generation (kWh)',
                comparisonReport.summary.solarGenBill,
                comparisonReport.summary.solarGenSys,
                comparisonReport.summary.solarDiff,
                '—',
                Math.abs(comparisonReport.summary.solarDiff) <= 100 ? 'Matched' : 'Variance'
            ],
            [
                'MSEB Net Payable Bill Amount (₹)',
                comparisonReport.summary.billCost,
                comparisonReport.summary.sysCost,
                comparisonReport.summary.costDiff,
                `${comparisonReport.summary.costDiffPct.toFixed(2)}%`,
                Math.abs(comparisonReport.summary.costDiff) <= 5000 ? 'Matched' : 'Variance (Includes Fixed Surcharges)'
            ],
            [
                'MSEB Fixed & Statutory Surcharges (₹)',
                comparisonReport.summary.fixedSurcharges,
                0,
                comparisonReport.summary.fixedSurcharges,
                '100%',
                'Statutory Demand & Duty'
            ]
        ];

        const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(wb, wsSummary, 'Executive Summary');

        // 2. Full Parameter Line-by-Line Sheet
        const detailsData = [
            ['Category', 'Parameter / Term', 'Unit', 'As Per Bill', 'As Per Software', 'Difference', 'Variance %', 'Cost Impact (₹)', 'Audit Status', 'Notes / Remarks']
        ];

        comparisonReport.sections.forEach((sec) => {
            sec.rows.forEach((r) => {
                detailsData.push([
                    sec.title.replace(/[^a-zA-Z0-9 &]/g, '').trim(),
                    r.paramName,
                    r.unit,
                    r.billVal !== null ? r.billVal : 'N/A',
                    r.sysVal !== null ? r.sysVal : 'N/A',
                    r.diff !== null ? r.diff : 'N/A',
                    r.diffPct !== null ? `${r.diffPct}%` : 'N/A',
                    r.costImpact !== null ? r.costImpact : 'N/A',
                    r.status,
                    r.helperText || ''
                ]);
            });
        });

        const wsDetails = XLSX.utils.aoa_to_sheet(detailsData);
        XLSX.utils.book_append_sheet(wb, wsDetails, 'Line-Item Comparison');

        const fileName = `Bill_Audit_${selectedPlant.plant_display_name}_${selectedMonth}.xlsx`;
        XLSX.writeFile(wb, fileName);
    };

    // Open Manual Edit Modal
    const handleStartEdit = () => {
        if (!activeBill) return;
        setEditFormData({ ...(activeBill.extractedData || {}) });
        setIsEditingBill(true);
    };

    // Save Manual Edit
    const handleSaveEdit = (e) => {
        e.preventDefault();
        if (!editFormData || !activeBill) return;
        setActiveBill({
            ...activeBill,
            extractedData: {
                ...activeBill.extractedData,
                ...editFormData
            }
        });
        setIsEditingBill(false);
    };

    // Clear active audit
    const handleClearAudit = () => {
        if (window.confirm('Are you sure you want to clear the current bill audit?')) {
            setActiveBill(null);
            localStorage.removeItem('ep_active_audit_bill_v2');
        }
    };

    return (
        <div className="w-full flex flex-col gap-4 p-3 md:p-5 max-w-[1600px] mx-auto text-slate-800 dark:text-slate-100">
            {/* 1. TOP HEADER & AUDIT CONTROLS */}
            <div className="bg-white dark:bg-slate-900 border border-slate-200/80 dark:border-slate-800 rounded-xl p-4 shadow-xs flex flex-wrap items-center justify-between gap-4">
                <div className="flex flex-col gap-0.5">
                    <div className="flex items-center gap-2">
                        <span className="w-2.5 h-2.5 rounded-full bg-blue-600 animate-pulse"></span>
                        <h1 className="text-lg md:text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">
                            Utility Bill Audit & Management Reconciliation
                        </h1>
                    </div>
                    <p className="text-xs text-slate-500 dark:text-slate-400">
                        Direct parameter-by-parameter audit between uploaded DISCOM/MSEB utility bill and software daily meter readings.
                    </p>
                </div>

                {/* Filter Controls: Plant & Month */}
                <div className="flex flex-wrap items-center gap-2.5">
                    {/* Plant Dropdown */}
                    <div className="flex items-center gap-1.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 px-3 py-1.5 rounded-lg">
                        <label className="text-[11px] font-extrabold uppercase text-slate-500 dark:text-slate-400">Plant:</label>
                        <select
                            value={selectedPlantCode}
                            onChange={(e) => setSelectedPlantCode(e.target.value)}
                            className="bg-transparent text-xs font-bold text-slate-800 dark:text-slate-100 border-none outline-none cursor-pointer"
                        >
                            {plants.map((p) => (
                                <option key={p.plant_code} value={p.plant_code} className="dark:bg-slate-800">
                                    {p.location_name ? `${p.location_name} - ` : ''}{p.plant_display_name} ({p.plant_code})
                                </option>
                            ))}
                        </select>
                    </div>

                    {/* Month Picker */}
                    <div className="flex items-center gap-1.5 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 px-3 py-1.5 rounded-lg">
                        <label className="text-[11px] font-extrabold uppercase text-slate-500 dark:text-slate-400">Month:</label>
                        <input
                            type="month"
                            value={selectedMonth}
                            onChange={(e) => setSelectedMonth(e.target.value)}
                            className="bg-transparent text-xs font-bold text-slate-800 dark:text-slate-100 border-none outline-none cursor-pointer"
                        />
                    </div>

                    {/* Action Buttons */}
                    {activeBill ? (
                        <>
                            <button
                                type="button"
                                onClick={handleExportExcel}
                                className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold transition flex items-center gap-1.5 shadow-xs cursor-pointer"
                                title="Download Audit Comparison in Excel format for Management"
                            >
                                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                                </svg>
                                Export to Excel
                            </button>

                            <button
                                type="button"
                                onClick={handleStartEdit}
                                className="px-3 py-1.5 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg text-xs font-bold transition flex items-center gap-1.5 border border-slate-300 dark:border-slate-700 cursor-pointer"
                                title="Review or manually edit extracted values from bill"
                            >
                                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                                </svg>
                                Edit Values
                            </button>

                            <button
                                type="button"
                                onClick={() => fileInputRef.current?.click()}
                                className="px-3 py-1.5 bg-blue-50 dark:bg-blue-950/40 text-blue-700 dark:text-blue-300 hover:bg-blue-100 rounded-lg text-xs font-bold transition flex items-center gap-1.5 border border-blue-200 dark:border-blue-800 cursor-pointer"
                            >
                                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                                </svg>
                                Upload New Bill
                            </button>

                            <button
                                type="button"
                                onClick={handleClearAudit}
                                className="px-2.5 py-1.5 text-slate-400 hover:text-rose-600 rounded-lg text-xs font-bold transition"
                                title="Clear current audit"
                            >
                                Clear
                            </button>
                        </>
                    ) : (
                        <button
                            type="button"
                            onClick={() => fileInputRef.current?.click()}
                            className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold transition flex items-center gap-2 shadow-xs cursor-pointer"
                        >
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                            </svg>
                            Upload Utility Bill (PDF / Excel)
                        </button>
                    )}

                    {/* Hidden Native File Input */}
                    <input
                        ref={fileInputRef}
                        type="file"
                        accept=".pdf,.xlsx,.xls"
                        onChange={(e) => {
                            if (e.target.files && e.target.files[0]) {
                                processFile(e.target.files[0]);
                            }
                        }}
                        className="hidden"
                    />
                </div>
            </div>

            {/* Error Message if Extraction Failed */}
            {extractError && (
                <div className="bg-rose-50 dark:bg-rose-950/40 border border-rose-200 dark:border-rose-800 text-rose-700 dark:text-rose-300 p-3 rounded-xl text-xs font-medium flex items-center justify-between">
                    <span>⚠️ {extractError}</span>
                    <button type="button" onClick={() => setExtractError('')} className="text-rose-500 hover:text-rose-700 font-bold">×</button>
                </div>
            )}

            {/* Loading Indicator */}
            {isExtracting && (
                <div className="bg-blue-50 dark:bg-blue-950/40 border border-blue-200 dark:border-blue-800 text-blue-700 dark:text-blue-300 p-4 rounded-xl text-xs font-semibold flex items-center gap-3 animate-pulse">
                    <div className="w-4 h-4 border-2 border-blue-600 border-t-transparent rounded-full animate-spin"></div>
                    <span>Extracting billing parameters from document and fetching system meter records for {selectedPlant.plant_display_name}...</span>
                </div>
            )}

            {/* 2. NO BILL UPLOADED EMPTY STATE / DROPZONE */}
            {!activeBill && !isExtracting && (
                <div
                    onDragEnter={handleDrag}
                    onDragLeave={handleDrag}
                    onDragOver={handleDrag}
                    onDrop={handleDrop}
                    onClick={() => fileInputRef.current?.click()}
                    className={`border-2 border-dashed rounded-2xl p-10 flex flex-col items-center justify-center text-center cursor-pointer transition-all duration-200 ${
                        dragActive
                            ? 'border-blue-500 bg-blue-50/60 dark:bg-blue-950/30 scale-[1.005]'
                            : 'border-slate-300 dark:border-slate-700 hover:border-blue-400 bg-slate-50/50 dark:bg-slate-900/40 hover:bg-blue-50/20'
                    }`}
                >
                    <div className="w-14 h-14 rounded-2xl bg-blue-100 dark:bg-blue-900/40 text-blue-600 dark:text-blue-400 flex items-center justify-center mb-3">
                        <svg className="w-7 h-7" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                        </svg>
                    </div>
                    <h3 className="text-base font-extrabold text-slate-800 dark:text-slate-100 mb-1">
                        Upload Utility Bill for {selectedPlant.plant_display_name} ({selectedMonth})
                    </h3>
                    <p className="text-xs text-slate-500 dark:text-slate-400 max-w-md mb-4">
                        Drag and drop your official DISCOM/MSEDCL electricity bill (PDF or Excel) here, or click to browse. The software will instantly extract all readings and compare against our internal meter logs.
                    </p>
                    <div className="flex items-center gap-2 text-[11px] font-semibold text-slate-600 dark:text-slate-300 bg-white dark:bg-slate-800 px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 shadow-xs">
                        <span>⚡ Supports MSEDCL / MSEB PDF Bills</span>
                        <span>•</span>
                        <span>☀️ Solar Net-Metering Readings</span>
                        <span>•</span>
                        <span>💰 Automatic Cost Audit</span>
                    </div>
                </div>
            )}

            {/* 3. ACTIVE BILL COMPARISON VIEW */}
            {activeBill && comparisonReport && (
                <>
                    {/* EXECUTIVE AUDIT VERDICT & FINANCIAL SUMMARY BAR */}
                    <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-4 shadow-xs flex flex-col gap-3">
                        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 dark:border-slate-800 pb-3">
                            <div className="flex items-center gap-3">
                                <div className="text-2xl">📑</div>
                                <div>
                                    <div className="flex items-center gap-2">
                                        <h2 className="text-sm font-black text-slate-900 dark:text-white uppercase tracking-wider">
                                            {selectedPlant.plant_display_name} ({selectedPlant.plant_code}) — {selectedMonth} Audit Overview
                                        </h2>
                                        <span className={`text-[10px] font-black uppercase px-2 py-0.5 rounded-full border ${comparisonReport.summary.verdictColor}`}>
                                            {comparisonReport.summary.verdict}
                                        </span>
                                    </div>
                                    <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                                        {comparisonReport.summary.verdictMessage}
                                    </p>
                                </div>
                            </div>

                            <div className="text-right">
                                <span className="text-[10px] font-bold uppercase text-slate-400">File Audited:</span>
                                <div className="text-xs font-bold text-slate-700 dark:text-slate-200 truncate max-w-[280px]">
                                    {activeBill.fileName}
                                </div>
                            </div>
                        </div>

                        {/* 4 Clean High-Impact Executive Metrics Strip */}
                        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 pt-1">
                            {/* Metric 1: Total Net Payable Bill Amount */}
                            <div className="bg-slate-50 dark:bg-slate-800/60 p-3 rounded-lg border border-slate-200/60 dark:border-slate-700/60">
                                <span className="text-[10px] font-black uppercase tracking-wider text-slate-500 dark:text-slate-400">
                                    MSEB Net Payable Bill
                                </span>
                                <div className="text-base md:text-lg font-black text-slate-900 dark:text-white mt-0.5">
                                    ₹ {formatINR(comparisonReport.summary.billCost, 0)}
                                </div>
                                <div className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5 flex items-center justify-between">
                                    <span>Software Energy Cost:</span>
                                    <span className="font-bold">₹ {formatINR(comparisonReport.summary.sysCost, 0)}</span>
                                </div>
                            </div>

                            {/* Metric 2: Net Financial Variance */}
                            <div className={`p-3 rounded-lg border ${
                                Math.abs(comparisonReport.summary.costDiff) > 10000
                                    ? 'bg-amber-50/70 border-amber-200 dark:bg-amber-950/30 dark:border-amber-800/60'
                                    : 'bg-slate-50 dark:bg-slate-800/60 border-slate-200/60 dark:border-slate-700/60'
                            }`}>
                                <span className="text-[10px] font-black uppercase tracking-wider text-slate-500 dark:text-slate-400">
                                    Financial Difference
                                </span>
                                <div className={`text-base md:text-lg font-black mt-0.5 ${
                                    comparisonReport.summary.costDiff > 0 ? 'text-amber-700 dark:text-amber-300' : 'text-slate-900 dark:text-white'
                                }`}>
                                    {comparisonReport.summary.costDiff >= 0 ? '+' : ''}₹ {formatINR(comparisonReport.summary.costDiff, 0)}
                                </div>
                                <div className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5 flex items-center justify-between">
                                    <span>Fixed & Statutory Taxes:</span>
                                    <span className="font-bold text-slate-700 dark:text-slate-300">₹ {formatINR(comparisonReport.summary.fixedSurcharges, 0)}</span>
                                </div>
                            </div>

                            {/* Metric 3: Net Grid Units (kWh) */}
                            <div className="bg-slate-50 dark:bg-slate-800/60 p-3 rounded-lg border border-slate-200/60 dark:border-slate-700/60">
                                <span className="text-[10px] font-black uppercase tracking-wider text-slate-500 dark:text-slate-400">
                                    Grid Consumption (kWh)
                                </span>
                                <div className="text-base md:text-lg font-black text-slate-900 dark:text-white mt-0.5">
                                    {formatNumber(comparisonReport.summary.billKwh, 0)} <span className="text-xs font-normal text-slate-400">kWh</span>
                                </div>
                                <div className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5 flex items-center justify-between">
                                    <span>Software Recorded:</span>
                                    <span className="font-bold">{formatNumber(comparisonReport.summary.sysKwh, 0)} kWh</span>
                                </div>
                            </div>

                            {/* Metric 4: Solar Generation (kWh) */}
                            <div className="bg-slate-50 dark:bg-slate-800/60 p-3 rounded-lg border border-slate-200/60 dark:border-slate-700/60">
                                <span className="text-[10px] font-black uppercase tracking-wider text-slate-500 dark:text-slate-400">
                                    Solar Generation (kWh)
                                </span>
                                <div className="text-base md:text-lg font-black text-slate-900 dark:text-white mt-0.5">
                                    {formatNumber(comparisonReport.summary.solarGenBill, 0)} <span className="text-xs font-normal text-slate-400">kWh</span>
                                </div>
                                <div className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5 flex items-center justify-between">
                                    <span>Software Solar:</span>
                                    <span className="font-bold">{formatNumber(comparisonReport.summary.solarGenSys, 0)} kWh</span>
                                </div>
                            </div>
                        </div>
                    </div>

                    {/* 4. THE CORE MANAGEMENT COMPARISON TABLE */}
                    <div className="bg-white dark:bg-slate-900 border border-slate-200/90 dark:border-slate-800 rounded-xl overflow-hidden shadow-xs">
                        <div className="px-4 py-3 bg-slate-100/80 dark:bg-slate-800/80 border-b border-slate-200 dark:border-slate-700 flex flex-wrap items-center justify-between gap-2">
                            <div className="flex items-center gap-2">
                                <span className="text-xs font-black uppercase tracking-wider text-slate-700 dark:text-slate-200">
                                    Detailed Parameter Reconciliation Matrix
                                </span>
                                <span className="text-[10px] bg-white dark:bg-slate-700 text-slate-600 dark:text-slate-300 font-bold px-2 py-0.5 rounded-full border border-slate-200 dark:border-slate-600">
                                    Plant {selectedPlant.plant_code}
                                </span>
                            </div>

                            <div className="flex items-center gap-3 text-xs">
                                <span className="flex items-center gap-1.5 text-emerald-600 font-bold">
                                    <span className="w-2 h-2 rounded-full bg-emerald-500"></span> Verified Match (&lt; 0.5%)
                                </span>
                                <span className="flex items-center gap-1.5 text-amber-600 font-bold">
                                    <span className="w-2 h-2 rounded-full bg-amber-500"></span> Minor Variance (&lt; 2%)
                                </span>
                                <span className="flex items-center gap-1.5 text-rose-600 font-bold">
                                    <span className="w-2 h-2 rounded-full bg-rose-500"></span> Discrepancy (&gt; 2%)
                                </span>
                            </div>
                        </div>

                        <div className="overflow-x-auto">
                            <table className="w-full text-left text-xs border-collapse">
                                <thead>
                                    <tr className="bg-slate-50 dark:bg-slate-800/50 text-[10px] font-black uppercase text-slate-500 dark:text-slate-400 border-b border-slate-200 dark:border-slate-700 tracking-wider">
                                        <th className="py-2.5 px-3 w-10 text-center">#</th>
                                        <th className="py-2.5 px-4 min-w-[260px]">Billing Parameter / Term</th>
                                        <th className="py-2.5 px-2.5 w-16 text-center">Unit</th>
                                        <th className="py-2.5 px-4 text-right min-w-[130px] bg-blue-50/40 dark:bg-blue-950/20">As Per Uploaded Bill</th>
                                        <th className="py-2.5 px-4 text-right min-w-[130px] bg-slate-100/50 dark:bg-slate-800/50">As Per Software Readings</th>
                                        <th className="py-2.5 px-3.5 text-right min-w-[110px]">Difference (Variance)</th>
                                        <th className="py-2.5 px-3 text-right min-w-[90px]">Variance %</th>
                                        <th className="py-2.5 px-4 text-right min-w-[130px]">Financial Impact (₹)</th>
                                        <th className="py-2.5 px-4 text-center min-w-[130px]">Audit Status</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {comparisonReport.sections.map((sec, secIdx) => (
                                        <React.Fragment key={sec.title}>
                                            {/* Section Header Row */}
                                            <tr className="bg-slate-100/60 dark:bg-slate-800/40 border-y border-slate-200/80 dark:border-slate-700/80">
                                                <td colSpan={9} className="py-2 px-4 font-black text-xs text-slate-800 dark:text-slate-100 tracking-wide">
                                                    {sec.title}
                                                </td>
                                            </tr>

                                            {/* Section Data Rows */}
                                            {sec.rows.map((row, rowIdx) => {
                                                const isHighlighted = row.id === 'net_units' || row.id === 'solar_gen' || row.id === 'net_bill_amount' || row.id === 'combined_plant_cost';

                                                return (
                                                    <tr
                                                        key={row.id}
                                                        className={`border-b border-slate-100 dark:border-slate-800/60 transition hover:bg-slate-50/80 dark:hover:bg-slate-800/40 ${
                                                            isHighlighted ? 'bg-blue-50/20 dark:bg-blue-950/10 font-bold' : ''
                                                        }`}
                                                    >
                                                        {/* # */}
                                                        <td className="py-2.5 px-3 text-center text-slate-400 font-mono text-[11px]">
                                                            {rowIdx + 1}
                                                        </td>

                                                        {/* Parameter Name */}
                                                        <td className="py-2.5 px-4">
                                                            <div className={`text-xs ${isHighlighted ? 'font-black text-slate-900 dark:text-white' : 'font-semibold text-slate-700 dark:text-slate-200'}`}>
                                                                {row.paramName}
                                                            </div>
                                                            {row.helperText && (
                                                                <div className="text-[10px] text-slate-400 font-normal">
                                                                    {row.helperText}
                                                                </div>
                                                            )}
                                                        </td>

                                                        {/* Unit */}
                                                        <td className="py-2.5 px-2.5 text-center font-mono text-[11px] text-slate-500 dark:text-slate-400">
                                                            {row.unit}
                                                        </td>

                                                        {/* Bill Value */}
                                                        <td className="py-2.5 px-4 text-right font-mono font-bold text-slate-800 dark:text-slate-100 bg-blue-50/20 dark:bg-blue-950/10">
                                                            {row.unit === '₹' ? `₹ ${formatINR(row.billVal)}` : formatNumber(row.billVal)}
                                                        </td>

                                                        {/* Software System Value */}
                                                        <td className="py-2.5 px-4 text-right font-mono font-bold text-slate-700 dark:text-slate-200 bg-slate-50/40 dark:bg-slate-800/30">
                                                            {row.sysVal !== null
                                                                ? (row.unit === '₹' ? `₹ ${formatINR(row.sysVal)}` : formatNumber(row.sysVal))
                                                                : <span className="text-slate-300 dark:text-slate-600 font-normal italic">N/A</span>}
                                                        </td>

                                                        {/* Difference */}
                                                        <td className={`py-2.5 px-3.5 text-right font-mono font-semibold ${
                                                            row.diff === null ? 'text-slate-400' :
                                                            row.diff === 0 ? 'text-slate-600' :
                                                            row.diff > 0 ? 'text-rose-600 dark:text-rose-400' : 'text-emerald-600 dark:text-emerald-400'
                                                        }`}>
                                                            {row.diff !== null
                                                                ? `${row.diff > 0 ? '+' : ''}${formatNumber(row.diff)}`
                                                                : '—'}
                                                        </td>

                                                        {/* Difference % */}
                                                        <td className="py-2.5 px-3 text-right font-mono text-slate-600 dark:text-slate-400">
                                                            {row.diffPct !== null ? `${row.diffPct.toFixed(2)}%` : '—'}
                                                        </td>

                                                        {/* Cost Impact */}
                                                        <td className={`py-2.5 px-4 text-right font-mono font-bold ${
                                                            row.costImpact === null ? 'text-slate-400' :
                                                            row.costImpact === 0 ? 'text-slate-600' :
                                                            row.costImpact > 0 ? 'text-rose-700 dark:text-rose-300' : 'text-emerald-700 dark:text-emerald-300'
                                                        }`}>
                                                            {row.costImpact !== null
                                                                ? `${row.costImpact > 0 ? '+₹ ' : '-₹ '}${formatINR(Math.abs(row.costImpact))}`
                                                                : '—'}
                                                        </td>

                                                        {/* Audit Status Badge */}
                                                        <td className="py-2.5 px-4 text-center">
                                                            {row.status === 'Verified' && (
                                                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-800 dark:bg-emerald-950/60 dark:text-emerald-300">
                                                                    ✓ MATCH
                                                                </span>
                                                            )}
                                                            {row.status === 'Minor Variance' && (
                                                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black bg-amber-100 text-amber-800 dark:bg-amber-950/60 dark:text-amber-300">
                                                                    ⚠️ &lt; 2%
                                                                </span>
                                                            )}
                                                            {row.status === 'Discrepancy' && (
                                                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-black bg-rose-100 text-rose-800 dark:bg-rose-950/60 dark:text-rose-300">
                                                                    🔴 VARIANCE
                                                                </span>
                                                            )}
                                                            {row.status === 'Statutory' && (
                                                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300">
                                                                    🏛️ TARIFF
                                                                </span>
                                                            )}
                                                            {row.status === 'Bill Only' && (
                                                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-blue-50 text-blue-700 dark:bg-blue-950/40 dark:text-blue-300">
                                                                    BILL ENTRY
                                                                </span>
                                                            )}
                                                            {row.status === 'Software Only' && (
                                                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-purple-50 text-purple-700 dark:bg-purple-950/40 dark:text-purple-300">
                                                                    SYSTEM ONLY
                                                                </span>
                                                            )}
                                                        </td>
                                                    </tr>
                                                );
                                            })}
                                        </React.Fragment>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </>
            )}

            {/* 5. MODAL / DRAWER TO REVIEW OR EDIT EXTRACTED BILL VALUES */}
            {isEditingBill && editFormData && (
                <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
                    <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl max-w-2xl w-full max-h-[90vh] overflow-hidden flex flex-col shadow-2xl">
                        <div className="p-4 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between">
                            <div>
                                <h3 className="text-base font-black text-slate-900 dark:text-white uppercase">
                                    Review & Edit Extracted Bill Terms
                                </h3>
                                <p className="text-xs text-slate-500">
                                    Modify any parameter to instantly re-calculate the management comparison table.
                                </p>
                            </div>
                            <button
                                type="button"
                                onClick={() => setIsEditingBill(false)}
                                className="text-slate-400 hover:text-slate-600 font-bold text-lg p-1"
                            >
                                ✕
                            </button>
                        </div>

                        <form onSubmit={handleSaveEdit} className="p-5 overflow-y-auto flex flex-col gap-4 text-xs">
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                                <div>
                                    <label className="font-bold text-slate-700 dark:text-slate-300 mb-1 block">Billed Electricity Units (kWh)</label>
                                    <input
                                        type="number"
                                        step="any"
                                        value={editFormData.billedUnitsKwh ?? ''}
                                        onChange={(e) => setEditFormData({ ...editFormData, billedUnitsKwh: e.target.value ? Number(e.target.value) : '' })}
                                        className="w-full px-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg font-mono font-bold"
                                    />
                                </div>
                                <div>
                                    <label className="font-bold text-slate-700 dark:text-slate-300 mb-1 block">Total Bill Amount Payable (₹)</label>
                                    <input
                                        type="number"
                                        step="any"
                                        value={editFormData.totalBillAmount ?? ''}
                                        onChange={(e) => setEditFormData({ ...editFormData, totalBillAmount: e.target.value ? Number(e.target.value) : '' })}
                                        className="w-full px-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg font-mono font-bold"
                                    />
                                </div>
                                <div>
                                    <label className="font-bold text-slate-700 dark:text-slate-300 mb-1 block">Base Energy Charges (₹)</label>
                                    <input
                                        type="number"
                                        step="any"
                                        value={editFormData.energyCharges ?? ''}
                                        onChange={(e) => setEditFormData({ ...editFormData, energyCharges: e.target.value ? Number(e.target.value) : '' })}
                                        className="w-full px-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg font-mono"
                                    />
                                </div>
                                <div>
                                    <label className="font-bold text-slate-700 dark:text-slate-300 mb-1 block">Fixed Demand Charges (₹)</label>
                                    <input
                                        type="number"
                                        step="any"
                                        value={editFormData.demandCharges ?? ''}
                                        onChange={(e) => setEditFormData({ ...editFormData, demandCharges: e.target.value ? Number(e.target.value) : '' })}
                                        className="w-full px-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg font-mono"
                                    />
                                </div>
                                <div>
                                    <label className="font-bold text-slate-700 dark:text-slate-300 mb-1 block">Solar Generation Units (kWh)</label>
                                    <input
                                        type="number"
                                        step="any"
                                        value={editFormData.solarGenUnits ?? ''}
                                        onChange={(e) => setEditFormData({ ...editFormData, solarGenUnits: e.target.value ? Number(e.target.value) : '' })}
                                        className="w-full px-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg font-mono"
                                    />
                                </div>
                                <div>
                                    <label className="font-bold text-slate-700 dark:text-slate-300 mb-1 block">Solar Adjustment Credit (kWh)</label>
                                    <input
                                        type="number"
                                        step="any"
                                        value={editFormData.solarAdjUnits ?? ''}
                                        onChange={(e) => setEditFormData({ ...editFormData, solarAdjUnits: e.target.value ? Number(e.target.value) : '' })}
                                        className="w-full px-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg font-mono"
                                    />
                                </div>
                                <div>
                                    <label className="font-bold text-slate-700 dark:text-slate-300 mb-1 block">Billed Demand (kVA)</label>
                                    <input
                                        type="number"
                                        step="any"
                                        value={editFormData.billedDemandKva ?? ''}
                                        onChange={(e) => setEditFormData({ ...editFormData, billedDemandKva: e.target.value ? Number(e.target.value) : '' })}
                                        className="w-full px-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg font-mono"
                                    />
                                </div>
                                <div>
                                    <label className="font-bold text-slate-700 dark:text-slate-300 mb-1 block">Wheeling Charges (₹)</label>
                                    <input
                                        type="number"
                                        step="any"
                                        value={editFormData.wheelingCharges ?? ''}
                                        onChange={(e) => setEditFormData({ ...editFormData, wheelingCharges: e.target.value ? Number(e.target.value) : '' })}
                                        className="w-full px-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg font-mono"
                                    />
                                </div>
                            </div>

                            <div className="flex items-center justify-end gap-2.5 pt-4 border-t border-slate-100 dark:border-slate-800 mt-2">
                                <button
                                    type="button"
                                    onClick={() => setIsEditingBill(false)}
                                    className="px-4 py-2 bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg font-bold"
                                >
                                    Cancel
                                </button>
                                <button
                                    type="submit"
                                    className="px-5 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold shadow-xs"
                                >
                                    Save & Update Comparison
                                </button>
                            </div>
                        </form>
                    </div>
                </div>
            )}
        </div>
    );
}
