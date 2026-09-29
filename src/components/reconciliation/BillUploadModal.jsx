// UtilitySense Smart Bill Upload & Data Review Modal
// Implements the 9-Step upload & review flow with PDF extraction and audit trail

import React, { useState, useEffect } from 'react';
import { extractTextFromPdfBuffer, parseMsedclBillText, getBillParametersList } from '../../utils/msedclBillParser';
import { saveBill } from './reconciliationService';

export default function BillUploadModal({
    isOpen,
    onClose,
    onBillSaved,
    locations = [],
    plants = [],
    currentUser,
    initialLocation = "",
    initialPlant = "",
    initialUtility = "electricity",
    initialMonth = ""
}) {
    // Current wizard step: 1 = "Setup & File Upload", 2 = "Extraction & Review Form"
    const [step, setStep] = useState(1);

    // Form selection states
    const [selectedLocation, setSelectedLocation] = useState(initialLocation || (locations[0] || ""));
    const [selectedPlant, setSelectedPlant] = useState(initialPlant || "");
    const [selectedUtility, setSelectedUtility] = useState(initialUtility || "electricity");
    const [selectedMonth, setSelectedMonth] = useState(initialMonth || "2026-07");

    // File upload & extraction states
    const [billFile, setBillFile] = useState(null);
    const [fileName, setFileName] = useState("");
    const [fileDataUrl, setFileDataUrl] = useState(null);
    const [isExtracting, setIsExtracting] = useState(false);
    const [extractError, setExtractError] = useState("");
    const [rawPdfText, setRawPdfText] = useState("");

    // Extracted / Editable Bill Data State
    const [billData, setBillData] = useState({
        consumerNo: "",
        consumerName: "",
        meterNo: "",
        tariffCategory: "",
        billDate: "",
        
        // Consumption
        grossUnitsKwh: "",
        billedUnitsKwh: "",
        billedUnitsKvah: "",
        openingMeter: "",
        closingMeter: "",
        meterDifference: "",
        multiplyingFactor: "1",
        rkvahLag: "",
        rkvahLead: "",

        // Solar
        solarGenUnits: "",
        solarExportUnits: "",
        solarAdjUnits: "",
        solarCapacity: "",

        // Demand & PF
        recordedDemandKva: "",
        billedDemandKva: "",
        contractDemandKva: "",
        kwMaxDemand: "",
        powerFactor: "",

        // Costs
        energyCharges: "",
        demandCharges: "",
        excessDemandCharges: "",
        wheelingCharges: "",
        facCharges: "",
        electricityDuty: "",
        todCharges: "",
        gridSupportCharges: "",
        promptPaymentDiscount: "",
        subsidiesTotal: "",
        currentBillAmount: "",
        totalBillAmount: "",
        notes: ""
    });

    const [activeReviewTab, setActiveReviewTab] = useState("consumption"); // "consumption" | "solar" | "demand" | "cost"

    // Sync initial props
    useEffect(() => {
        if (isOpen) {
            setStep(1);
            setBillFile(null);
            setFileName("");
            setFileDataUrl(null);
            setExtractError("");
            if (initialLocation) setSelectedLocation(initialLocation);
            if (initialPlant) setSelectedPlant(initialPlant);
            if (initialUtility) setSelectedUtility(initialUtility);
            if (initialMonth) setSelectedMonth(initialMonth);
        }
    }, [isOpen, initialLocation, initialPlant, initialUtility, initialMonth]);

    // Filter plants based on chosen location
    const availablePlants = plants.filter(p => {
        if (!selectedLocation || selectedLocation === "all") return true;
        return String(p.location || "").toUpperCase() === String(selectedLocation).toUpperCase();
    });

    // Auto-select first plant when location changes if current plant is invalid
    useEffect(() => {
        if (availablePlants.length > 0) {
            const exists = availablePlants.some(p => p.plant_code === selectedPlant);
            if (!exists) {
                setSelectedPlant(availablePlants[0].plant_code);
            }
        }
    }, [selectedLocation, availablePlants, selectedPlant]);

    if (!isOpen) return null;

    // Handle PDF File Selection & Extraction
    const handleFileChange = async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;

        if (!file.name.toLowerCase().endsWith(".pdf")) {
            setExtractError("Please upload a valid PDF utility bill document.");
            return;
        }

        setBillFile(file);
        setFileName(file.name);
        setExtractError("");
        setIsExtracting(true);

        try {
            // Read as data URL for PDF preview
            const reader = new FileReader();
            reader.onload = (event) => {
                setFileDataUrl(event.target.result);
            };
            reader.readAsDataURL(file);

            // Read array buffer for PDF.js text extraction
            const arrayBuffer = await file.arrayBuffer();
            const extractedText = await extractTextFromPdfBuffer(arrayBuffer);
            setRawPdfText(extractedText);

            if (!extractedText || extractedText.trim().length === 0) {
                setExtractError("Could not extract digital text from this PDF. You can still fill in the values manually below.");
                setStep(2);
                setIsExtracting(false);
                return;
            }

            // Run MSEDCL Smart Parser
            const parsed = parseMsedclBillText(extractedText);
            if (parsed) {
                // Auto-suggest plant or location if detected in bill and user hasn't locked
                if (parsed.suggestedPlantCode && plants.some(p => p.plant_code === parsed.suggestedPlantCode)) {
                    setSelectedPlant(parsed.suggestedPlantCode);
                }
                if (parsed.suggestedLocation && locations.includes(parsed.suggestedLocation)) {
                    setSelectedLocation(parsed.suggestedLocation);
                }
                if (parsed.billMonth) {
                    setSelectedMonth(parsed.billMonth);
                }

                // Populate form state
                const finalBilledKwh = parsed.billedUnitsKwh ?? parsed.billedUnitsKvah ?? "";
                const finalGrossKwh = parsed.grossUnitsKwh ?? finalBilledKwh;

                setBillData({
                    consumerNo: parsed.consumerNo || "",
                    consumerName: parsed.consumerName || "",
                    meterNo: parsed.meterNo || "",
                    tariffCategory: parsed.tariffCategory || "",
                    billDate: parsed.billDate || new Date().toISOString().split("T")[0],
                    grossUnitsKwh: finalGrossKwh,
                    billedUnitsKwh: finalBilledKwh,
                    billedUnitsKvah: parsed.billedUnitsKvah ?? "",
                    openingMeter: parsed.openingMeter ?? "",
                    closingMeter: parsed.closingMeter ?? "",
                    meterDifference: parsed.meterDifference ?? (finalGrossKwh ? Number((finalGrossKwh / (parsed.multiplyingFactor || 1)).toFixed(3)) : ""),
                    multiplyingFactor: parsed.multiplyingFactor ?? "1",
                    rkvahLag: parsed.rkvahLag ?? "",
                    rkvahLead: parsed.rkvahLead ?? "",
                    solarGenUnits: parsed.solarGenUnits ?? "",
                    solarExportUnits: parsed.solarExportUnits ?? "",
                    solarAdjUnits: parsed.solarAdjUnits ?? "",
                    solarCapacity: parsed.solarCapacity ?? "",
                    recordedDemandKva: parsed.recordedDemandKva ?? "",
                    billedDemandKva: parsed.billedDemandKva ?? "",
                    contractDemandKva: parsed.contractDemandKva ?? "",
                    kwMaxDemand: parsed.kwMaxDemand ?? "",
                    powerFactor: parsed.powerFactor ?? "",
                    energyCharges: parsed.energyCharges ?? "",
                    demandCharges: parsed.demandCharges ?? "",
                    excessDemandCharges: parsed.excessDemandCharges ?? "",
                    wheelingCharges: parsed.wheelingCharges ?? "",
                    facCharges: parsed.facCharges ?? "",
                    electricityDuty: parsed.electricityDuty ?? "",
                    todCharges: parsed.todCharges ?? "",
                    gridSupportCharges: parsed.gridSupportCharges ?? "",
                    promptPaymentDiscount: parsed.promptPaymentDiscount ?? "",
                    subsidiesTotal: parsed.subsidiesTotal ?? "",
                    currentBillAmount: parsed.currentBillAmount ?? "",
                    totalBillAmount: parsed.totalBillAmount ?? "",
                    notes: `Extracted from ${file.name}`
                });
            }

            // Advance to Review Step
            setStep(2);
        } catch (err) {
            console.error("Extraction error:", err);
            setExtractError("Extraction failed: " + err.message + ". You can continue with manual entry.");
            setStep(2);
        } finally {
            setIsExtracting(false);
        }
    };

    // Calculate meter diff on manual change
    const handleReadingChange = (field, val) => {
        const next = { ...billData, [field]: val };
        const op = Number(field === "openingMeter" ? val : next.openingMeter);
        const cl = Number(field === "closingMeter" ? val : next.closingMeter);
        if (!isNaN(op) && !isNaN(cl) && cl >= op) {
            next.meterDifference = cl - op;
        }
        setBillData(next);
    };

    // Handle Final Confirmation & Save
    const handleConfirmSave = async () => {
        if (!selectedLocation || !selectedPlant || !selectedMonth) {
            alert("Please select Location, Plant, and Bill Month.");
            return;
        }

        const toNum = (val) => (val === "" || val === null || isNaN(val) ? null : Number(val));

        const structuredData = {
            billMonth: selectedMonth,
            billDate: billData.billDate || new Date().toISOString().split("T")[0],
            consumerNo: billData.consumerNo,
            consumerName: billData.consumerName,
            meterNo: billData.meterNo,
            tariffCategory: billData.tariffCategory,
            
            grossUnitsKwh: toNum(billData.grossUnitsKwh),
            billedUnitsKwh: toNum(billData.billedUnitsKwh),
            billedUnitsKvah: toNum(billData.billedUnitsKvah),
            openingMeter: toNum(billData.openingMeter),
            closingMeter: toNum(billData.closingMeter),
            meterDifference: toNum(billData.meterDifference),
            multiplyingFactor: toNum(billData.multiplyingFactor) || 1,
            rkvahLag: toNum(billData.rkvahLag),
            rkvahLead: toNum(billData.rkvahLead),

            solarGenUnits: toNum(billData.solarGenUnits),
            solarExportUnits: toNum(billData.solarExportUnits),
            solarAdjUnits: toNum(billData.solarAdjUnits),
            solarCapacity: toNum(billData.solarCapacity),

            recordedDemandKva: toNum(billData.recordedDemandKva),
            billedDemandKva: toNum(billData.billedDemandKva),
            contractDemandKva: toNum(billData.contractDemandKva),
            kwMaxDemand: toNum(billData.kwMaxDemand),
            powerFactor: toNum(billData.powerFactor),

            energyCharges: toNum(billData.energyCharges),
            demandCharges: toNum(billData.demandCharges),
            excessDemandCharges: toNum(billData.excessDemandCharges),
            wheelingCharges: toNum(billData.wheelingCharges),
            facCharges: toNum(billData.facCharges),
            electricityDuty: toNum(billData.electricityDuty),
            todCharges: toNum(billData.todCharges),
            gridSupportCharges: toNum(billData.gridSupportCharges),
            promptPaymentDiscount: toNum(billData.promptPaymentDiscount),
            subsidiesTotal: toNum(billData.subsidiesTotal),
            currentBillAmount: toNum(billData.currentBillAmount),
            totalBillAmount: toNum(billData.totalBillAmount)
        };

        const billId = `bill_${selectedLocation}_${selectedPlant}_${selectedUtility}_${selectedMonth}`.replace(/\s+/g, '_');

        const newBillRecord = {
            id: billId,
            location: selectedLocation,
            plant: selectedPlant,
            utility: selectedUtility,
            billMonth: selectedMonth,
            billDate: billData.billDate || new Date().toISOString().split("T")[0],
            consumerNumber: billData.consumerNo || "—",
            meterNumber: billData.meterNo || "—",
            fileName: fileName || "Manual_Entry.pdf",
            fileData: fileDataUrl,
            status: "Matched",
            notes: billData.notes || "Verified by " + (currentUser?.name || "User"),
            uploadedBy: currentUser?.name || "IT Admin",
            uploadedAt: new Date().toISOString(),
            lastModifiedBy: currentUser?.name || "IT Admin",
            lastModifiedAt: new Date().toISOString(),
            extractedData: structuredData
        };

        await saveBill(newBillRecord);
        onBillSaved(newBillRecord);
        onClose();
    };

    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-3 sm:p-4 overflow-y-auto">
            <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-4xl max-h-[92vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
                
                {/* Modal Header */}
                <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/70">
                    <div className="flex items-center gap-2.5">
                        <div className="w-8 h-8 rounded-xl bg-sky-100 text-sky-700 flex items-center justify-center font-bold">
                            <span className="material-symbols-outlined text-[20px]">upload_file</span>
                        </div>
                        <div>
                            <h3 className="text-sm font-extrabold text-slate-900 uppercase tracking-tight">Upload Utility Bill & Reconcile</h3>
                            <p className="text-[11px] text-slate-500 font-medium">Location & Plant monthly utility invoice verification</p>
                        </div>
                    </div>
                    <button
                        onClick={onClose}
                        className="w-7 h-7 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 flex items-center justify-center transition border-none cursor-pointer bg-transparent"
                    >
                        <span className="material-symbols-outlined text-[18px]">close</span>
                    </button>
                </div>

                {/* Progress Step Header */}
                <div className="px-6 py-2.5 bg-white border-b border-slate-100 flex items-center justify-between text-xs font-bold text-slate-600">
                    <div className="flex items-center gap-6">
                        <div className={`flex items-center gap-2 ${step === 1 ? 'text-sky-700 font-black' : 'text-slate-400'}`}>
                            <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] ${step === 1 ? 'bg-sky-600 text-white' : 'bg-slate-200 text-slate-600'}`}>1</span>
                            <span>Target & File Upload</span>
                        </div>
                        <span className="text-slate-300">→</span>
                        <div className={`flex items-center gap-2 ${step === 2 ? 'text-sky-700 font-black' : 'text-slate-400'}`}>
                            <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] ${step === 2 ? 'bg-sky-600 text-white' : 'bg-slate-200 text-slate-600'}`}>2</span>
                            <span>Extracted Data Review & Audit</span>
                        </div>
                    </div>
                    {step === 2 && (
                        <button
                            type="button"
                            onClick={() => setStep(1)}
                            className="text-xs text-slate-500 hover:text-sky-700 bg-transparent border-none cursor-pointer font-bold flex items-center gap-1"
                        >
                            <span className="material-symbols-outlined text-[14px]">arrow_back</span>
                            <span>Change File / Targets</span>
                        </button>
                    )}
                </div>

                {/* Modal Body */}
                <div className="p-6 overflow-y-auto flex-1 text-slate-800">

                    {/* ================= STEP 1: HIERARCHY SELECT & UPLOAD ================= */}
                    {step === 1 && (
                        <div className="space-y-6">
                            {/* Hierarchy Form Controls */}
                            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 bg-slate-50 p-4 rounded-xl border border-slate-200/80">
                                <div>
                                    <label className="block text-[11px] font-extrabold uppercase tracking-wider text-slate-600 mb-1">
                                        Step 1: Location <span className="text-rose-500">*</span>
                                    </label>
                                    <select
                                        value={selectedLocation}
                                        onChange={(e) => setSelectedLocation(e.target.value)}
                                        className="w-full h-9 rounded-lg border border-slate-200 px-2.5 text-xs bg-white font-bold text-slate-800 focus:outline-sky-500"
                                    >
                                        {locations.map(loc => (
                                            <option key={loc} value={loc}>{loc}</option>
                                        ))}
                                    </select>
                                </div>

                                <div>
                                    <label className="block text-[11px] font-extrabold uppercase tracking-wider text-slate-600 mb-1">
                                        Step 2: Plant <span className="text-rose-500">*</span>
                                    </label>
                                    <select
                                        value={selectedPlant}
                                        onChange={(e) => setSelectedPlant(e.target.value)}
                                        className="w-full h-9 rounded-lg border border-slate-200 px-2.5 text-xs bg-white font-bold text-slate-800 focus:outline-sky-500"
                                    >
                                        {availablePlants.map(p => (
                                            <option key={p.plant_code} value={p.plant_code}>
                                                {p.plant_code} - {p.plant_display_name || p.plant_name}
                                            </option>
                                        ))}
                                    </select>
                                </div>

                                <div>
                                    <label className="block text-[11px] font-extrabold uppercase tracking-wider text-slate-600 mb-1">
                                        Step 3: Utility <span className="text-rose-500">*</span>
                                    </label>
                                    <select
                                        value={selectedUtility}
                                        onChange={(e) => setSelectedUtility(e.target.value)}
                                        className="w-full h-9 rounded-lg border border-slate-200 px-2.5 text-xs bg-white font-bold text-slate-800 focus:outline-sky-500"
                                    >
                                        <option value="electricity">Electricity (Grid / MSEDCL)</option>
                                        <option value="solar">Solar Power</option>
                                        <option value="water">Water</option>
                                        <option value="png">PNG Gas</option>
                                        <option value="diesel">Diesel</option>
                                    </select>
                                </div>

                                <div>
                                    <label className="block text-[11px] font-extrabold uppercase tracking-wider text-slate-600 mb-1">
                                        Step 4: Bill Month <span className="text-rose-500">*</span>
                                    </label>
                                    <input
                                        type="month"
                                        value={selectedMonth}
                                        onChange={(e) => setSelectedMonth(e.target.value)}
                                        className="w-full h-9 rounded-lg border border-slate-200 px-2.5 text-xs bg-white font-bold text-slate-800 focus:outline-sky-500"
                                    />
                                </div>
                            </div>

                            {/* Step 5: PDF Bill Drop Zone */}
                            <div className="space-y-2">
                                <label className="block text-xs font-extrabold uppercase tracking-wider text-slate-700">
                                    Step 5: Upload PDF Bill
                                </label>
                                
                                <div className="border-2 border-dashed border-slate-200 hover:border-sky-400 bg-slate-50/50 hover:bg-sky-50/30 rounded-2xl p-8 flex flex-col items-center justify-center text-center transition cursor-pointer relative">
                                    <input
                                        type="file"
                                        accept=".pdf,application/pdf"
                                        onChange={handleFileChange}
                                        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                                    />
                                    <div className="w-12 h-12 rounded-2xl bg-white shadow-sm border border-slate-200 flex items-center justify-center text-sky-600 mb-3">
                                        <span className="material-symbols-outlined text-[28px]">picture_as_pdf</span>
                                    </div>
                                    <h4 className="text-sm font-bold text-slate-800 mb-1">
                                        {fileName ? fileName : "Drag & drop PDF utility bill here, or click to browse"}
                                    </h4>
                                    <p className="text-xs text-slate-500 max-w-md">
                                        Supports MSEDCL industrial electricity bills, net metering solar adjustments, and utility statements.
                                    </p>
                                </div>

                                {isExtracting && (
                                    <div className="p-4 bg-sky-50 border border-sky-200 rounded-xl flex items-center gap-3 text-sky-800 text-xs font-bold animate-pulse">
                                        <span className="material-symbols-outlined text-[20px] animate-spin">progress_activity</span>
                                        <span>Parsing PDF bill structure, extracting units, demand, solar credits, and tariff breakdown...</span>
                                    </div>
                                )}

                                {extractError && (
                                    <div className="p-3 bg-amber-50 border border-amber-200 rounded-xl flex items-center gap-2.5 text-amber-800 text-xs font-semibold">
                                        <span className="material-symbols-outlined text-[18px] text-amber-600">warning</span>
                                        <span>{extractError}</span>
                                    </div>
                                )}
                            </div>

                            {/* Manual Entry Fallback Button */}
                            <div className="flex items-center justify-between pt-2 border-t border-slate-100">
                                <span className="text-[11px] text-slate-400">Do not have PDF ready? Enter bill values manually.</span>
                                <button
                                    type="button"
                                    onClick={() => setStep(2)}
                                    className="px-3.5 py-1.5 rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-100 text-xs font-bold transition cursor-pointer bg-white"
                                >
                                    Proceed to Manual Entry Form →
                                </button>
                            </div>
                        </div>
                    )}

                    {/* ================= STEP 2: REVIEW & CONFIRM FORM ================= */}
                    {step === 2 && (
                        <div className="space-y-5">
                            {/* Summary Target Banner */}
                            <div className="flex flex-wrap items-center justify-between gap-3 p-3 bg-sky-50/70 border border-sky-100 rounded-xl text-xs">
                                <div className="flex items-center gap-3">
                                    <span className="px-2 py-0.5 rounded bg-sky-600 text-white font-extrabold uppercase text-[10px]">Active Reconcile Target</span>
                                    <span className="font-bold text-slate-800">
                                        {selectedLocation} • Plant {selectedPlant} • {selectedUtility.toUpperCase()} • Month {selectedMonth}
                                    </span>
                                </div>
                                <div className="text-[11px] text-slate-500">
                                    File: <span className="font-bold text-slate-700">{fileName || "Manual Entry"}</span>
                                </div>
                            </div>

                            {/* Header Bill Info Cards */}
                            <div className="bg-slate-50/80 p-3 rounded-xl border border-slate-200">
                                <div className="text-[10px] font-black uppercase tracking-wider text-slate-500 mb-2">
                                    Bill Identification & Period
                                </div>
                                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                                    <div>
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Consumer No.</label>
                                        <input
                                            type="text"
                                            value={billData.consumerNo}
                                            onChange={(e) => setBillData({ ...billData, consumerNo: e.target.value })}
                                            placeholder="e.g. 170019014520"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Meter Serial No.</label>
                                        <input
                                            type="text"
                                            value={billData.meterNo}
                                            onChange={(e) => setBillData({ ...billData, meterNo: e.target.value })}
                                            placeholder="e.g. LT-884912"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Bill Date</label>
                                        <input
                                            type="date"
                                            value={billData.billDate}
                                            onChange={(e) => setBillData({ ...billData, billDate: e.target.value })}
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Tariff Category</label>
                                        <input
                                            type="text"
                                            value={billData.tariffCategory}
                                            onChange={(e) => setBillData({ ...billData, tariffCategory: e.target.value })}
                                            placeholder="e.g. HT-1 (A)"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                </div>
                            </div>

                            {/* Section Sub-Tabs */}
                            <div className="flex items-center gap-1 border-b border-slate-200 pb-1">
                                <button
                                    type="button"
                                    onClick={() => setActiveReviewTab("consumption")}
                                    className={`px-3 py-1.5 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeReviewTab === "consumption" ? "bg-sky-600 text-white shadow-xs" : "bg-transparent text-slate-600 hover:bg-slate-100"}`}
                                >
                                    1. Consumption & Meters
                                </button>
                                <button
                                    type="button"
                                    onClick={() => setActiveReviewTab("solar")}
                                    className={`px-3 py-1.5 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeReviewTab === "solar" ? "bg-sky-600 text-white shadow-xs" : "bg-transparent text-slate-600 hover:bg-slate-100"}`}
                                >
                                    2. Solar Reconciliation
                                </button>
                                <button
                                    type="button"
                                    onClick={() => setActiveReviewTab("demand")}
                                    className={`px-3 py-1.5 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeReviewTab === "demand" ? "bg-sky-600 text-white shadow-xs" : "bg-transparent text-slate-600 hover:bg-slate-100"}`}
                                >
                                    3. Demand & Power Factor
                                </button>
                                <button
                                    type="button"
                                    onClick={() => setActiveReviewTab("cost")}
                                    className={`px-3 py-1.5 text-xs font-bold rounded-lg border-none cursor-pointer transition ${activeReviewTab === "cost" ? "bg-sky-600 text-white shadow-xs" : "bg-transparent text-slate-600 hover:bg-slate-100"}`}
                                >
                                    4. Cost & Charges (₹)
                                </button>
                            </div>

                            {/* Tab 1: Consumption Parameters */}
                            {activeReviewTab === "consumption" && (
                                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Gross Consumption (kWh)</label>
                                        <input
                                            type="number"
                                            value={billData.grossUnitsKwh}
                                            onChange={(e) => setBillData({ ...billData, grossUnitsKwh: e.target.value })}
                                            placeholder="e.g. 145910"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Net / Billed Units (kWh)</label>
                                        <input
                                            type="number"
                                            value={billData.billedUnitsKwh}
                                            onChange={(e) => setBillData({ ...billData, billedUnitsKwh: e.target.value })}
                                            placeholder="e.g. 124800"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Billed Units (kVAh)</label>
                                        <input
                                            type="number"
                                            value={billData.billedUnitsKvah}
                                            onChange={(e) => setBillData({ ...billData, billedUnitsKvah: e.target.value })}
                                            placeholder="e.g. 128450"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Previous Reading (kWh)</label>
                                        <input
                                            type="number"
                                            value={billData.openingMeter}
                                            onChange={(e) => handleReadingChange("openingMeter", e.target.value)}
                                            placeholder="e.g. 48920"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Current Reading (kWh)</label>
                                        <input
                                            type="number"
                                            value={billData.closingMeter}
                                            onChange={(e) => handleReadingChange("closingMeter", e.target.value)}
                                            placeholder="e.g. 51338"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Meter Difference (kWh)</label>
                                        <input
                                            type="number"
                                            value={billData.meterDifference}
                                            onChange={(e) => setBillData({ ...billData, meterDifference: e.target.value })}
                                            placeholder="Calculated or read"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Multiplying Factor (MF)</label>
                                        <input
                                            type="number"
                                            value={billData.multiplyingFactor}
                                            onChange={(e) => setBillData({ ...billData, multiplyingFactor: e.target.value })}
                                            placeholder="e.g. 60"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">RKVAH Lag</label>
                                        <input
                                            type="number"
                                            value={billData.rkvahLag}
                                            onChange={(e) => setBillData({ ...billData, rkvahLag: e.target.value })}
                                            placeholder="e.g. 4200"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">RKVAH Lead</label>
                                        <input
                                            type="number"
                                            value={billData.rkvahLead}
                                            onChange={(e) => setBillData({ ...billData, rkvahLead: e.target.value })}
                                            placeholder="e.g. 1100"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                </div>
                            )}

                            {/* Tab 2: Solar Reconciliation */}
                            {activeReviewTab === "solar" && (
                                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                                    <div className="p-3 bg-amber-50/50 border border-amber-200/70 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-amber-900 mb-1">Solar Generation (kWh)</label>
                                        <input
                                            type="number"
                                            value={billData.solarGenUnits}
                                            onChange={(e) => setBillData({ ...billData, solarGenUnits: e.target.value })}
                                            placeholder="e.g. 21110"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-amber-50/50 border border-amber-200/70 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-amber-900 mb-1">Solar Export (kWh)</label>
                                        <input
                                            type="number"
                                            value={billData.solarExportUnits}
                                            onChange={(e) => setBillData({ ...billData, solarExportUnits: e.target.value })}
                                            placeholder="e.g. 2401"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-amber-50/50 border border-amber-200/70 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-amber-900 mb-1">Solar Adjustment (kWh)</label>
                                        <input
                                            type="number"
                                            value={billData.solarAdjUnits}
                                            onChange={(e) => setBillData({ ...billData, solarAdjUnits: e.target.value })}
                                            placeholder="e.g. 2401"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-amber-50/50 border border-amber-200/70 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-amber-900 mb-1">Sanctioned Solar (kWp)</label>
                                        <input
                                            type="number"
                                            value={billData.solarCapacity}
                                            onChange={(e) => setBillData({ ...billData, solarCapacity: e.target.value })}
                                            placeholder="e.g. 250"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                </div>
                            )}

                            {/* Tab 3: Demand & Power Factor */}
                            {activeReviewTab === "demand" && (
                                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Recorded MD (kVA)</label>
                                        <input
                                            type="number"
                                            value={billData.recordedDemandKva}
                                            onChange={(e) => setBillData({ ...billData, recordedDemandKva: e.target.value })}
                                            placeholder="e.g. 940"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Billed Demand (kVA)</label>
                                        <input
                                            type="number"
                                            value={billData.billedDemandKva}
                                            onChange={(e) => setBillData({ ...billData, billedDemandKva: e.target.value })}
                                            placeholder="e.g. 980"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Contract Demand (kVA)</label>
                                        <input
                                            type="number"
                                            value={billData.contractDemandKva}
                                            onChange={(e) => setBillData({ ...billData, contractDemandKva: e.target.value })}
                                            placeholder="e.g. 1250"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Power Factor (PF)</label>
                                        <input
                                            type="number"
                                            step="0.001"
                                            value={billData.powerFactor}
                                            onChange={(e) => setBillData({ ...billData, powerFactor: e.target.value })}
                                            placeholder="e.g. 0.995"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                </div>
                            )}

                            {/* Tab 4: Cost & Charges Breakdown */}
                            {activeReviewTab === "cost" && (
                                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Energy Charges (₹)</label>
                                        <input
                                            type="number"
                                            value={billData.energyCharges}
                                            onChange={(e) => setBillData({ ...billData, energyCharges: e.target.value })}
                                            placeholder="e.g. 1048320"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Demand Charges (₹)</label>
                                        <input
                                            type="number"
                                            value={billData.demandCharges}
                                            onChange={(e) => setBillData({ ...billData, demandCharges: e.target.value })}
                                            placeholder="e.g. 480200"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Wheeling Charges (₹)</label>
                                        <input
                                            type="number"
                                            value={billData.wheelingCharges}
                                            onChange={(e) => setBillData({ ...billData, wheelingCharges: e.target.value })}
                                            placeholder="e.g. 147264"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">FAC Surcharge (₹)</label>
                                        <input
                                            type="number"
                                            value={billData.facCharges}
                                            onChange={(e) => setBillData({ ...billData, facCharges: e.target.value })}
                                            placeholder="e.g. 68640"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">Electricity Duty (₹)</label>
                                        <input
                                            type="number"
                                            value={billData.electricityDuty}
                                            onChange={(e) => setBillData({ ...billData, electricityDuty: e.target.value })}
                                            placeholder="e.g. 115200"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">PPD Rebate / Discount (₹)</label>
                                        <input
                                            type="number"
                                            value={billData.promptPaymentDiscount}
                                            onChange={(e) => setBillData({ ...billData, promptPaymentDiscount: e.target.value })}
                                            placeholder="e.g. 18420"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-sky-50/70 border border-sky-200 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-sky-800 mb-1">Current Bill Amount (₹)</label>
                                        <input
                                            type="number"
                                            value={billData.currentBillAmount}
                                            onChange={(e) => setBillData({ ...billData, currentBillAmount: e.target.value })}
                                            placeholder="e.g. 1831254"
                                            className="w-full h-8 rounded-lg border border-slate-200 px-2.5 text-xs font-bold text-slate-800 bg-white"
                                        />
                                    </div>
                                    <div className="p-3 bg-emerald-50/70 border border-emerald-200 rounded-xl">
                                        <label className="block text-[10px] font-extrabold uppercase text-emerald-800 mb-1">Net Payable Amount (₹)</label>
                                        <input
                                            type="number"
                                            value={billData.totalBillAmount}
                                            onChange={(e) => setBillData({ ...billData, totalBillAmount: e.target.value })}
                                            placeholder="e.g. 1812834"
                                            className="w-full h-8 rounded-lg border border-emerald-300 px-2.5 text-xs font-extrabold text-emerald-900 bg-white"
                                        />
                                    </div>
                                </div>
                            )}

                            {/* Audit Metadata Card */}
                            <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 text-xs flex flex-wrap items-center justify-between text-slate-500">
                                <div>
                                    Uploaded By: <span className="font-bold text-slate-700">{currentUser?.name || "IT Admin"}</span>
                                </div>
                                <div>
                                    Audit Timestamp: <span className="font-bold text-slate-700">{new Date().toLocaleString()}</span>
                                </div>
                                <div className="text-emerald-600 font-bold flex items-center gap-1">
                                    <span className="material-symbols-outlined text-[15px]">verified</span>
                                    <span>Ready for Reconciliation Calculation</span>
                                </div>
                            </div>
                        </div>
                    )}
                </div>

                {/* Modal Footer */}
                <div className="px-6 py-3.5 border-t border-slate-100 flex items-center justify-between bg-slate-50/50">
                    <button
                        type="button"
                        onClick={onClose}
                        className="px-4 py-2 rounded-xl border border-slate-200 text-slate-600 hover:bg-slate-100 text-xs font-bold transition bg-white cursor-pointer"
                    >
                        Cancel
                    </button>
                    {step === 2 && (
                        <button
                            type="button"
                            onClick={handleConfirmSave}
                            className="px-5 py-2 rounded-xl bg-sky-600 hover:bg-sky-700 text-white text-xs font-extrabold transition shadow-sm border-none cursor-pointer flex items-center gap-1.5"
                        >
                            <span className="material-symbols-outlined text-[16px]">check_circle</span>
                            <span>Confirm & Reconcile Bill</span>
                        </button>
                    )}
                </div>
            </div>
        </div>
    );
}
