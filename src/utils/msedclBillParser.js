// UtilitySense Smart MSEDCL & Industrial Utility Bill Parser
// Auto-extracts billing parameters from Maharashtra State Electricity Distribution (MSEDCL) & industrial utility bills

export async function extractTextFromPdfBuffer(arrayBuffer) {
    let fullText = "";
    
    // Primary: Try PDF.js
    try {
        const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.js');
        if (!pdfjsLib.GlobalWorkerOptions.workerSrc && typeof window !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version || '3.11.174'}/pdf.worker.min.js`;
        }
        const loadingTask = pdfjsLib.getDocument({
            data: arrayBuffer,
            disableFontFace: true,
            isEvalSupported: false,
            useSystemFonts: true
        });
        const pdf = await loadingTask.promise;
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            const pageText = content.items.map(item => item.str).join(" ");
            fullText += `\n--- PAGE ${i} ---\n` + pageText;
        }
        if (fullText.trim().length > 50) {
            return fullText;
        }
    } catch (err) {
        console.warn("PDF.js parse failed or worker restricted, using raw stream fallback:", err);
    }

    // Secondary / Fallback: Extract ASCII strings from PDF binary stream directly
    try {
        const bytes = new Uint8Array(arrayBuffer);
        let str = "";
        for (let i = 0; i < bytes.length; i++) {
            const c = bytes[i];
            if ((c >= 32 && c <= 126) || c === 10 || c === 13) {
                str += String.fromCharCode(c);
            }
        }
        return str;
    } catch (fallbackErr) {
        console.error("Binary stream extract failed:", fallbackErr);
        return "";
    }
}

export function parseMsedclBillText(rawText) {
    if (!rawText || typeof rawText !== "string") return null;

    const clean = rawText.replace(/,/g, '');
    const res = {
        rawDetected: true,
        billMonth: "",
        billMonthDisplay: "",
        startDate: "",
        endDate: "",
        billDate: "",
        consumerNo: "",
        consumerName: "",
        meterNo: "",
        tariffCategory: "",
        
        // Demand Parameters
        contractDemandKva: null,
        billedDemandKva: null,
        recordedDemandKva: null,
        kwMaxDemand: null,
        kvaMaxDemand: null,
        powerFactor: null,

        // Consumption Parameters
        billedUnitsKvah: null,
        billedUnitsKwh: null,
        grossUnitsKwh: null,
        openingMeter: null,
        closingMeter: null,
        meterDifference: null,
        multiplyingFactor: null,
        rkvahLag: null,
        rkvahLead: null,
        
        // TOD Slots (kWh)
        todSlotA: null, // 0000-0600 & 2200-2400
        todSlotB: null, // 0600-0900 & 1200-1800
        todSlotC: null, // 0900-1200
        todSlotD: null, // 1800-2200

        // Solar Reconciliation Parameters
        solarGenUnits: null,
        solarExportUnits: null,
        solarAdjUnits: null,
        solarCapacity: null,
        solarMeterReading: null,
        solarGenMeterConsumption: null,

        // Financial & Cost Breakdown
        energyCharges: null,
        demandCharges: null,
        excessDemandCharges: null,
        todCharges: null,
        wheelingCharges: null,
        facCharges: null,
        electricityDuty: null,
        taxOnSale: null,
        gridSupportCharges: null,
        rebatesTotal: null,
        subsidiesTotal: null,
        promptPaymentDiscount: null,
        otherCharges: null,
        currentBillAmount: null,
        totalBillAmount: null, // Total payable amount

        // Inferred location/plant mappings
        suggestedPlantCode: "",
        suggestedLocation: ""
    };

    // 1. Bill Month & Period
    const mMatch = rawText.match(/MONTH\s+OF\s+([A-Z]{3})[- ]?(\d{4})/i) ||
                   rawText.match(/Bill Month\s*:\s*([A-Z]{3})[- ]?(\d{4})/i) ||
                   rawText.match(/BILL OF SUPPLY FOR THE MONTH OF\s+([A-Z]{3})[- ]?(\d{4})/i);
    if (mMatch) {
        const months = { JAN:'01', FEB:'02', MAR:'03', APR:'04', MAY:'05', JUN:'06', JUL:'07', AUG:'08', SEP:'09', OCT:'10', NOV:'11', DEC:'12' };
        const mm = months[mMatch[1].toUpperCase()];
        const yr = mMatch[2];
        if (mm && yr) {
            res.billMonth = `${yr}-${mm}`;
            res.billMonthDisplay = `${mMatch[1].toUpperCase()} ${yr}`;
            const lastDay = new Date(yr, Number(mm), 0).getDate();
            res.startDate = `${yr}-${mm}-01`;
            res.endDate = `${yr}-${mm}-${String(lastDay).padStart(2, '0')}`;
        }
    }

    // Bill Date
    const bDate = rawText.match(/Bill Date\s*:\s*(\d{2}[-/]\d{2}[-/]\d{4})/i) ||
                  rawText.match(/Date\s*:\s*(\d{2}[-/]\d{2}[-/]\d{4})/i);
    if (bDate) {
        const parts = bDate[1].split(/[-/]/);
        if (parts.length === 3) {
            res.billDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
        }
    }

    // 2. Consumer Number
    const cNo = rawText.match(/Consumer No\.?\s*:\s*(\d{10,15})/i) || clean.match(/Consumer No[^\d]*(\d{10,15})/i);
    if (cNo) res.consumerNo = cNo[1];

    // Meter Number
    const mNo = rawText.match(/Meter No\.?\s*:\s*([A-Za-z0-9\-_]+)/i) || clean.match(/Meter Serial No\.?\s*:\s*([A-Za-z0-9\-_]+)/i);
    if (mNo) res.meterNo = mNo[1];

    // 3. Consumer Name & Plant Mapping
    const cName = rawText.match(/Consumer Name\s*:\s*([^\n\r]+)/i);
    if (cName) {
        res.consumerName = cName[1].trim();
        const upper = res.consumerName.toUpperCase();
        if (upper.includes("NEXT GEN") || upper.includes("NGM")) {
            res.suggestedPlantCode = "4010";
            res.suggestedLocation = "PUNE";
        } else if (upper.includes("TECHNOPLAST") || upper.includes("PGTL")) {
            res.suggestedPlantCode = "2020";
            res.suggestedLocation = "PUNE";
        } else if (upper.includes("ELECTROPLAST")) {
            res.suggestedPlantCode = "1040";
            res.suggestedLocation = "PUNE";
        }
    }

    // Tariff Category
    const tarMatch = rawText.match(/Tariff\s*:\s*([^\n\r,]+)/i);
    if (tarMatch) res.tariffCategory = tarMatch[1].trim();

    // 4. Financial: Total Payable Bill Amount & Current Bill
    const curBillMatch = clean.match(/TOTAL CURRENT BILL(?: AS PER TARIFF)?\s*(\d+(?:\.\d+)?)/i) ||
                         clean.match(/Current Bill Amount\s*[:=]?\s*Rs\.?\s*(\d+(?:\.\d+)?)/i);
    if (curBillMatch) res.currentBillAmount = Number(curBillMatch[1]);

    const totMatch = clean.match(/After PPD upto Due Date\s*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Total Bill Amount Payable Rs\.[^\d]*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Net Payable Amount\s*[:=]?\s*Rs\.?\s*(\d+(?:\.\d+)?)/i) ||
                     curBillMatch;
    if (totMatch) res.totalBillAmount = Number(totMatch[1]);

    // 5. Units (kVAh and kWh)
    const kvahMatch = clean.match(/KVAH\s*(\d+(?:\.\d+)?)/i) || 
                      clean.match(/Billed Units\s*\(?KVAH\)?\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                      clean.match(/Industrial\s*(\d+)\s*8\.44/i);
    if (kvahMatch) res.billedUnitsKvah = Number(kvahMatch[1]);

    const kwhMatch = clean.match(/Total Consumption\s*(\d+(?:\.\d+)?)/i) || 
                     clean.match(/Net Consumption\s*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Billed Units\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Units Billed\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Total Billed Units\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Net Billed Units\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Active Consumption\s*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Consumption\s*\(?KWH\)?\s*[:\s]*(\d+(?:\.\d+)?)/i);
    if (kwhMatch) {
        res.billedUnitsKwh = Number(kwhMatch[1]);
    } else if (res.billedUnitsKvah !== null) {
        // In Maharashtra HT industrial bills, energy charges are billed directly on kVAh
        res.billedUnitsKwh = res.billedUnitsKvah;
    }

    const grossMatch = clean.match(/Gross Units\s*(\d+(?:\.\d+)?)/i) || 
                       clean.match(/Gross Consumption\s*(\d+(?:\.\d+)?)/i) ||
                       clean.match(/Consumption\s*(\d+(?:\.\d+)?)\s*(?:L\.T|RKVAH)/i) || 
                       clean.match(/Consumption\s*(\d{4,8}\.\d+)/i);
    if (grossMatch) {
        res.grossUnitsKwh = Number(grossMatch[1]);
    } else if (res.billedUnitsKwh !== null) {
        res.grossUnitsKwh = res.billedUnitsKwh;
    }

    // Meter Readings & Multiplying Factor
    // 1. Current / Closing Reading
    const curRead = clean.match(/Current\s*(?:Reading|Rdg)?\s*[:\s]*\d{0,2}[-/.]?\d{0,2}[-/.]?\d{0,4}\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                    clean.match(/Current\s*(?:Reading|Rdg)?\s*[:\s]*(\d{4,8}(?:\.\d+)?)/i) ||
                    clean.match(/Closing\s*(?:Reading|Rdg)?\s*[:\s]*(\d{4,8}(?:\.\d+)?)/i) ||
                    clean.match(/Present\s*(?:Reading|Rdg)?\s*[:\s]*(\d{4,8}(?:\.\d+)?)/i) ||
                    clean.match(/Current\s*\d{2}[-/.]\d{2}[-/.]\d{2,4}\s*(\d+(?:\.\d+)?)/i);
    if (curRead) res.closingMeter = Number(curRead[1]);

    // 2. Previous / Opening Reading
    const prevRead = clean.match(/Previous\s*(?:Reading|Rdg)?\s*[:\s]*\d{0,2}[-/.]?\d{0,2}[-/.]?\d{0,4}\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Previous\s*(?:Reading|Rdg)?\s*[:\s]*(\d{4,8}(?:\.\d+)?)/i) ||
                     clean.match(/Opening\s*(?:Reading|Rdg)?\s*[:\s]*(\d{4,8}(?:\.\d+)?)/i) ||
                     clean.match(/Past\s*(?:Reading|Rdg)?\s*[:\s]*(\d{4,8}(?:\.\d+)?)/i) ||
                     clean.match(/Previous\s*\d{2}[-/.]\d{2}[-/.]\d{2,4}\s*(\d+(?:\.\d+)?)/i);
    if (prevRead) res.openingMeter = Number(prevRead[1]);

    // 3. Multi-column meter reading rows: KWH <val1> <val2> <val3>
    const kwhRow = clean.match(/KWH\s+(\d{3,8}(?:\.\d+)?)\s+(\d{3,8}(?:\.\d+)?)\s+(\d{3,8}(?:\.\d+)?)/i);
    if (kwhRow) {
        const v1 = Number(kwhRow[1]);
        const v2 = Number(kwhRow[2]);
        const v3 = Number(kwhRow[3]);
        if (v1 > v2) {
            if (res.closingMeter === null) res.closingMeter = v1;
            if (res.openingMeter === null) res.openingMeter = v2;
            if (res.meterDifference === null) res.meterDifference = v3;
        } else if (v2 > v1) {
            if (res.closingMeter === null) res.closingMeter = v2;
            if (res.openingMeter === null) res.openingMeter = v1;
            if (res.meterDifference === null) res.meterDifference = v3;
        }
    }

    const mfMatch = clean.match(/Multiplying Factor\s*[:\s]*(\d+(?:\.\d+)?)/i) || clean.match(/MF\s*[:\s]*(\d+(?:\.\d+)?)/i);
    if (mfMatch) res.multiplyingFactor = Number(mfMatch[1]);
    else if (res.multiplyingFactor === null) res.multiplyingFactor = 1;

    // Meter Difference calculation
    if (res.closingMeter !== null && res.openingMeter !== null) {
        res.meterDifference = Number(Math.max(0, res.closingMeter - res.openingMeter).toFixed(3));
    } else if (res.meterDifference === null && res.grossUnitsKwh !== null) {
        const mf = res.multiplyingFactor || 1;
        res.meterDifference = Number((res.grossUnitsKwh / mf).toFixed(3));
    }

    // Reactive Energy RKVAH
    const rkLag = clean.match(/RKVAH\s*[\(-]?\s*LAG\s*[\)-]?\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                  clean.match(/RKVAH\s*LAG\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                  clean.match(/Reactive\s*(?:Energy)?\s*[\(-]?\s*Lag\s*[\)-]?\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                  clean.match(/LAG\s*RKVAH\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                  clean.match(/RKVAH\s*(?:Lag)?\s*(\d+(?:\.\d+)?)/i);
    if (rkLag) res.rkvahLag = Number(rkLag[1]);

    const rkLead = clean.match(/RKVAH\s*[\(-]?\s*LEAD\s*[\)-]?\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                   clean.match(/RKVAH\s*LEAD\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                   clean.match(/Reactive\s*(?:Energy)?\s*[\(-]?\s*Lead\s*[\)-]?\s*[:\s]*(\d+(?:\.\d+)?)/i) ||
                   clean.match(/LEAD\s*RKVAH\s*[:\s]*(\d+(?:\.\d+)?)/i);
    if (rkLead) res.rkvahLead = Number(rkLead[1]);

    // 6. Contract & Billing Demand
    const bdMatch = clean.match(/Billed Demand\s*\(KVA\)[^\d]*(\d{2,5})/i);
    if (bdMatch) res.billedDemandKva = Number(bdMatch[1]);

    const recMdMatch = clean.match(/Recorded MD[^\d]*(\d{2,5})/i);
    if (recMdMatch) res.recordedDemandKva = Number(recMdMatch[1]);

    const cdMatch = clean.match(/Contract Demand\s*\(KVA\)[^\d]*(\d{2,5}(?:\.\d+)?)/i);
    if (cdMatch) res.contractDemandKva = Number(cdMatch[1]);

    const kwMd = clean.match(/KW\s*(?:MD|Demand)\s*(\d+(?:\.\d+)?)/i);
    if (kwMd) res.kwMaxDemand = Number(kwMd[1]);

    const kvaMd = clean.match(/KVA\s*(?:MD|Demand)\s*(\d+(?:\.\d+)?)/i);
    if (kvaMd) res.kvaMaxDemand = Number(kvaMd[1]);

    // Power Factor
    const pfMatch = clean.match(/Power Factor\s*\(?PF\)?\s*[:=]?\s*(\d+(?:\.\d+)?)/i) ||
                    clean.match(/P\.?F\.?\s*[:=]?\s*(0\.\d{2,4}|1\.000?)/i);
    if (pfMatch) res.powerFactor = Number(pfMatch[1]);

    // 7. Charges Breakdown
    const demChg = clean.match(/Demand Charges.*?Rs\.?[\d.]+\s*(\d+(?:\.\d+)?)/i) || clean.match(/Demand Charges\s*@\s*Rs\.?\d+\s*(\d+(?:\.\d+)?)/i);
    if (demChg) res.demandCharges = Number(demChg[1]);

    const excDemChg = clean.match(/Excess Demand Charges?\s*(\d+(?:\.\d+)?)/i);
    if (excDemChg) res.excessDemandCharges = Number(excDemChg[1]);

    const nrgChg = clean.match(/Energy Charges\s*(\d+(?:\.\d+)?)/i);
    if (nrgChg) res.energyCharges = Number(nrgChg[1]);

    const todChg = clean.match(/TOD Tariff EC\s*(\d+(?:\.\d+)?)/i);
    if (todChg) res.todCharges = Number(todChg[1]);

    const whlChg = clean.match(/Wheeling Charge.*?Rs\/U\s*(\d+(?:\.\d+)?)/i);
    if (whlChg) res.wheelingCharges = Number(whlChg[1]);

    const facChg = clean.match(/FAC\s*@\s*[\d.]+\s*Ps\.\/U\s*(\d+(?:\.\d+)?)/i);
    if (facChg) res.facCharges = Number(facChg[1]);

    const edChg = clean.match(/Electricity Duty\s*(\d+(?:\.\d+)?)/i);
    if (edChg) res.electricityDuty = Number(edChg[1]);

    const tosChg = clean.match(/Tax on Sale.*?Ps\.\/U\s*(\d+(?:\.\d+)?)/i);
    if (tosChg) res.taxOnSale = Number(tosChg[1]);

    const gsChg = clean.match(/Grid Support Charge\s*(\d+(?:\.\d+)?)/i);
    if (gsChg) res.gridSupportCharges = Number(gsChg[1]);

    // Subsidies, Discounts & Rebates
    const subMatch = clean.match(/Subsidy from Govt of Maharashtra.*?(\d+(?:\.\d+)?)/i);
    if (subMatch) res.subsidiesTotal = Number(subMatch[1]);

    const ppdMatch = clean.match(/PROMPT PAYMENT DISCOUNT\s*(\d+(?:\.\d+)?)/i) || clean.match(/Prompt Payment Discount\s*-?\s*(\d+(?:\.\d+)?)/i);
    if (ppdMatch) res.promptPaymentDiscount = Number(ppdMatch[1]);

    const incMatch = clean.match(/Incremental Consum\. Rebate\s*-?\s*(\d+(?:\.\d+)?)/i);
    if (incMatch) res.rebatesTotal = Number(incMatch[1]);

    // 8. Solar Net Metering Details
    const solGen = clean.match(/Total Solar Generation Units\s*:\s*(\d+)/i) || clean.match(/TOD SOLAR GENERATION METER.*?(\d{3,7}\.\d+)/i);
    if (solGen) res.solarGenUnits = Number(solGen[1]);

    const solAdj = clean.match(/Adjustment-Solar\s*(-?\d+(?:\.\d+)?)/i);
    if (solAdj) res.solarAdjUnits = Math.abs(Number(solAdj[1]));

    const solExp = clean.match(/Solar Export(?: Units)?\s*[:=]?\s*(\d+(?:\.\d+)?)/i) ||
                   clean.match(/Export Units\s*[:=]?\s*(\d+(?:\.\d+)?)/i);
    if (solExp) {
        res.solarExportUnits = Number(solExp[1]);
    } else if (res.solarAdjUnits !== null) {
        res.solarExportUnits = res.solarAdjUnits; // In many MSEDCL solar net metering bills, solar adjustment equals exported units
    }

    const solCap = clean.match(/Solar Capacity\s*\(?KWP?\)?\s*[:=]?\s*(\d+(?:\.\d+)?)/i);
    if (solCap) res.solarCapacity = Number(solCap[1]);

    // TOD Slots
    const todA = clean.match(/TOD\s*Zone\s*A\s*(\d+(?:\.\d+)?)/i) || clean.match(/0000-0600\s*hrs\s*(\d+(?:\.\d+)?)/i);
    if (todA) res.todSlotA = Number(todA[1]);

    const todB = clean.match(/TOD\s*Zone\s*B\s*(\d+(?:\.\d+)?)/i) || clean.match(/0600-0900\s*hrs\s*(\d+(?:\.\d+)?)/i);
    if (todB) res.todSlotB = Number(todB[1]);

    const todC = clean.match(/TOD\s*Zone\s*C\s*(\d+(?:\.\d+)?)/i) || clean.match(/0900-1200\s*hrs\s*(\d+(?:\.\d+)?)/i);
    if (todC) res.todSlotC = Number(todC[1]);

    const todD = clean.match(/TOD\s*Zone\s*D\s*(\d+(?:\.\d+)?)/i) || clean.match(/1800-2200\s*hrs\s*(\d+(?:\.\d+)?)/i);
    if (todD) res.todSlotD = Number(todD[1]);

    return res;
}

/**
 * Builds structured parameters array for electricity & utility bills
 * Each parameter has { id, category, name, value, unit, source, billRef, confidence }
 */
export function getBillParametersList(parsed) {
    if (!parsed) return [];

    const list = [
        // 1. Consumption
        { id: "grossUnitsKwh", category: "Consumption", name: "Gross Consumption", value: parsed.grossUnitsKwh, unit: "kWh", billRef: "Consumption / Gross Units", source: "PDF Extracted", status: parsed.grossUnitsKwh !== null ? "Extracted" : "N/A" },
        { id: "billedUnitsKwh", category: "Consumption", name: "Net Consumption", value: parsed.billedUnitsKwh, unit: "kWh", billRef: "Total Consumption / Net Units", source: "PDF Extracted", status: parsed.billedUnitsKwh !== null ? "Extracted" : "N/A" },
        { id: "billedUnitsKvah", category: "Consumption", name: "Billed Units (kVAh)", value: parsed.billedUnitsKvah, unit: "kVAh", billRef: "KVAH Units", source: "PDF Extracted", status: parsed.billedUnitsKvah !== null ? "Extracted" : "N/A" },
        { id: "openingMeter", category: "Consumption", name: "Previous Meter Reading", value: parsed.openingMeter, unit: "kWh", billRef: "Previous Reading", source: "PDF Extracted", status: parsed.openingMeter !== null ? "Extracted" : "N/A" },
        { id: "closingMeter", category: "Consumption", name: "Current Meter Reading", value: parsed.closingMeter, unit: "kWh", billRef: "Current Reading", source: "PDF Extracted", status: parsed.closingMeter !== null ? "Extracted" : "N/A" },
        { id: "meterDifference", category: "Consumption", name: "Meter Difference", value: parsed.meterDifference, unit: "kWh", billRef: "Current - Previous", source: "Calculated", status: parsed.meterDifference !== null ? "Calculated" : "N/A" },
        { id: "multiplyingFactor", category: "Consumption", name: "Multiplying Factor", value: parsed.multiplyingFactor, unit: "x", billRef: "Multiplying Factor (MF)", source: "PDF Extracted", status: parsed.multiplyingFactor !== null ? "Extracted" : "N/A" },
        { id: "rkvahLag", category: "Consumption", name: "RKVAH Lag", value: parsed.rkvahLag, unit: "RKVAh", billRef: "RKVAH (Lag)", source: "PDF Extracted", status: parsed.rkvahLag !== null ? "Extracted" : "N/A" },
        { id: "rkvahLead", category: "Consumption", name: "RKVAH Lead", value: parsed.rkvahLead, unit: "RKVAh", billRef: "RKVAH (Lead)", source: "PDF Extracted", status: parsed.rkvahLead !== null ? "Extracted" : "N/A" },
        
        // 2. Solar
        { id: "solarGenUnits", category: "Solar", name: "Solar Generation", value: parsed.solarGenUnits, unit: "kWh", billRef: "Total Solar Gen / Generation Meter", source: "PDF Extracted", status: parsed.solarGenUnits !== null ? "Extracted" : "N/A" },
        { id: "solarExportUnits", category: "Solar", name: "Solar Export", value: parsed.solarExportUnits, unit: "kWh", billRef: "Solar Export Units", source: "PDF Extracted", status: parsed.solarExportUnits !== null ? "Extracted" : "N/A" },
        { id: "solarAdjUnits", category: "Solar", name: "Solar Adjustment", value: parsed.solarAdjUnits, unit: "kWh", billRef: "Adjustment - Solar", source: "PDF Extracted", status: parsed.solarAdjUnits !== null ? "Extracted" : "N/A" },
        { id: "solarCapacity", category: "Solar", name: "Solar Capacity", value: parsed.solarCapacity, unit: "kWp", billRef: "Solar Sanctioned Capacity", source: "PDF Extracted", status: parsed.solarCapacity !== null ? "Extracted" : "N/A" },

        // 3. Demand & Power Factor
        { id: "recordedDemandKva", category: "Demand", name: "Recorded MD", value: parsed.recordedDemandKva, unit: "kVA", billRef: "Recorded Maximum Demand", source: "PDF Extracted", status: parsed.recordedDemandKva !== null ? "Extracted" : "N/A" },
        { id: "billedDemandKva", category: "Demand", name: "Billed Demand", value: parsed.billedDemandKva, unit: "kVA", billRef: "Billed Demand (KVA)", source: "PDF Extracted", status: parsed.billedDemandKva !== null ? "Extracted" : "N/A" },
        { id: "contractDemandKva", category: "Demand", name: "Contract Demand", value: parsed.contractDemandKva, unit: "kVA", billRef: "Contract Demand (KVA)", source: "PDF Extracted", status: parsed.contractDemandKva !== null ? "Extracted" : "N/A" },
        { id: "kwMaxDemand", category: "Demand", name: "KW Maximum Demand", value: parsed.kwMaxDemand, unit: "kW", billRef: "KW MD", source: "PDF Extracted", status: parsed.kwMaxDemand !== null ? "Extracted" : "N/A" },
        { id: "powerFactor", category: "Power Factor", name: "Power Factor (PF)", value: parsed.powerFactor, unit: "", billRef: "Average Power Factor", source: "PDF Extracted", status: parsed.powerFactor !== null ? "Extracted" : "N/A" },

        // 4. Cost Breakdown
        { id: "energyCharges", category: "Cost", name: "Energy Charges", value: parsed.energyCharges, unit: "₹", billRef: "Energy Charges", source: "PDF Extracted", status: parsed.energyCharges !== null ? "Extracted" : "N/A" },
        { id: "demandCharges", category: "Cost", name: "Demand Charges", value: parsed.demandCharges, unit: "₹", billRef: "Demand Charges", source: "PDF Extracted", status: parsed.demandCharges !== null ? "Extracted" : "N/A" },
        { id: "excessDemandCharges", category: "Cost", name: "Excess Demand Charges", value: parsed.excessDemandCharges, unit: "₹", billRef: "Excess Demand Charges", source: "PDF Extracted", status: parsed.excessDemandCharges !== null ? "Extracted" : "N/A" },
        { id: "wheelingCharges", category: "Cost", name: "Wheeling Charges", value: parsed.wheelingCharges, unit: "₹", billRef: "Wheeling Charge", source: "PDF Extracted", status: parsed.wheelingCharges !== null ? "Extracted" : "N/A" },
        { id: "facCharges", category: "Cost", name: "FAC (Fuel Adjustment)", value: parsed.facCharges, unit: "₹", billRef: "FAC Charges", source: "PDF Extracted", status: parsed.facCharges !== null ? "Extracted" : "N/A" },
        { id: "electricityDuty", category: "Cost", name: "Electricity Duty", value: parsed.electricityDuty, unit: "₹", billRef: "Electricity Duty", source: "PDF Extracted", status: parsed.electricityDuty !== null ? "Extracted" : "N/A" },
        { id: "todCharges", category: "Cost", name: "TOD Charges", value: parsed.todCharges, unit: "₹", billRef: "TOD Tariff EC", source: "PDF Extracted", status: parsed.todCharges !== null ? "Extracted" : "N/A" },
        { id: "gridSupportCharges", category: "Cost", name: "Grid Support Charges", value: parsed.gridSupportCharges, unit: "₹", billRef: "Grid Support Charge", source: "PDF Extracted", status: parsed.gridSupportCharges !== null ? "Extracted" : "N/A" },
        { id: "promptPaymentDiscount", category: "Cost", name: "Prompt Payment Discount (PPD)", value: parsed.promptPaymentDiscount, unit: "₹", billRef: "Prompt Payment Discount", source: "PDF Extracted", status: parsed.promptPaymentDiscount !== null ? "Extracted" : "N/A" },
        { id: "subsidiesTotal", category: "Cost", name: "Govt Subsidies / Rebates", value: parsed.subsidiesTotal, unit: "₹", billRef: "Subsidy from Govt", source: "PDF Extracted", status: parsed.subsidiesTotal !== null ? "Extracted" : "N/A" },
        { id: "currentBillAmount", category: "Cost", name: "Current Bill Amount", value: parsed.currentBillAmount, unit: "₹", billRef: "Current Bill Amount", source: "PDF Extracted", status: parsed.currentBillAmount !== null ? "Extracted" : "N/A" },
        { id: "totalBillAmount", category: "Cost", name: "Payable Bill Amount", value: parsed.totalBillAmount, unit: "₹", billRef: "Total Amount Payable", source: "PDF Extracted", status: parsed.totalBillAmount !== null ? "Extracted" : "N/A" }
    ];

    return list;
}
