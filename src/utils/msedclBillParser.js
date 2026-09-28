// UtilitySense Smart MSEDCL & Utility Bill Parser
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
        startDate: "",
        endDate: "",
        consumerNo: "",
        consumerName: "",
        tariffCategory: "",
        contractDemandKva: null,
        billedDemandKva: null,
        recordedDemandKva: null,
        billedUnitsKvah: null,
        billedUnitsKwh: null,
        grossUnitsKwh: null,
        totalBillAmount: null,
        energyCharges: null,
        demandCharges: null,
        todCharges: null,
        wheelingCharges: null,
        facCharges: null,
        electricityDuty: null,
        taxOnSale: null,
        gridSupportCharges: null,
        subsidiesTotal: null,
        solarGenUnits: null,
        solarAdjUnits: null,
        openingMeter: null,
        closingMeter: null,
        multiplyingFactor: null,
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
            res.billMonth = `${mMatch[1].toUpperCase()}-${yr}`;
            const lastDay = new Date(yr, Number(mm), 0).getDate();
            res.startDate = `${yr}-${mm}-01`;
            res.endDate = `${yr}-${mm}-${String(lastDay).padStart(2, '0')}`;
        }
    }

    // 2. Consumer Number
    const cNo = rawText.match(/Consumer No\.?\s*:\s*(\d{10,15})/i) || clean.match(/Consumer No[^\d]*(\d{10,15})/i);
    if (cNo) res.consumerNo = cNo[1];

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

    // 4. Financial: Total Payable Bill Amount
    const totMatch = clean.match(/After PPD upto Due Date\s*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/Total Bill Amount Payable Rs\.[^\d]*(\d+(?:\.\d+)?)/i) ||
                     clean.match(/TOTAL CURRENT BILL AS PER TARIFF\s*(\d+(?:\.\d+)?)/i);
    if (totMatch) res.totalBillAmount = Number(totMatch[1]);

    // 5. Units (kVAh and kWh)
    const kvahMatch = clean.match(/KVAH\s*(\d+(?:\.\d+)?)/i) || clean.match(/Industrial\s*(\d+)\s*8\.44/i);
    if (kvahMatch) res.billedUnitsKvah = Number(kvahMatch[1]);

    const kwhMatch = clean.match(/Total Consumption\s*(\d+(?:\.\d+)?)/i);
    if (kwhMatch) res.billedUnitsKwh = Number(kwhMatch[1]);

    const grossMatch = clean.match(/Consumption\s*(\d+(?:\.\d+)?)\s*(?:L\.T|RKVAH)/i) || clean.match(/Consumption\s*(\d{4,8}\.\d+)/i);
    if (grossMatch) res.grossUnitsKwh = Number(grossMatch[1]);

    // 6. Contract & Billing Demand
    const bdMatch = clean.match(/Billed Demand\s*\(KVA\)[^\d]*(\d{2,5})/i);
    if (bdMatch) res.billedDemandKva = Number(bdMatch[1]);

    const recMdMatch = clean.match(/Recorded MD[^\d]*(\d{2,5})/i);
    if (recMdMatch) res.recordedDemandKva = Number(recMdMatch[1]);

    const cdMatch = clean.match(/Contract Demand\s*\(KVA\)[^\d]*(\d{2,5}(?:\.\d+)?)/i);
    if (cdMatch) res.contractDemandKva = Number(cdMatch[1]);

    // 7. Charges Breakdown
    const demChg = clean.match(/Demand Charges.*?Rs\.?[\d.]+\s*(\d+(?:\.\d+)?)/i) || clean.match(/Demand Charges\s*@\s*Rs\.?\d+\s*(\d+(?:\.\d+)?)/i);
    if (demChg) res.demandCharges = Number(demChg[1]);

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

    // Subsidies & Rebates
    const subMatch = clean.match(/Subsidy from Govt of Maharashtra.*?(\d+(?:\.\d+)?)/i);
    const ppdMatch = clean.match(/PROMPT PAYMENT DISCOUNT\s*(\d+(?:\.\d+)?)/i) || clean.match(/Prompt Payment Discount\s*-?\s*(\d+(?:\.\d+)?)/i);
    const incMatch = clean.match(/Incremental Consum\. Rebate\s*-?\s*(\d+(?:\.\d+)?)/i);
    
    let subSum = 0;
    if (subMatch) subSum += Number(subMatch[1]);
    if (ppdMatch) subSum += Number(ppdMatch[1]);
    if (incMatch) subSum += Number(incMatch[1]);
    if (subSum > 0) res.subsidiesTotal = subSum;

    // 8. Solar Net Metering
    const solGen = clean.match(/Total Solar Generation Units\s*:\s*(\d+)/i) || clean.match(/TOD SOLAR GENERATION METER.*?(\d{3,7}\.\d+)/i);
    if (solGen) res.solarGenUnits = Number(solGen[1]);

    const solAdj = clean.match(/Adjustment-Solar\s*(-?\d+(?:\.\d+)?)/i);
    if (solAdj) res.solarAdjUnits = Math.abs(Number(solAdj[1]));

    // 9. Meter Readings & Multiplying Factor
    const curRead = clean.match(/Current\s*\d{2}\/\d{2}\/\d{4}\s*(\d+(?:\.\d+)?)/i);
    if (curRead) res.closingMeter = Number(curRead[1]);

    const prevRead = clean.match(/Previous\s*\d{2}\/\d{2}\/\d{4}\s*(\d+(?:\.\d+)?)/i);
    if (prevRead) res.openingMeter = Number(prevRead[1]);

    const mfMatch = clean.match(/Multiplying Factor\s*(\d+(?:\.\d+)?)/i);
    if (mfMatch) res.multiplyingFactor = Number(mfMatch[1]);

    return res;
}
