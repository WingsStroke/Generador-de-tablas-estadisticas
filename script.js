let rawData = [];
let numClasses = 0;
let classesData = [];
let minVal = 0;
let maxVal = 0;
let amplitude = 0;
let activeMethod = 'sturges'; 

const cleanNum = (num, decimals = 4) => {
    if (isNaN(num)) return 0;
    const fixed = parseFloat(num.toFixed(decimals));
    return Number.isInteger(fixed) ? fixed : fixed;
};

// Activar/desactivar el input manual
document.querySelectorAll('input[name="kMethod"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
        const inputManual = document.getElementById('kManualValue');
        if (e.target.value === 'manual') {
            inputManual.disabled = false;
            inputManual.focus();
        } else {
            inputManual.disabled = true;
        }
    });
});

document.getElementById('processBtn').addEventListener('click', processExcel);
document.getElementById('exportBtn').addEventListener('click', exportToAdvancedExcel);

function processExcel() {
    const fileInput = document.getElementById('fileInput');
    if (!fileInput.files.length) {
        alert("Sube un archivo Excel primero."); return;
    }

    activeMethod = document.querySelector('input[name="kMethod"]:checked').value;
    if (activeMethod === 'manual') {
        const manualK = parseInt(document.getElementById('kManualValue').value);
        if (isNaN(manualK) || manualK < 1) {
            alert("Por favor, ingresa un número de intervalos válido (mayor a 0).");
            return;
        }
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        
        rawData = [];
        json.forEach(row => {
            row.forEach(cell => {
                let num = parseFloat(cell);
                if (!isNaN(num)) rawData.push(num);
            });
        });

        if (rawData.length === 0) return alert("No se encontraron números en el archivo.");

        rawData.sort((a, b) => a - b);
        generateFrequencyTable();
        renderWebStats();
        
        document.getElementById('resultsArea').classList.remove('hidden');
        document.getElementById('exportBtn').classList.remove('hidden');
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
}

function generateFrequencyTable() {
    const n = rawData.length;
    minVal = rawData[0];
    maxVal = rawData[n - 1];
    const range = maxVal - minVal;
    
    if (activeMethod === 'manual') {
        numClasses = parseInt(document.getElementById('kManualValue').value);
    } else {
        numClasses = Math.round(1 + 3.322 * Math.log10(n));
        if (numClasses < 1) numClasses = 1;
    }
    
    amplitude = range / numClasses;
    const tbody = document.querySelector('#freqTable tbody');
    tbody.innerHTML = '';
    classesData = [];

    let currentMin = minVal;
    let cumulativeFreq = 0;

    for (let i = 0; i < numClasses; i++) {
        let currentMax = currentMin + amplitude;
        let isLast = (i === numClasses - 1);
        if (isLast) currentMax = maxVal; 

        let count = rawData.filter(x => x >= currentMin && (isLast ? x <= currentMax : x < currentMax)).length;
        let xi = (currentMin + currentMax) / 2;
        cumulativeFreq += count;
        let hi = count / n;
        let Hi = cumulativeFreq / n;

        classesData.push({ min: currentMin, max: currentMax, isLast, xi, fi: count, Fi: cumulativeFreq, hi, Hi });

        tbody.innerHTML += `
            <tr>
                <td>${cleanNum(currentMin)}</td>
                <td>${cleanNum(currentMax)}</td>
                <td>${cleanNum(xi)}</td>
                <td>${count}</td>
                <td>${cumulativeFreq}</td>
                <td>${cleanNum(hi)}</td>
                <td>${cleanNum(Hi)}</td>
            </tr>
        `;
        currentMin = currentMax;
    }
}

function createStatRow(label, value, formula) {
    return `
        <div class="stat-row">
            <span class="tooltip">${label}
                <span class="tooltiptext">${formula}</span>
            </span>
            <b>${value}</b>
        </div>
    `;
}

function renderWebStats() {
    const n = rawData.length;
    const sum = rawData.reduce((a, b) => a + b, 0);
    const mean = sum / n;
    const geoMean = Math.exp(rawData.reduce((s, x) => s + Math.log(x), 0) / n);
    const harMean = n / rawData.reduce((s, x) => s + (1 / x), 0);
    const mid = Math.floor(n / 2);
    const median = n % 2 !== 0 ? rawData[mid] : (rawData[mid - 1] + rawData[mid]) / 2;

    let freqMap = {}; let maxFreq = 0; let mode = [];
    rawData.forEach(num => {
        freqMap[num] = (freqMap[num] || 0) + 1;
        if (freqMap[num] > maxFreq) maxFreq = freqMap[num];
    });
    for (const key in freqMap) if (freqMap[key] === maxFreq) mode.push(Number(key));

    const variance = rawData.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (n - 1);
    const stdDev = Math.sqrt(variance);
    const range = maxVal - minVal;
    const cv = (stdDev / mean) * 100;

    let skewness = 0;
    if (n > 2 && stdDev > 0) {
        const sum3 = rawData.reduce((acc, val) => acc + Math.pow((val - mean) / stdDev, 3), 0);
        skewness = (n / ((n - 1) * (n - 2))) * sum3;
    }

    const getP = (p) => {
        const idx = (p / 100) * (n - 1);
        const l = Math.floor(idx);
        return l + 1 >= n ? rawData[l] : rawData[l] * (1 - (idx % 1)) + rawData[l + 1] * (idx % 1);
    };

    let methodLabel = activeMethod === 'sturges' ? ' (Sturges)' : ' (Manual)';
    let kFormula = activeMethod === 'sturges' ? 'k ≈ 1 + 3.322 · log₁₀(n)' : 'Asignado por el usuario';

    document.getElementById('statsArea').innerHTML = `
        <div class="stats-grid">
            <div class="stat-card">
                <h3>Parámetros Base</h3>
                ${createStatRow('Valor Mínimo:', cleanNum(minVal), 'min(xᵢ)')}
                ${createStatRow('Valor Máximo:', cleanNum(maxVal), 'max(xᵢ)')}
                ${createStatRow(`Intervalos (k)${methodLabel}:`, numClasses, kFormula)}
                ${createStatRow('Amplitud (A):', cleanNum(amplitude), 'A = Rango / k')}
            </div>
            <div class="stat-card">
                <h3>Tendencia Central</h3>
                ${createStatRow('Media Aritmética:', cleanNum(mean), 'x̄ = (Σxᵢ) / n')}
                ${createStatRow('Media Geométrica:', cleanNum(geoMean), 'MG = ⁿ√(x₁·x₂···xₙ)')}
                ${createStatRow('Media Armónica:', cleanNum(harMean), 'MH = n / Σ(1/xᵢ)')}
                ${createStatRow('Mediana:', cleanNum(median), 'Me = Lᵢ + A·[(n/2 - Fᵢ₋₁)/fᵢ]')}
                ${createStatRow('Moda:', mode.map(m=>cleanNum(m)).join(', '), 'Mo = Lᵢ + A·[(fᵢ - fᵢ₋₁)/(2fᵢ - fᵢ₋₁ - fᵢ₊₁)]')}
            </div>
            <div class="stat-card">
                <h3>Dispersión y Forma</h3>
                ${createStatRow('Rango:', cleanNum(range), 'R = x_max - x_min')}
                ${createStatRow('Varianza:', cleanNum(variance), 's² = Σ(xᵢ - x̄)² / (n - 1)')}
                ${createStatRow('Desv. Estándar:', cleanNum(stdDev), 's = √s²')}
                ${createStatRow('Coef. Variación (CV):', cleanNum(cv, 2) + '%', 'CV = (s / x̄) · 100%')}
                ${createStatRow('Asimetría:', cleanNum(skewness), 'As = [n/((n-1)(n-2))] · Σ[(xᵢ-x̄)/s]³')}
            </div>
            <div class="stat-card">
                <h3>Posición (Percentiles)</h3>
                ${createStatRow('P10 (10%):', cleanNum(getP(10)), 'P₁₀ = Lᵢ + A·[(10n/100 - Fᵢ₋₁)/fᵢ]')}
                ${createStatRow('Q1 (25%):', cleanNum(getP(25)), 'Q₁ = P₂₅')}
                ${createStatRow('Q2 (50%):', cleanNum(getP(50)), 'Q₂ = P₅₀ = Mediana')}
                ${createStatRow('Q3 (75%):', cleanNum(getP(75)), 'Q₃ = P₇₅')}
                ${createStatRow('P90 (90%):', cleanNum(getP(90)), 'P₉₀ = Lᵢ + A·[(90n/100 - Fᵢ₋₁)/fᵢ]')}
            </div>
        </div>
    `;
}

async function exportToAdvancedExcel() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Generador Estadístico';

    const wsDatos = wb.addWorksheet('Datos Base');
    wsDatos.getCell('A1').value = "DATOS ORDENADOS";
    wsDatos.getCell('A1').font = { bold: true, color: { argb: 'FFFFFFFF' } };
    wsDatos.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
    
    rawData.forEach((val, i) => { wsDatos.getCell(`A${i + 2}`).value = val; });
    wsDatos.getColumn('A').width = 25;
    const n = rawData.length;
    const dataRange = `'Datos Base'!A2:A${n + 1}`;

    const ws = wb.addWorksheet('Análisis');
    
    const headerStyle = {
        font: { bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } },
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    };

    const headers = ['LÍMITE INF. (LI)', 'LÍMITE SUP. (LS)', 'MARCA DE CLASE (XI)', 'FREC. ABSOLUTA (FI)', 'FREC. ACUMULADA (FI)', 'FREC. RELATIVA (HI)', 'FREC. REL. ACUM. (HI)'];
    ws.addRow(headers);
    ws.getRow(1).eachCell((cell) => { Object.assign(cell, headerStyle); });

    classesData.forEach((cls, index) => {
        let rowNum = index + 2; 
        let conditionOperator = cls.isLast ? "<=" : "<";

        const row = ws.addRow([
            cls.min, cls.max,
            { formula: `(A${rowNum}+B${rowNum})/2`, result: cls.xi },
            { formula: `COUNTIFS(${dataRange},">="&A${rowNum},${dataRange},"${conditionOperator}"&B${rowNum})`, result: cls.fi },
            index === 0 ? { formula: `D2`, result: cls.Fi } : { formula: `E${rowNum - 1}+D${rowNum}`, result: cls.Fi },
            { formula: `D${rowNum}/COUNT(${dataRange})`, result: cls.hi },
            index === 0 ? { formula: `F2`, result: cls.Hi } : { formula: `G${rowNum - 1}+F${rowNum}`, result: cls.Hi }
        ]);
        row.eachCell(cell => cell.alignment = { horizontal: 'center' });
    });

    for(let i = 1; i <= 8; i++) { ws.getColumn(i).width = 24; }

    let startRow = numClasses + 4;
    ws.getCell(`A${startRow}`).value = "PARÁMETROS BASE";
    ws.getCell(`C${startRow}`).value = "TENDENCIA CENTRAL";
    ws.getCell(`E${startRow}`).value = "DISPERSIÓN Y FORMA";
    ws.getCell(`G${startRow}`).value = "POSICIÓN";
    
    [ws.getCell(`A${startRow}`), ws.getCell(`C${startRow}`), ws.getCell(`E${startRow}`), ws.getCell(`G${startRow}`)].forEach(cell => {
        cell.font = { bold: true }; cell.border = { bottom: { style: 'medium' } };
    });

    let formulaK = activeMethod === 'sturges' ? `ROUND(1+3.322*LOG10(COUNT(${dataRange})),0)` : numClasses;
    let filaK = startRow + 4; 
    let formulaAmplitud = `(MAX(${dataRange})-MIN(${dataRange}))/B${filaK}`;

    const stats = [
        { col1: 'A', lbl1: 'Valor Mínimo:', f1: `MIN(${dataRange})`, col2: 'C', lbl2: 'Media Aritmética:', f2: `AVERAGE(${dataRange})`, col3: 'E', lbl3: 'Rango:', f3: `MAX(${dataRange})-MIN(${dataRange})`, col4: 'G', lbl4: 'P10 (10%):', f4: `PERCENTILE(${dataRange}, 0.1)` },
        { col1: 'A', lbl1: 'Valor Máximo:', f1: `MAX(${dataRange})`, col2: 'C', lbl2: 'Media Geométrica:', f2: `GEOMEAN(${dataRange})`, col3: 'E', lbl3: 'Varianza:', f3: `VAR(${dataRange})`, col4: 'G', lbl4: 'Q1 (25%):', f4: `QUARTILE(${dataRange}, 1)` },
        { col1: 'A', lbl1: `N° Intervalos (k):`, f1: formulaK, col2: 'C', lbl2: 'Media Armónica:', f2: `HARMEAN(${dataRange})`, col3: 'E', lbl3: 'Desv. Estándar:', f3: `STDEV(${dataRange})`, col4: 'G', lbl4: 'Q2 (50%):', f4: `MEDIAN(${dataRange})` },
        { col1: 'A', lbl1: 'Amplitud (A):', f1: formulaAmplitud, col2: 'C', lbl2: 'Mediana:', f2: `MEDIAN(${dataRange})`, col3: 'E', lbl3: 'Coef. Variación (CV):', f3: `STDEV(${dataRange})/AVERAGE(${dataRange})`, col4: 'G', lbl4: 'Q3 (75%):', f4: `QUARTILE(${dataRange}, 3)` },
        { col1: 'A', lbl1: '', f1: '', col2: 'C', lbl2: 'Moda:', f2: `MODE(${dataRange})`, col3: 'E', lbl3: 'Asimetría:', f3: `SKEW(${dataRange})`, col4: 'G', lbl4: 'P90 (90%):', f4: `PERCENTILE(${dataRange}, 0.9)` }
    ];

    stats.forEach((stat, i) => {
        let r = startRow + 2 + i;
        if(stat.lbl1) { ws.getCell(`${stat.col1}${r}`).value = stat.lbl1; ws.getCell(`B${r}`).value = (stat.col1 === 'A' && r === filaK && activeMethod === 'manual') ? stat.f1 : { formula: stat.f1 }; }
        if(stat.lbl2) { ws.getCell(`${stat.col2}${r}`).value = stat.lbl2; ws.getCell(`D${r}`).value = { formula: stat.f2 }; }
        if(stat.lbl3) { 
            ws.getCell(`${stat.col3}${r}`).value = stat.lbl3; 
            ws.getCell(`F${r}`).value = { formula: stat.f3 }; 
            if (stat.lbl3 === 'Coef. Variación (CV):') ws.getCell(`F${r}`).numFmt = '0.00%';
        }
        if(stat.lbl4) { ws.getCell(`${stat.col4}${r}`).value = stat.lbl4; ws.getCell(`H${r}`).value = { formula: stat.f4 }; }
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'Analisis_Estadistico_Avanzado.xlsx');
}
