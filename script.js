// ==========================================
// VARIABLES GLOBALES Y UTILIDADES
// ==========================================
let activeMethod = 'sturges'; 
let globalDatasets = []; // Arreglo que guardará hasta 10 datasets
const MAX_DATASETS = 10;

// Utilidades Numéricas
const cleanNum = (num, decimals = 4) => {
    if (isNaN(num)) return 0;
    const fixed = parseFloat(num.toFixed(decimals));
    return Number.isInteger(fixed) ? fixed : fixed;
};

const getPercentile = (data, p) => {
    const n = data.length;
    const idx = (p / 100) * (n - 1);
    const l = Math.floor(idx);
    return l + 1 >= n ? data[l] : data[l] * (1 - (idx % 1)) + data[l + 1] * (idx % 1);
};

function createStatRow(label, value, formula) {
    return `
        <div class="stat-row">
            <span class="tooltip">${label}<span class="tooltiptext">${formula}</span></span>
            <b>${value}</b>
        </div>
    `;
}

// ==========================================
// CONTROLADORES DE UI (MODOS E INPUTS)
// ==========================================
document.querySelectorAll('input[name="kMethod"]').forEach(r => {
    r.addEventListener('change', (e) => {
        const inputManual = document.getElementById('kManualValue');
        inputManual.disabled = e.target.value !== 'manual';
        if (!inputManual.disabled) inputManual.focus();
    });
});

document.querySelectorAll('input[name="uploadMode"]').forEach(r => {
    r.addEventListener('change', (e) => {
        const fileInput = document.getElementById('fileInput');
        // Si es manual, restringir a 1 archivo
        if(e.target.value === 'manual') {
            fileInput.removeAttribute('multiple');
        } else {
            fileInput.setAttribute('multiple', 'multiple');
        }
    });
});

document.getElementById('processBtn').addEventListener('click', handleProcessClick);
document.getElementById('exportBtn').addEventListener('click', exportAllToExcel);

// ==========================================
// LÓGICA PRINCIPAL DE PROCESAMIENTO
// ==========================================
async function handleProcessClick() {
    const fileInput = document.getElementById('fileInput');
    const uploadMode = document.querySelector('input[name="uploadMode"]:checked').value;
    activeMethod = document.querySelector('input[name="kMethod"]:checked').value;

    if (!fileInput.files.length) return alert("Sube al menos un archivo Excel.");
    if (fileInput.files.length > MAX_DATASETS) return alert(`Límite excedido. Máximo ${MAX_DATASETS} archivos.`);
    
    if (activeMethod === 'manual') {
        const manualK = parseInt(document.getElementById('kManualValue').value);
        if (isNaN(manualK) || manualK < 1) return alert("Ingresa un número de intervalos (k) válido.");
    }

    globalDatasets = []; // Reiniciar lotes
    document.getElementById('resultsArea').innerHTML = '';

    if (uploadMode === 'auto') {
        // MODO AUTO: Leer todos los archivos subidos
        for (let i = 0; i < fileInput.files.length; i++) {
            const raw = await extractNumbersFromFile(fileInput.files[i]);
            
            // Detección Inteligente
            if (raw.length < 5 && fileInput.files.length === 1) {
                const conf = confirm("No se detectó una columna limpia de números. ¿Deseas abrir el archivo y seleccionar los datos visualmente?");
                if(conf) {
                    document.querySelector('input[value="manual"][name="uploadMode"]').checked = true;
                    return openExcelModal(fileInput.files[0]);
                }
            }
            
            if (raw.length > 0) {
                globalDatasets.push(calculateStatsForDataset(raw, `Archivo ${i+1}: ${fileInput.files[i].name}`));
            }
        }
        renderAllDatasets();
    } else {
        // MODO MANUAL: Abrir modal con el primer archivo
        openExcelModal(fileInput.files[0]);
    }
}

function extractNumbersFromFile(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});
            
            let nums = [];
            json.forEach(row => {
                row.forEach(cell => {
                    let num = parseFloat(cell);
                    if (!isNaN(num)) nums.push(num);
                });
            });
            resolve(nums);
        };
        reader.readAsArrayBuffer(file);
    });
}

// ==========================================
// LÓGICA DEL MODAL VISUAL (SELECCIÓN DE RANGOS)
// ==========================================
let preview2DArray = [];
let isDragging = false;
let startCell = null;
let endCell = null;
let savedRangesData = [];

function openExcelModal(file) {
    savedRangesData = [];
    updateRangeCount();
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        preview2DArray = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1, defval: ""});
        
        renderPreviewTable();
        document.getElementById('previewModal').classList.remove('hidden');
    };
    reader.readAsArrayBuffer(file);
}

document.getElementById('closeModalBtn').onclick = () => document.getElementById('previewModal').classList.add('hidden');

function renderPreviewTable() {
    const container = document.getElementById('tableContainer');
    let html = '<table id="interactiveTable">';
    
    preview2DArray.forEach((row, rIdx) => {
        html += '<tr>';
        row.forEach((cell, cIdx) => {
            html += `<td data-r="${rIdx}" data-c="${cIdx}">${cell !== undefined ? cell : ''}</td>`;
        });
        html += '</tr>';
    });
    html += '</table>';
    container.innerHTML = html;

    const table = document.getElementById('interactiveTable');
    table.addEventListener('mousedown', (e) => {
        if(e.target.tagName === 'TD') {
            isDragging = true;
            startCell = { r: parseInt(e.target.dataset.r), c: parseInt(e.target.dataset.c) };
            endCell = startCell;
            highlightSelection();
        }
    });
    table.addEventListener('mouseover', (e) => {
        if(isDragging && e.target.tagName === 'TD') {
            endCell = { r: parseInt(e.target.dataset.r), c: parseInt(e.target.dataset.c) };
            highlightSelection();
        }
    });
    window.addEventListener('mouseup', () => { isDragging = false; });
}

function highlightSelection() {
    const tds = document.querySelectorAll('#interactiveTable td');
    tds.forEach(td => td.classList.remove('cell-selected'));
    
    if(!startCell || !endCell) return;
    
    const minR = Math.min(startCell.r, endCell.r);
    const maxR = Math.max(startCell.r, endCell.r);
    const minC = Math.min(startCell.c, endCell.c);
    const maxC = Math.max(startCell.c, endCell.c);

    for(let r = minR; r <= maxR; r++) {
        for(let c = minC; c <= maxC; c++) {
            const cell = document.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if(cell && !cell.classList.contains('cell-saved')) cell.classList.add('cell-selected');
        }
    }
}

document.getElementById('saveRangeBtn').addEventListener('click', () => {
    if(savedRangesData.length >= MAX_DATASETS) return alert(`Máximo ${MAX_DATASETS} rangos permitidos.`);
    
    const selected = document.querySelectorAll('.cell-selected');
    if(selected.length === 0) return alert("Selecciona un rango arrastrando el ratón primero.");

    let nums = [];
    selected.forEach(td => {
        let val = parseFloat(td.innerText);
        if(!isNaN(val)) nums.push(val);
        td.classList.remove('cell-selected');
        td.classList.add('cell-saved');
    });

    if(nums.length === 0) return alert("No hay números válidos en tu selección.");

    savedRangesData.push(nums);
    updateRangeCount();
});

function updateRangeCount() {
    document.getElementById('rangeCount').innerText = `Rangos guardados: ${savedRangesData.length}/${MAX_DATASETS}`;
}

document.getElementById('finishRangesBtn').addEventListener('click', () => {
    if(savedRangesData.length === 0) return alert("No has guardado ningún rango.");
    
    globalDatasets = savedRangesData.map((raw, index) => calculateStatsForDataset(raw, `Conjunto Seleccionado ${index + 1}`));
    document.getElementById('previewModal').classList.add('hidden');
    renderAllDatasets();
});


// ==========================================
// MOTOR ESTADÍSTICO POR LOTE
// ==========================================
function calculateStatsForDataset(raw, datasetName) {
    let data = [...raw].sort((a, b) => a - b);
    const n = data.length;
    const minVal = data[0];
    const maxVal = data[n - 1];
    const range = maxVal - minVal;
    
    let numClasses = activeMethod === 'manual' ? parseInt(document.getElementById('kManualValue').value) : Math.round(1 + 3.322 * Math.log10(n));
    if (numClasses < 1) numClasses = 1;
    
    const amplitude = range / numClasses;
    let classesData = [];
    let currentMin = minVal;
    let cumulativeFreq = 0;

    for (let i = 0; i < numClasses; i++) {
        let currentMax = currentMin + amplitude;
        let isLast = (i === numClasses - 1);
        if (isLast) currentMax = maxVal; 

        let count = data.filter(x => x >= currentMin && (isLast ? x <= currentMax : x < currentMax)).length;
        let xi = (currentMin + currentMax) / 2;
        cumulativeFreq += count;
        
        classesData.push({ min: currentMin, max: currentMax, isLast, xi, fi: count, Fi: cumulativeFreq, hi: count/n, Hi: cumulativeFreq/n });
        currentMin = currentMax;
    }

    // Estadísticas
    const sum = data.reduce((a, b) => a + b, 0);
    const mean = sum / n;
    const geoMean = Math.exp(data.reduce((s, x) => s + Math.log(x), 0) / n);
    const harMean = n / data.reduce((s, x) => s + (1 / x), 0);
    const median = n % 2 !== 0 ? data[Math.floor(n/2)] : (data[Math.floor(n/2)-1] + data[Math.floor(n/2)]) / 2;

    let freqMap = {}; let maxFreq = 0; let mode = [];
    data.forEach(num => { freqMap[num] = (freqMap[num] || 0) + 1; if (freqMap[num] > maxFreq) maxFreq = freqMap[num]; });
    for (const key in freqMap) if (freqMap[key] === maxFreq) mode.push(Number(key));

    const variance = data.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (n - 1);
    const stdDev = Math.sqrt(variance);
    const cv = (stdDev / mean) * 100;

    let skewness = 0;
    if (n > 2 && stdDev > 0) skewness = (n / ((n - 1) * (n - 2))) * data.reduce((acc, val) => acc + Math.pow((val - mean) / stdDev, 3), 0);

    return { name: datasetName, data, n, minVal, maxVal, range, numClasses, amplitude, classesData, stats: { mean, geoMean, harMean, median, mode, variance, stdDev, cv, skewness, p10: getPercentile(data,10), q1: getPercentile(data,25), q2: getPercentile(data,50), q3: getPercentile(data,75), p90: getPercentile(data,90) }};
}

// ==========================================
// RENDERIZADO HTML MULTIPLE
// ==========================================
function renderAllDatasets() {
    const resultsArea = document.getElementById('resultsArea');
    resultsArea.innerHTML = '<h2>Resultados del Análisis</h2>';
    
    let kFormula = activeMethod === 'sturges' ? 'k ≈ 1 + 3.322 · log₁₀(n)' : 'Manual';
    let methodLabel = activeMethod === 'sturges' ? ' (Sturges)' : ' (Manual)';

    globalDatasets.forEach((ds, index) => {
        let block = document.createElement('div');
        block.className = 'dataset-block';
        
        let freqHtml = `
            <h3>${ds.name} - Tabla de Frecuencias</h3>
            <table>
                <thead>
                    <tr><th>Límite Inf. (Li)</th><th>Límite Sup. (Ls)</th><th>Marca de Clase (Xi)</th><th>Frec. Abs. (fi)</th><th>Frec. Acum. (Fi)</th><th>Frec. Rel. (hi)</th><th>Frec. Rel. Acum. (Hi)</th></tr>
                </thead>
                <tbody>
        `;
        
        ds.classesData.forEach(c => {
            freqHtml += `<tr><td>${cleanNum(c.min)}</td><td>${cleanNum(c.max)}</td><td>${cleanNum(c.xi)}</td><td>${c.fi}</td><td>${c.Fi}</td><td>${cleanNum(c.hi)}</td><td>${cleanNum(c.Hi)}</td></tr>`;
        });
        freqHtml += `</tbody></table>`;

        let statsHtml = `
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>Parámetros Base</h3>
                    ${createStatRow('Mínimo:', cleanNum(ds.minVal), 'min(xᵢ)')}
                    ${createStatRow('Máximo:', cleanNum(ds.maxVal), 'max(xᵢ)')}
                    ${createStatRow(`Intervalos (k)${methodLabel}:`, ds.numClasses, kFormula)}
                    ${createStatRow('Amplitud (A):', cleanNum(ds.amplitude), 'A = Rango / k')}
                </div>
                <div class="stat-card">
                    <h3>Tendencia Central</h3>
                    ${createStatRow('Media Arit.:', cleanNum(ds.stats.mean), 'x̄ = (Σxᵢ) / n')}
                    ${createStatRow('Media Geom.:', cleanNum(ds.stats.geoMean), 'MG = ⁿ√(x₁···xₙ)')}
                    ${createStatRow('Media Arm.:', cleanNum(ds.stats.harMean), 'MH = n / Σ(1/xᵢ)')}
                    ${createStatRow('Mediana:', cleanNum(ds.stats.median), 'Me')}
                    ${createStatRow('Moda:', ds.stats.mode.map(m=>cleanNum(m)).join(','), 'Mo')}
                </div>
                <div class="stat-card">
                    <h3>Dispersión y Forma</h3>
                    ${createStatRow('Rango:', cleanNum(ds.range), 'R = x_max - x_min')}
                    ${createStatRow('Varianza:', cleanNum(ds.stats.variance), 's²')}
                    ${createStatRow('Desv. Est.:', cleanNum(ds.stats.stdDev), 's = √s²')}
                    ${createStatRow('C. Variación:', cleanNum(ds.stats.cv, 2) + '%', 'CV = (s/x̄)·100%')}
                    ${createStatRow('Asimetría:', cleanNum(ds.stats.skewness), 'As')}
                </div>
                <div class="stat-card">
                    <h3>Posición (Percentiles)</h3>
                    ${createStatRow('P10 (10%):', cleanNum(ds.stats.p10), 'P₁₀')}
                    ${createStatRow('Q1 (25%):', cleanNum(ds.stats.q1), 'Q₁')}
                    ${createStatRow('Q2 (50%):', cleanNum(ds.stats.q2), 'Q₂')}
                    ${createStatRow('Q3 (75%):', cleanNum(ds.stats.q3), 'Q₃')}
                    ${createStatRow('P90 (90%):', cleanNum(ds.stats.p90), 'P₉₀')}
                </div>
            </div>
        `;

        block.innerHTML = freqHtml + statsHtml;
        resultsArea.appendChild(block);
    });

    resultsArea.classList.remove('hidden');
    document.getElementById('exportBtn').classList.remove('hidden');
}

// ==========================================
// EXPORTACIÓN MÚLTIPLE A EXCELJS
// ==========================================
async function exportAllToExcel() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Generador Estadístico Lotes';

    const headerStyle = {
        font: { bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } },
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    };

    globalDatasets.forEach((ds, idx) => {
        let shortName = ds.name.substring(0, 20).replace(/[:*?/"<>|]/g, ''); // Limpiar para tab de excel
        
        // Hoja de Datos
        const wsDatos = wb.addWorksheet(`D_${idx+1}_${shortName}`);
        wsDatos.getCell('A1').value = "DATOS ORDENADOS";
        wsDatos.getCell('A1').font = headerStyle.font; wsDatos.getCell('A1').fill = headerStyle.fill;
        ds.data.forEach((val, i) => { wsDatos.getCell(`A${i + 2}`).value = val; });
        wsDatos.getColumn('A').width = 20;
        
        const dataRange = `'D_${idx+1}_${shortName}'!A2:A${ds.n + 1}`;

        // Hoja de Análisis
        const ws = wb.addWorksheet(`A_${idx+1}_${shortName}`);
        const headers = ['LÍMITE INF. (LI)', 'LÍMITE SUP. (LS)', 'MARCA DE CLASE (XI)', 'FREC. ABS (FI)', 'FREC. ACUM (FI)', 'FREC. REL (HI)', 'FREC. REL ACUM (HI)'];
        ws.addRow(headers);
        ws.getRow(1).eachCell((cell) => { Object.assign(cell, headerStyle); });

        ds.classesData.forEach((cls, i) => {
            let rowNum = i + 2; 
            let cond = cls.isLast ? "<=" : "<";
            ws.addRow([
                cls.min, cls.max,
                { formula: `(A${rowNum}+B${rowNum})/2`, result: cls.xi },
                { formula: `COUNTIFS(${dataRange},">="&A${rowNum},${dataRange},"${cond}"&B${rowNum})`, result: cls.fi },
                i === 0 ? { formula: `D2`, result: cls.Fi } : { formula: `E${rowNum - 1}+D${rowNum}`, result: cls.Fi },
                { formula: `D${rowNum}/COUNT(${dataRange})`, result: cls.hi },
                i === 0 ? { formula: `F2`, result: cls.Hi } : { formula: `G${rowNum - 1}+F${rowNum}`, result: cls.Hi }
            ]).eachCell(cell => cell.alignment = { horizontal: 'center' });
        });

        for(let c = 1; c <= 8; c++) ws.getColumn(c).width = 20;

        let startRow = ds.numClasses + 4;
        ws.getCell(`A${startRow}`).value = "PARÁMETROS"; ws.getCell(`C${startRow}`).value = "TENDENCIA C."; 
        ws.getCell(`E${startRow}`).value = "DISPERSIÓN"; ws.getCell(`G${startRow}`).value = "POSICIÓN";
        [ws.getCell(`A${startRow}`), ws.getCell(`C${startRow}`), ws.getCell(`E${startRow}`), ws.getCell(`G${startRow}`)].forEach(c => { c.font = { bold: true }; c.border = { bottom: { style: 'medium' } }; });

        let formulaK = activeMethod === 'sturges' ? `ROUND(1+3.322*LOG10(COUNT(${dataRange})),0)` : ds.numClasses;
        let filaK = startRow + 4; 
        let formAmp = `(MAX(${dataRange})-MIN(${dataRange}))/B${filaK}`;

        const statsGrid = [
            { c1: 'A', l1: 'Mínimo:', f1: `MIN(${dataRange})`, c2: 'C', l2: 'Media Arit.:', f2: `AVERAGE(${dataRange})`, c3: 'E', l3: 'Rango:', f3: `MAX(${dataRange})-MIN(${dataRange})`, c4: 'G', l4: 'P10:', f4: `PERCENTILE(${dataRange}, 0.1)` },
            { c1: 'A', l1: 'Máximo:', f1: `MAX(${dataRange})`, c2: 'C', l2: 'Media Geom.:', f2: `GEOMEAN(${dataRange})`, c3: 'E', l3: 'Varianza:', f3: `VAR(${dataRange})`, c4: 'G', l4: 'Q1 (25%):', f4: `QUARTILE(${dataRange}, 1)` },
            { c1: 'A', l1: `Int. (k):`, f1: formulaK, c2: 'C', l2: 'Media Arm.:', f2: `HARMEAN(${dataRange})`, c3: 'E', l3: 'Desv. Est.:', f3: `STDEV(${dataRange})`, c4: 'G', l4: 'Q2 (50%):', f4: `MEDIAN(${dataRange})` },
            { c1: 'A', l1: 'Amplitud:', f1: formAmp, c2: 'C', l2: 'Mediana:', f2: `MEDIAN(${dataRange})`, c3: 'E', l3: 'CV:', f3: `STDEV(${dataRange})/AVERAGE(${dataRange})`, c4: 'G', l4: 'Q3 (75%):', f4: `QUARTILE(${dataRange}, 3)` },
            { c1: 'A', l1: '', f1: '', c2: 'C', l2: 'Moda:', f2: `MODE(${dataRange})`, c3: 'E', l3: 'Asimetría:', f3: `SKEW(${dataRange})`, c4: 'G', l4: 'P90:', f4: `PERCENTILE(${dataRange}, 0.9)` }
        ];

        statsGrid.forEach((st, i) => {
            let r = startRow + 2 + i;
            if(st.l1) { ws.getCell(`${st.c1}${r}`).value = st.l1; ws.getCell(`B${r}`).value = (st.c1 === 'A' && r === filaK && activeMethod === 'manual') ? st.f1 : { formula: st.f1 }; }
            if(st.l2) { ws.getCell(`${st.c2}${r}`).value = st.l2; ws.getCell(`D${r}`).value = { formula: st.f2 }; }
            if(st.l3) { ws.getCell(`${st.c3}${r}`).value = st.l3; ws.getCell(`F${r}`).value = { formula: st.f3 }; if(st.l3 === 'CV:') ws.getCell(`F${r}`).numFmt = '0.00%'; }
            if(st.l4) { ws.getCell(`${st.c4}${r}`).value = st.l4; ws.getCell(`H${r}`).value = { formula: st.f4 }; }
        });
    });

    const buffer = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'Analisis_Lotes_Avanzado.xlsx');
}
