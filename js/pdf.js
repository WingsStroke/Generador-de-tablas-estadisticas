import { cleanNum } from './math.js';

export async function exportToPDF(dataset, slideIndex) {
    const element = document.createElement('div');
    element.style.padding = '30px';
    element.style.fontFamily = 'Arial, sans-serif';
    element.style.color = '#000';
    element.style.backgroundColor = '#fff';

    let html = `
        <h1 style="text-align: center; border-bottom: 2px solid #000; padding-bottom: 10px; text-transform: uppercase; font-size: 24px;">
            Reporte Estadístico: ${dataset.name}
        </h1>
        <p style="text-align: center; color: #555; margin-bottom: 30px;">Generado automáticamente por el Generador de Tablas de Frecuencia</p>
        
        <h3 style="border-bottom: 1px solid #ccc; padding-bottom: 5px;">1. Tabla de Frecuencias ${dataset.isGrouped ? '(Datos Agrupados)' : '(Frecuencias Simples)'}</h3>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 30px; text-align: center; font-size: 12px;">
            <thead>
                <tr style="background-color: #000; color: #fff;">
    `;
    
    if (dataset.isGrouped) {
        html += `<th style="padding: 8px; border: 1px solid #000;">Límite Inf. (Li)</th><th style="padding: 8px; border: 1px solid #000;">Límite Sup. (Ls)</th><th style="padding: 8px; border: 1px solid #000;">Marca Clase (Xi)</th>`;
    } else {
        html += `<th style="padding: 8px; border: 1px solid #000;">Dato (Xi)</th>`;
    }
    
    html += `<th style="padding: 8px; border: 1px solid #000;">Frec. Abs. (fi)</th><th style="padding: 8px; border: 1px solid #000;">Frec. Acum. (Fi)</th><th style="padding: 8px; border: 1px solid #000;">Frec. Rel. (hi)</th><th style="padding: 8px; border: 1px solid #000;">Frec. Rel. Acum. (Hi)</th></tr></thead><tbody>`;
    
    dataset.classesData.forEach(c => {
        html += `<tr>`;
        if (dataset.isGrouped) {
            html += `<td style="padding: 6px; border: 1px solid #000;">${cleanNum(c.min)}</td><td style="padding: 6px; border: 1px solid #000;">${cleanNum(c.max)}</td>`;
        }
        html += `<td style="padding: 6px; border: 1px solid #000;">${cleanNum(c.xi)}</td><td style="padding: 6px; border: 1px solid #000;">${c.fi}</td><td style="padding: 6px; border: 1px solid #000;">${c.Fi}</td><td style="padding: 6px; border: 1px solid #000;">${cleanNum(c.hi)}</td><td style="padding: 6px; border: 1px solid #000;">${cleanNum(c.Hi)}</td></tr>`;
    });
    html += `</tbody></table>`;

    html += `
        <h3 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 20px;">2. Medidas Estadísticas</h3>
        <div style="display: flex; justify-content: space-between; font-size: 13px; margin-bottom: 30px; border: 1px solid #000; padding: 15px;">
            <div style="width: 48%;">
                <p><b>Total de datos (n):</b> ${dataset.n}</p>
                <p><b>Mínimo:</b> ${cleanNum(dataset.minVal)}</p>
                <p><b>Máximo:</b> ${cleanNum(dataset.maxVal)}</p>
                <p><b>Rango:</b> ${cleanNum(dataset.range)}</p>
                <p style="margin-top: 10px;"><b>Media Aritmética:</b> ${cleanNum(dataset.stats.mean)}</p>
                <p><b>Mediana:</b> ${cleanNum(dataset.stats.median)}</p>
                <p><b>Moda:</b> ${dataset.stats.mode.map(m=>cleanNum(m)).join(', ')}</p>
            </div>
            <div style="width: 48%;">
                <p><b>Varianza:</b> ${cleanNum(dataset.stats.variance)}</p>
                <p><b>Desviación Estándar:</b> ${cleanNum(dataset.stats.stdDev)}</p>
                <p><b>Coeficiente Variación:</b> ${cleanNum(dataset.stats.cv, 2)}%</p>
                <p><b>Asimetría:</b> ${cleanNum(dataset.stats.skewness)}</p>
                <p style="margin-top: 10px;"><b>Cuartil 1 (Q1):</b> ${cleanNum(dataset.stats.q1)}</p>
                <p><b>Cuartil 2 (Q2):</b> ${cleanNum(dataset.stats.q2)}</p>
                <p><b>Cuartil 3 (Q3):</b> ${cleanNum(dataset.stats.q3)}</p>
            </div>
        </div>
    `;

    const histCanvas = document.getElementById(`chartHist-${slideIndex}`);
    const ojivaCanvas = document.getElementById(`chartOjiva-${slideIndex}`);
    const boxCanvas = document.getElementById(`chartBox-${slideIndex}`);

    html += `<div style="page-break-before: always;"></div>`;
    html += `<h3 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-bottom: 20px;">3. Gráficos Estadísticos</h3>`;
    html += `<div style="text-align: center;">`;
    
    if (histCanvas) {
        html += `<h4 style="margin: 15px 0 10px 0; text-transform: uppercase;">Histograma y Polígono de Frecuencias</h4>`;
        html += `<img src="${histCanvas.toDataURL('image/png', 1.0)}" style="max-width: 90%; height: auto; border: 1px solid #ccc; margin-bottom: 30px;">`;
    }
    if (ojivaCanvas) {
        html += `<h4 style="margin: 15px 0 10px 0; text-transform: uppercase;">Ojiva (Menor que)</h4>`;
        html += `<img src="${ojivaCanvas.toDataURL('image/png', 1.0)}" style="max-width: 90%; height: auto; border: 1px solid #ccc; margin-bottom: 30px;">`;
    }
    if (boxCanvas) {
        html += `<h4 style="margin: 15px 0 10px 0; text-transform: uppercase;">Diagrama de Caja y Bigotes</h4>`;
        html += `<img src="${boxCanvas.toDataURL('image/png', 1.0)}" style="max-width: 90%; height: auto; border: 1px solid #ccc; margin-bottom: 30px;">`;
    }
    
    html += `</div>`;
    element.innerHTML = html;

    const opt = {
        margin:       15,
        filename:     `Reporte_Estadistico_${dataset.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.pdf`,
        image:        { type: 'jpeg', quality: 1.0 },
        html2canvas:  { scale: 2, useCORS: true },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    const btn = document.getElementById('exportPdfBtn');
    const originalText = btn.innerText;
    btn.innerText = 'Generando Reporte...';
    btn.disabled = true;

    html2pdf().set(opt).from(element).save().then(() => {
        btn.innerText = originalText;
        btn.disabled = false;
    });
}