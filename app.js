// app.js

let datosProcesadosGlobal = []; // Almacena los resultados para exportar

// Función auxiliar para leer un archivo Excel como Workbook
function leerExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      resolve(workbook);
    };
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

// Evento Principal: Procesar y Cruzar
document.getElementById('btnProcesar').addEventListener('click', async () => {
  const inputComprobantes = document.getElementById('fileComprobantes').files[0];
  const inputBancos = document.getElementById('fileBancos').files[0];

  if (!inputComprobantes || !inputBancos) {
    alert("Por favor selecciona ambos archivos de Excel para continuar.");
    return;
  }

  try {
    const wbComprobantes = await leerExcel(inputComprobantes);
    const wbBancos = await leerExcel(inputBancos);

    // 1. Cargar Maestro de Bancos en un Map
    const hojaBancos = wbBancos.Sheets[wbBancos.SheetNames[0]];
    const datosBancos = XLSX.utils.sheet_to_json(hojaBancos);
    
    const mapaBancos = new Map();
    datosBancos.forEach(row => {
      const rifKey = String(row[EXCEL_CONFIG.maestroBancos.colRif] || '').replace(/[\s-]/g, '').toUpperCase();
      if (rifKey) {
        mapaBancos.set(rifKey, {
          banco: row[EXCEL_CONFIG.maestroBancos.colBanco] || 'NO ENCONTRADO',
          cuenta: row[EXCEL_CONFIG.maestroBancos.colCuenta] || 'NO ENCONTRADO'
        });
      }
    });

    // 2. Recorrer Pestañas de Comprobantes
    datosProcesadosGlobal = [];

    wbComprobantes.SheetNames.forEach(sheetName => {
      const sheet = wbComprobantes.Sheets[sheetName];

      const proveedor = sheet[EXCEL_CONFIG.comprobante.proveedor]?.v || '';
      const rifRaw = String(sheet[EXCEL_CONFIG.comprobante.rif]?.v || '');
      const rifClean = rifRaw.replace(/[\s-]/g, '').toUpperCase();
      const montoBs = sheet[EXCEL_CONFIG.comprobante.montoBs]?.v || 0;
      const montoUSD = sheet[EXCEL_CONFIG.comprobante.montoUSD]?.v || 0;
      const centroCosto = sheet[EXCEL_CONFIG.comprobante.centroCosto]?.v || '';
      const textoServicio = sheet[EXCEL_CONFIG.comprobante.compraServicio]?.v || '';

      // Extraer Órdenes de Compra
      let ordenesCompra = [];
      let fila = EXCEL_CONFIG.comprobante.tablaFacturas.filaInicio;
      
      while (fila <= 20) {
        const celdaOC = sheet[`${EXCEL_CONFIG.comprobante.tablaFacturas.colOrdenCompra}${fila}`];
        if (celdaOC && celdaOC.v) {
          ordenesCompra.push(celdaOC.v);
        }
        fila++;
      }

      const conceptoFinal = `${ordenesCompra.join(',')} - ${textoServicio}`;
      const infoBanco = mapaBancos.get(rifClean) || { banco: 'NO REGISTRADO', cuenta: 'NO REGISTRADO' };

      datosProcesadosGlobal.push({
        FECHA: new Date().toLocaleDateString('es-VE'),
        'C.I./RIF': rifRaw,
        'NOMBRE O RAZON SOCIAL': proveedor,
        CONCEPTO: conceptoFinal,
        BANCO: infoBanco.banco,
        'N° CUENTA': infoBanco.cuenta,
        'CENTRO DE COSTO': centroCosto,
        'MONTO Bs.': montoBs,
        'MONTO $': montoUSD
      });
    });

    // 3. Renderizar Vista Previa y habilitar botón de descarga
    renderizarTabla(datosProcesadosGlobal);
    document.getElementById('btnExportar').classList.remove('hidden');

  } catch (error) {
    console.error(error);
    alert("Ocurrió un error al procesar los archivos. Revisa la consola.");
  }
});

// Evento: Exportar a Excel
document.getElementById('btnExportar').addEventListener('click', () => {
  if (datosProcesadosGlobal.length === 0) return;

  // Crear una nueva hoja de cálculo a partir de los datos procesados
  const worksheet = XLSX.utils.json_to_sheet(datosProcesadosGlobal);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Relación de Pagos");

  // Generar y descargar el archivo .xlsx
  XLSX.writeFile(workbook, `Relacion_de_Pagos_${new Date().toISOString().slice(0, 10)}.xlsx`);
});

// Renderizar Vista Previa
function renderizarTabla(datos) {
  const tbody = document.getElementById('tablaVistaPrevia');
  tbody.innerHTML = '';

  datos.forEach(row => {
    const tr = document.createElement('tr');
    tr.className = 'border-b border-slate-700/50 hover:bg-slate-800/50';
    tr.innerHTML = `
      <td class="p-2">${row.FECHA}</td>
      <td class="p-2 font-mono">${row['C.I./RIF']}</td>
      <td class="p-2 font-semibold">${row['NOMBRE O RAZON SOCIAL']}</td>
      <td class="p-2 text-slate-400 max-w-xs truncate" title="${row.CONCEPTO}">${row.CONCEPTO}</td>
      <td class="p-2 text-sky-400">${row.BANCO}</td>
      <td class="p-2 font-mono text-xs">${row['N° CUENTA']}</td>
      <td class="p-2 text-right font-semibold text-emerald-400">${Number(row['MONTO Bs.']).toLocaleString('es-VE', {minimumFractionDigits: 2})}</td>
    `;
    tbody.appendChild(tr);
  });
}
