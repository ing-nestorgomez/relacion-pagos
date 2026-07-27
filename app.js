let resultadosConsolidados = [];

// Función para limpiar el RIFs (ej: "J-401800440 " -> "J401800440")
function normalizarRif(rif) {
  if (!rif) return "";
  return String(rif).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// Función para formatear fechas a DD/MM/YYYY
function formatearFecha(fechaRaw) {
  if (!fechaRaw) return "23/07/2026";
  if (typeof fechaRaw === 'string' && fechaRaw.includes('-')) {
    const partes = fechaRaw.split('T')[0].split('-');
    if (partes.length === 3) return `${partes[2]}/${partes[1]}/${partes[0]}`;
  }
  return String(fechaRaw);
}

// Función auxiliar para leer archivos Excel
function leerExcel(inputElement) {
  return new Promise((resolve, reject) => {
    const file = inputElement?.files?.[0];
    if (!file) {
      resolve(null);
      return;
    }
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

// 1. Botón Procesar y Cruzar Datos
document.getElementById('btnProcesar')?.addEventListener('click', async () => {
  console.log("🚀 Iniciando procesamiento de 49 pestañas...");

  const inputComprobante = document.getElementById('fileComprobante') || document.querySelectorAll('input[type="file"]')[0];
  const inputBancos = document.getElementById('fileBancos') || document.querySelectorAll('input[type="file"]')[1];

  try {
    const comprobanteWB = await leerExcel(inputComprobante);
    const bancosWB = await leerExcel(inputBancos);

    if (!comprobanteWB || !bancosWB) {
      alert("⚠️ Por favor, selecciona y carga ambos archivos antes de procesar.");
      return;
    }

    // 1. Crear Mapa de Proveedores desde el Directorio (Hoja 1)
    const sheetBancos = bancosWB.Sheets[bancosWB.SheetNames[0]];
    const rowsBancos = XLSX.utils.sheet_to_json(sheetBancos, { header: 1 });
    
    const mapaProveedores = {};
    rowsBancos.slice(1).forEach(row => {
      if (row && row[0]) {
        const rifNorm = normalizarRif(row[0]);
        mapaProveedores[rifNorm] = {
          rifOriginal: String(row[0]).trim(),
          proveedor: row[1] ? String(row[1]).trim() : "",
          banco: row[2] ? String(row[2]).trim() : "",
          cuenta: row[3] ? String(row[3]).trim() : "",
          tipo: row[4] ? String(row[4]).trim() : ""
        };
      }
    });

    resultadosConsolidados = [];

    // 2. Iterar por TODAS las pestañas del Comprobante
    comprobanteWB.SheetNames.forEach(sheetName => {
      const sheet = comprobanteWB.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      let proveedor = "";
      let rif = "";
      let fecha = "";
      let concepto = "";
      let centroCosto = "";
      let ppto = "SEM 29";
      let totalBs = 0;
      let totalUsd = 0;

      rows.forEach(row => {
        if (!row || row.length === 0) return;

        // Convertir cada celda a string limpia para evaluar
        const col1 = String(row[1] || '').trim().toUpperCase();

        if (col1.includes("PROVEEDOR")) proveedor = String(row[3] || row[2] || '').trim();
        if (col1.includes("RIF")) rif = String(row[3] || row[2] || '').trim();

        if (col1.includes("TOTAL GENERAL") && !col1.includes("USD")) {
          totalBs = row[13] || row[10] || row[row.length - 1] || 0;
        }

        if (col1.includes("TOTAL GENERAL EN USD")) {
          totalUsd = row[13] || row[10] || row[row.length - 1] || 0;
        }

        if (col1.includes("CENTRO DE COSTO")) {
          centroCosto = String(row[3] || '').trim();
        }

        // Extraer concepto (ubicar en fila 31 o por etiqueta)
        if (col1.includes("FORMA DE PAGO")) {
          concepto = String(row[6] || row[5] || '').trim();
        }

        // Fecha de elaboración
        if (col1.includes("FECHA:")) {
          fecha = formatearFecha(row[3]);
        }

        // Buscar Presupuesto
        row.forEach(cell => {
          const strCell = String(cell || '').trim();
          if (strCell.startsWith("SEM")) ppto = strCell;
        });
      });

      // Si encontramos RIF o Proveedor en la pestaña, procesamos la fila
      if (rif || proveedor) {
        const rifNorm = normalizarRif(rif);
        const datosBanco = mapaProveedores[rifNorm] || {};

        resultadosConsolidados.push({
          fecha: fecha || "23/07/2026",
          rif: rif || datosBanco.rifOriginal || "N/A",
          proveedor: proveedor || datosBanco.proveedor || "N/A",
          concepto: concepto || "PAGO DE FACTURAS Y SERVICIOS",
          banco: datosBanco.banco || "NO ENCONTRADO",
          numCuenta: datosBanco.cuenta || "NO ENCONTRADO",
          anexoDirect: datosBanco.banco ? "SI" : "NO",
          centroCosto: centroCosto || "ADMINISTRACION",
          montoBs: typeof totalBs === 'number' ? totalBs : parseFloat(totalBs) || 0,
          montoUsd: typeof totalUsd === 'number' ? totalUsd : parseFloat(totalUsd) || 0,
          ppto: ppto
        });
      }
    });

    // 3. Renderizar resultados en la Vista Previa
    const tbody = document.querySelector('tbody');
    if (tbody) {
      tbody.innerHTML = resultadosConsolidados.map(item => `
        <tr class="border-b border-slate-700/50 hover:bg-slate-800/50 transition-colors">
          <td class="py-2.5 px-3 text-slate-300 text-xs">${item.fecha}</td>
          <td class="py-2.5 px-3 font-mono text-sky-400 font-semibold text-xs">${item.rif}</td>
          <td class="py-2.5 px-3 font-medium text-slate-100 text-xs">${item.proveedor}</td>
          <td class="py-2.5 px-3 text-slate-300 text-xs truncate max-w-xs">${item.concepto}</td>
          <td class="py-2.5 px-3 text-emerald-400 font-medium text-xs">${item.banco}</td>
          <td class="py-2.5 px-3 font-mono text-slate-300 text-xs">${item.numCuenta}</td>
          <td class="py-2.5 px-3 font-bold text-slate-100 text-right text-xs">${item.montoBs.toLocaleString('es-VE', {minimumFractionDigits:2})} Bs.</td>
        </tr>
      `).join('');
    }

    // Mostrar botón de exportación
    const btnExportar = document.getElementById('btnExportar');
    if (btnExportar) btnExportar.classList.remove('hidden');

    alert(`✅ ¡Éxito! Se procesaron ${resultadosConsolidados.length} comprobantes/pestañas correctamente.`);

  } catch (error) {
    console.error("❌ Error en procesamiento:", error);
    alert("Ocurrió un error al procesar los archivos. Consulta la consola.");
  }
});

// 2. Botón Descargar Excel Final
document.getElementById('btnExportar')?.addEventListener('click', () => {
  if (!resultadosConsolidados.length) {
    alert("No hay datos cargados para exportar.");
    return;
  }

  const dataFinal = [
    ["FECHA", "C.I./R.I.F.", "NOMBRE Y/O RAZON SOCIAL", "CONCEPTO", "BANCO", "Nro. CUENTA", "ANEXO DIRECT.", "CENTRO DE COSTO", "MONTO Bs.", "MONTO $", "PPTO N°"]
  ];

  resultadosConsolidados.forEach(item => {
    dataFinal.push([
      item.fecha,
      item.rif,
      item.proveedor,
      item.concepto,
      item.banco,
      item.numCuenta,
      item.anexoDirect,
      item.centroCosto,
      item.montoBs,
      item.montoUsd,
      item.ppto
    ]);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(dataFinal);

  ws['!cols'] = [
    { wch: 12 }, { wch: 16 }, { wch: 35 }, { wch: 40 },
    { wch: 18 }, { wch: 26 }, { wch: 14 }, { wch: 18 },
    { wch: 18 }, { wch: 14 }, { wch: 12 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, "Solicitud_de_Pagos");

  const fechaHoy = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `SOLICITUD_DE_PAGOS_PROCESADA_${fechaHoy}.xlsx`);
});
