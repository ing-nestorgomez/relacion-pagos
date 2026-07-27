let resultadosConsolidados = [];

// Función para limpiar RIFs y hacerlos comparables (ej: "J-401800440" -> "J401800440")
function normalizarRif(rif) {
  if (!rif) return "";
  return String(rif).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// Función auxiliar para leer un archivo Excel
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
  console.log("🚀 Iniciando procesamiento masivo...");

  const inputComprobante = document.getElementById('fileComprobante') || document.querySelectorAll('input[type="file"]')[0];
  const inputBancos = document.getElementById('fileBancos') || document.querySelectorAll('input[type="file"]')[1];

  try {
    const comprobanteWB = await leerExcel(inputComprobante);
    const bancosWB = await leerExcel(inputBancos);

    if (!comprobanteWB || !bancosWB) {
      alert("⚠️ Por favor, selecciona y carga ambos archivos antes de procesar.");
      return;
    }

    // Indexar el Directorio de Proveedores por RIF Normalizado
    const sheetBancos = bancosWB.Sheets[bancosWB.SheetNames[0]];
    const rowsBancos = XLSX.utils.sheet_to_json(sheetBancos, { header: 1 });
    
    const mapaProveedores = {};
    rowsBancos.slice(1).forEach(row => {
      if (row && row[0]) {
        const rifNorm = normalizarRif(row[0]);
        mapaProveedores[rifNorm] = {
          rifOriginal: row[0],
          proveedor: row[1] || "",
          banco: row[2] || "",
          cuenta: row[3] || "",
          tipo: row[4] || ""
        };
      }
    });

    resultadosConsolidados = [];

    // Recorrer TODAS las hojas del Libro de Comprobantes
    comprobanteWB.SheetNames.forEach(sheetName => {
      const sheet = comprobanteWB.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      let proveedor = "";
      let rif = "";
      let fecha = "";
      let concepto = "";
      let centroCosto = "";
      let ppto = "";
      let totalBs = 0;
      let totalUsd = 0;

      rows.forEach(row => {
        if (!row || row.length === 0) return;

        const col1 = String(row[1] || '').trim().toUpperCase();
        const col0 = String(row[0] || '').trim().toUpperCase();

        if (col1 === "PROVEEDOR:") proveedor = row[3] || row[2] || "";
        if (col1 === "RIF:") rif = row[3] || row[2] || "";
        if (col1 === "FECHA:") fecha = row[3] || "";
        if (col1 === "CENTRO DE COSTO:") centroCosto = row[3] || "";
        if (col1 === "CONCEPTO DEL PAGO:") concepto = row[6] || row[5] || "";

        if (col1 === "TOTAL GENERAL") {
          totalBs = row[row.length - 2] || row[13] || 0;
        }
        if (col1 === "TOTAL GENERAL EN USD $ (B.C.V)") {
          totalUsd = row[row.length - 2] || row[13] || 0;
        }

        // Buscar Presupuesto (PPTO N°)
        row.forEach(cell => {
          if (cell && String(cell).includes("SEM")) {
            ppto = cell;
          }
        });
      });

      if (rif || proveedor) {
        const rifNorm = normalizarRif(rif);
        const datosBanco = mapaProveedores[rifNorm] || {};

        resultadosConsolidados.push({
          fecha: fecha || new Date().toLocaleDateString('es-VE'),
          rif: rif || datosBanco.rifOriginal || "N/A",
          proveedor: proveedor || datosBanco.proveedor || "N/A",
          concepto: concepto || "PAGO DE FACTURAS Y SERVICIOS",
          banco: datosBanco.banco || "NO ENCONTRADO",
          numCuenta: datosBanco.cuenta || "NO ENCONTRADO",
          anexoDirect: datosBanco.banco ? "SI" : "NO",
          centroCosto: centroCosto || "ADMINISTRACION",
          montoBs: typeof totalBs === 'number' ? totalBs : parseFloat(totalBs) || 0,
          montoUsd: typeof totalUsd === 'number' ? totalUsd : parseFloat(totalUsd) || 0,
          ppto: ppto || "SEM 29"
        });
      }
    });

    // Renderizar Resultados en la Tabla Vista Previa
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

    // Mostrar Botón de Exportar
    const btnExportar = document.getElementById('btnExportar');
    if (btnExportar) btnExportar.classList.remove('hidden');

    alert(`✅ ¡Éxito! Se procesaron ${resultadosConsolidados.length} comprobantes/pestañas correctamente.`);

  } catch (error) {
    console.error("❌ Error en procesamiento:", error);
    alert("Ocurrió un error al procesar los archivos. Consulta la consola.");
  }
});

// 2. Botón Descargar Excel Final (Estructura idéntica a Solicitud de Pagos)
document.getElementById('btnExportar')?.addEventListener('click', () => {
  if (!resultadosConsolidados.length) {
    alert("No hay datos cargados para exportar.");
    return;
  }

  // Encabezados exactamente como en SOLICITUD DE PAGOS
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

  // Anchos de columnas
  ws['!cols'] = [
    { wch: 12 }, { wch: 16 }, { wch: 35 }, { wch: 40 },
    { wch: 18 }, { wch: 26 }, { wch: 14 }, { wch: 18 },
    { wch: 18 }, { wch: 14 }, { wch: 12 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, "Solicitud_de_Pagos");

  const fechaHoy = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `SOLICITUD_DE_PAGOS_PROCESADA_${fechaHoy}.xlsx`);
});
