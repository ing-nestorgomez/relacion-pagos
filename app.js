let resultadosConsolidados = [];

// Normalizar RIF (ej: "J-401800440 " -> "J401800440")
function normalizarRif(rif) {
  if (!rif) return "";
  return String(rif).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// Formatear fechas a DD/MM/YYYY
function formatearFecha(fechaRaw) {
  if (!fechaRaw) return "23/07/2026";
  if (typeof fechaRaw === 'string' && fechaRaw.includes('-')) {
    const partes = fechaRaw.split('T')[0].split('-');
    if (partes.length === 3) return `${partes[2]}/${partes[1]}/${partes[0]}`;
  }
  return String(fechaRaw);
}

// Leer Excel desde input
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

// Evento Principal
document.addEventListener('DOMContentLoaded', () => {
  const btnProcesar = document.getElementById('btnProcesar') || document.querySelector('button');
  
  btnProcesar?.addEventListener('click', async () => {
    console.log("🚀 Iniciando procesamiento...");

    const inputs = document.querySelectorAll('input[type="file"]');
    const inputComprobante = inputs[0];
    const inputBancos = inputs[1];

    try {
      const comprobanteWB = await leerExcel(inputComprobante);
      const bancosWB = await leerExcel(inputBancos);

      if (!comprobanteWB || !bancosWB) {
        alert("⚠️ Por favor, selecciona y carga ambos archivos Excel antes de procesar.");
        return;
      }

      // 1. Cargar Maestro de Proveedores
      const sheetBancos = bancosWB.Sheets[bancosWB.SheetNames[0]];
      const rowsBancos = XLSX.utils.sheet_to_json(sheetBancos, { header: 1 });
      
      const mapaProveedores = {};
      rowsBancos.forEach(row => {
        if (row && row.length > 0) {
          const celda0 = String(row[0] || '').trim();
          if (celda0 && celda0.toUpperCase() !== 'RIF' && celda0.toUpperCase() !== 'C.I./R.I.F.') {
            const rifNorm = normalizarRif(celda0);
            mapaProveedores[rifNorm] = {
              rifOriginal: celda0,
              proveedor: row[1] ? String(row[1]).trim() : "",
              banco: row[2] ? String(row[2]).trim() : "",
              cuenta: row[3] ? String(row[3]).trim() : "",
              tipo: row[4] ? String(row[4]).trim() : ""
            };
          }
        }
      });

      resultadosConsolidados = [];

      // 2. Iterar sobre las pestañas
      comprobanteWB.SheetNames.forEach(sheetName => {
        const sheet = comprobanteWB.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let proveedor = "";
        let rif = "";
        let fecha = "23/07/2026";
        let conceptoPartes = [];
        let centroCosto = "ADMINISTRACION";
        let ppto = "SEM 29";
        let totalBs = 0;
        let totalUsd = 0;

        let colConcepto = -1;
        let enBloqueConcepto = false;

        rows.forEach((row, rIdx) => {
          if (!row || row.length === 0) return;

          const filaTexto = row.map(c => String(c || '').trim()).join(' | ');

          // Extraer PROVEEDOR
          if (filaTexto.toUpperCase().includes("PROVEEDOR:")) {
            const idx = row.findIndex(c => String(c).toUpperCase().includes("PROVEEDOR:"));
            if (idx !== -1 && row[idx + 1]) proveedor = String(row[idx + 1]).trim();
            else if (idx !== -1 && row[idx + 2]) proveedor = String(row[idx + 2]).trim();
            else if (idx !== -1 && row[idx + 3]) proveedor = String(row[idx + 3]).trim();
          }

          // Extraer RIF
          if (filaTexto.toUpperCase().includes("RIF:")) {
            const idx = row.findIndex(c => String(c).toUpperCase().includes("RIF:"));
            if (idx !== -1 && row[idx + 1]) rif = String(row[idx + 1]).trim();
            else if (idx !== -1 && row[idx + 2]) rif = String(row[idx + 2]).trim();
            else if (idx !== -1 && row[idx + 3]) rif = String(row[idx + 3]).trim();
          }

          // Extraer TOTAL GENERAL en Bs.
          if (filaTexto.toUpperCase().includes("TOTAL GENERAL") && !filaTexto.toUpperCase().includes("USD")) {
            const numeros = row.filter(val => typeof val === 'number' && val > 0);
            if (numeros.length > 0) totalBs = numeros[numeros.length - 1];
          }

          // Extraer TOTAL GENERAL EN USD
          if (filaTexto.toUpperCase().includes("TOTAL GENERAL EN USD") || filaTexto.toUpperCase().includes("USD $")) {
            const numeros = row.filter(val => typeof val === 'number' && val > 0);
            if (numeros.length > 0) totalUsd = numeros[numeros.length - 1];
          }

          // Extraer CENTRO DE COSTO
          if (filaTexto.toUpperCase().includes("CENTRO DE COSTO:")) {
            const idx = row.findIndex(c => String(c).toUpperCase().includes("CENTRO DE COSTO:"));
            if (idx !== -1 && row[idx + 1]) centroCosto = String(row[idx + 1]).trim();
            else if (idx !== -1 && row[idx + 2]) centroCosto = String(row[idx + 2]).trim();
          }

          // Detectar columna de COMPRA / SERVICIO
          if (filaTexto.toUpperCase().includes("COMPRA / SERVICIO")) {
            colConcepto = row.findIndex(c => String(c).toUpperCase().includes("COMPRA / SERVICIO"));
            enBloqueConcepto = true;
            return;
          }

          // Capturar el contenido del cuadro COMPRA / SERVICIO
          if (enBloqueConcepto) {
            // Detener lectura si llegamos a la sección de firmas o presupuestos
            if (filaTexto.toUpperCase().includes("ELABORADO POR") || filaTexto.toUpperCase().includes("DATOS DE ANTICIPOS")) {
              enBloqueConcepto = false;
            } else {
              let valCelda = "";
              if (colConcepto !== -1 && row[colConcepto]) {
                valCelda = String(row[colConcepto]).trim();
              } else {
                // Fallback: buscar cualquier celda con texto descriptivo en la fila
                const candidato = row.find((c, i) => i >= 4 && String(c).trim().length > 3 && !String(c).toUpperCase().includes("OBSERVACION"));
                if (candidato) valCelda = String(candidato).trim();
              }

              if (valCelda && valCelda.toUpperCase() !== "COMPRA / SERVICIO" && valCelda.toUpperCase() !== "OBSERVACION") {
                // Reemplazar saltos de línea por espacios si existen
                conceptoPartes.push(valCelda.replace(/\r?\n|\r/g, ' '));
              }
            }
          }

          // Extraer Presupuesto
          row.forEach(cell => {
            const strCell = String(cell || '').trim();
            if (strCell.toUpperCase().startsWith("SEM")) ppto = strCell;
          });
        });

        // Construir el concepto final unido por espacios
        let conceptoFinal = conceptoPartes.join(" ").trim();
        if (!conceptoFinal) conceptoFinal = "PAGO DE FACTURAS Y SERVICIOS";

        // Registrar pestaña procesada
        if (rif || proveedor || totalBs > 0) {
          const rifNorm = normalizarRif(rif);
          const datosBanco = mapaProveedores[rifNorm] || {};

          resultadosConsolidados.push({
            fecha: fecha,
            rif: rif || datosBanco.rifOriginal || "N/A",
            proveedor: proveedor || datosBanco.proveedor || sheetName,
            concepto: conceptoFinal,
            banco: datosBanco.banco || "NO ENCONTRADO",
            numCuenta: datosBanco.cuenta || "NO ENCONTRADO",
            anexoDirect: datosBanco.banco ? "SI" : "NO",
            centroCosto: centroCosto,
            montoBs: typeof totalBs === 'number' ? totalBs : parseFloat(totalBs) || 0,
            montoUsd: typeof totalUsd === 'number' ? totalUsd : parseFloat(totalUsd) || 0,
            ppto: ppto
          });
        }
      });

      // 3. Renderizar en la Tabla HTML
      const tbody = document.querySelector('tbody');
      if (tbody) {
        tbody.innerHTML = resultadosConsolidados.map(item => `
          <tr class="border-b border-slate-700/50 hover:bg-slate-800/50 transition-colors">
            <td class="py-2.5 px-3 text-slate-300 text-xs">${item.fecha}</td>
            <td class="py-2.5 px-3 font-mono text-sky-400 font-semibold text-xs">${item.rif}</td>
            <td class="py-2.5 px-3 font-medium text-slate-100 text-xs">${item.proveedor}</td>
            <td class="py-2.5 px-3 text-slate-300 text-xs truncate max-w-xs" title="${item.concepto}">${item.concepto}</td>
            <td class="py-2.5 px-3 text-emerald-400 font-medium text-xs">${item.banco}</td>
            <td class="py-2.5 px-3 font-mono text-slate-300 text-xs">${item.numCuenta}</td>
            <td class="py-2.5 px-3 font-bold text-slate-100 text-right text-xs">${item.montoBs.toLocaleString('es-VE', {minimumFractionDigits:2, maximumFractionDigits:2})} Bs.</td>
          </tr>
        `).join('');
      }

      // Habilitar botón de descarga
      const btnExportar = document.getElementById('btnExportar');
      if (btnExportar) btnExportar.classList.remove('hidden');

      alert(`✅ ¡Éxito! Se procesaron ${resultadosConsolidados.length} comprobantes/pestañas correctamente.`);

    } catch (error) {
      console.error("❌ Error en procesamiento:", error);
      alert("Ocurrió un error al procesar los archivos: " + error.message);
    }
  });

  // Botón Exportar Excel Final
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
      { wch: 12 }, { wch: 16 }, { wch: 35 }, { wch: 50 },
      { wch: 18 }, { wch: 26 }, { wch: 14 }, { wch: 18 },
      { wch: 18 }, { wch: 14 }, { wch: 12 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Solicitud_de_Pagos");

    const fechaHoy = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `SOLICITUD_DE_PAGOS_PROCESADA_${fechaHoy}.xlsx`);
  });
});
