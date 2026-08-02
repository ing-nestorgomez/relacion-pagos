let resultadosConsolidadosV2 = [];

// Formatear fechas a DD/MM/YYYY
function formatearFechaExcel(val) {
  if (!val) return "";
  
  if (val instanceof Date) {
    const dia = String(val.getDate()).padStart(2, '0');
    const mes = String(val.getMonth() + 1).padStart(2, '0');
    const anio = val.getFullYear();
    return `${dia}/${mes}/${anio}`;
  }

  if (typeof val === 'number') {
    const fechaObj = XLSX.SSF.parse_date_code(val);
    if (fechaObj) {
      const dia = String(fechaObj.d).padStart(2, '0');
      const mes = String(fechaObj.m).padStart(2, '0');
      const anio = fechaObj.y;
      return `${dia}/${mes}/${anio}`;
    }
  }

  let str = String(val).trim();
  if (str.includes('-') && str.length >= 10) {
    const partes = str.split('T')[0].split('-');
    if (partes.length === 3 && partes[0].length === 4) {
      return `${partes[2].padStart(2, '0')}/${partes[1].padStart(2, '0')}/${partes[0]}`;
    }
  }

  return str;
}

function limpiarIdentificacion(ident) {
  if (!ident) return "";
  let str = String(ident).toUpperCase().trim();
  const letraMatch = str.match(/^[JVGEPC]/);
  const letra = letraMatch ? letraMatch[0] : "";
  let numeros = str.replace(/[^0-9]/g, '');
  numeros = numeros.replace(/^0+/, '');
  return letra + numeros;
}

function obtenerBaseNum(limpio) {
  if (!limpio) return { exacto: "", baseConLetra: "", baseNumerica: "", soloNumeros: "" };
  const letra = limpio[0] && isNaN(limpio[0]) ? limpio[0] : "";
  const numeros = letra ? limpio.slice(1) : limpio;
  const base = numeros.length > 7 ? numeros.slice(0, -1) : numeros;

  return {
    exacto: limpio,
    baseConLetra: letra + base,
    baseNumerica: base,
    soloNumeros: numeros
  };
}

function limpiarNumeroCuenta(cuenta) {
  if (!cuenta) return "NO ENCONTRADO";
  const limpia = String(cuenta).replace(/[^0-9]/g, '');
  return limpia.length > 0 ? limpia : String(cuenta).trim();
}

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
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

function obtenerValorCelda(sheet, celdaCoord) {
  if (sheet && sheet[celdaCoord]) {
    return sheet[celdaCoord].v !== undefined ? sheet[celdaCoord].v : sheet[celdaCoord].w;
  }
  return null;
}

document.addEventListener('DOMContentLoaded', () => {
  const btnProcesar = document.getElementById('btnProcesar') || document.querySelector('button');
  
  btnProcesar?.addEventListener('click', async () => {
    console.log("🚀 [SISTEMA V2] Procesando comprobantes...");

    const inputComprobante = document.getElementById('fileComprobantes');
    const inputBancos = document.getElementById('fileBancos');

    try {
      const comprobanteWB = await leerExcel(inputComprobante);
      const bancosWB = await leerExcel(inputBancos);

      if (!comprobanteWB || !bancosWB) {
        alert("⚠️ Por favor, selecciona y carga los 2 archivos Excel (Comprobantes y Maestro de Bancos) antes de procesar.");
        return;
      }

      // 1. Cargar Maestro
      const sheetBancos = bancosWB.Sheets[bancosWB.SheetNames[0]];
      const rowsBancos = XLSX.utils.sheet_to_json(sheetBancos, { header: 1 });
      const listaMaestro = [];

      rowsBancos.forEach(row => {
        if (row && row.length > 0) {
          const celdaRif = String(row[0] || row[1] || '').trim();
          if (celdaRif && !celdaRif.toUpperCase().includes('RIF') && !celdaRif.toUpperCase().includes('C.I.')) {
            const limpio = limpiarIdentificacion(celdaRif);
            const variante = obtenerBaseNum(limpio);
            listaMaestro.push({
              rifOriginal: celdaRif,
              limpio: variante.exacto,
              baseConLetra: variante.baseConLetra,
              baseNumerica: variante.baseNumerica,
              soloNumeros: variante.soloNumeros,
              proveedor: row[1] ? String(row[1]).trim() : "",
              banco: row[2] ? String(row[2]).trim() : "NO ENCONTRADO",
              cuenta: limpiarNumeroCuenta(row[3])
            });
          }
        }
      });

      resultadosConsolidadosV2 = [];

      // 2. Procesar Comprobantes
      comprobanteWB.SheetNames.forEach(sheetName => {
        const sheet = comprobanteWB.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let proveedor = "", rif = "", conceptoPartes = [], centroCosto = "", ppto = "", totalBs = 0;

        let valB10 = obtenerValorCelda(sheet, 'B10');
        let valF28 = obtenerValorCelda(sheet, 'F28');
        let valN35 = obtenerValorCelda(sheet, 'N35');

        let fecha = formatearFechaExcel(valB10);
        let montoUsd = (typeof valF28 === 'number') ? valF28 : (parseFloat(valF28) || 0);
        let semana = valN35 !== null ? String(valN35).trim() : "";

        let valN26 = obtenerValorCelda(sheet, 'N26');
        let valN25 = obtenerValorCelda(sheet, 'N25');
        let valI26 = obtenerValorCelda(sheet, 'I26');
        let valI25 = obtenerValorCelda(sheet, 'I25');

        if (typeof valN26 === 'number' && valN26 > 10) totalBs = valN26;
        else if (typeof valN25 === 'number' && valN25 > 10) totalBs = valN25;
        else if (typeof valI26 === 'number' && valI26 > 10) totalBs = valI26;
        else if (typeof valI25 === 'number' && valI25 > 10) totalBs = valI25;

        let colConcepto = -1, enBloqueConcepto = false;

        rows.forEach((row) => {
          if (!row || row.length === 0) return;

          const filaTexto = row.map(c => String(c || '').trim()).join(' | ');

          if (filaTexto.toUpperCase().includes("PROVEEDOR:")) {
            const idx = row.findIndex(c => String(c).toUpperCase().includes("PROVEEDOR:"));
            if (idx !== -1 && row[idx + 1]) proveedor = String(row[idx + 1]).trim();
            else if (idx !== -1 && row[idx + 2]) proveedor = String(row[idx + 2]).trim();
          }

          if (filaTexto.toUpperCase().includes("RIF:")) {
            const idx = row.findIndex(c => String(c).toUpperCase().includes("RIF:"));
            if (idx !== -1 && row[idx + 1]) rif = String(row[idx + 1]).trim();
            else if (idx !== -1 && row[idx + 2]) rif = String(row[idx + 2]).trim();
          }

          if (totalBs === 0 && filaTexto.toUpperCase().includes("TOTAL GENERAL") && !filaTexto.toUpperCase().includes("USD")) {
            for (let colIdx = 13; colIdx >= 5; colIdx--) {
              const val = row[colIdx];
              if (typeof val === 'number' && val > 10) { totalBs = val; break; }
            }
          }

          if (filaTexto.toUpperCase().includes("CENTRO DE COSTO:")) {
            const idx = row.findIndex(c => String(c).toUpperCase().includes("CENTRO DE COSTO:"));
            if (idx !== -1 && row[idx + 1]) centroCosto = String(row[idx + 1]).trim();
          }

          if (filaTexto.toUpperCase().includes("COMPRA / SERVICIO")) {
            colConcepto = row.findIndex(c => String(c).toUpperCase().includes("COMPRA / SERVICIO"));
            enBloqueConcepto = true;
            return;
          }

          if (enBloqueConcepto) {
            if (filaTexto.toUpperCase().includes("ELABORADO POR") || filaTexto.toUpperCase().includes("DATOS DE ANTICIPOS")) {
              enBloqueConcepto = false;
            } else {
              let valCelda = (colConcepto !== -1 && row[colConcepto]) ? String(row[colConcepto]).trim() : "";
              if (valCelda && valCelda.toUpperCase() !== "COMPRA / SERVICIO" && valCelda.toUpperCase() !== "OBSERVACION") {
                conceptoPartes.push(valCelda.replace(/\r?\n|\r/g, ' '));
              }
            }
          }

          row.forEach(cell => {
            const strCell = String(cell || '').trim();
            if (strCell.toUpperCase().startsWith("SEM")) ppto = strCell;
          });
        });

        let conceptoFinal = conceptoPartes.join(" ").trim() || "PAGO DE FACTURAS Y SERVICIOS";
        if (!semana) semana = ppto;

        if (rif || proveedor || totalBs > 0) {
          const limpioComp = limpiarIdentificacion(rif);
          const varComp = obtenerBaseNum(limpioComp);
          
          let datosBanco = null;
          if (varComp.exacto) datosBanco = listaMaestro.find(item => item.limpio === varComp.exacto);
          if (!datosBanco && varComp.baseConLetra) datosBanco = listaMaestro.find(item => item.baseConLetra === varComp.baseConLetra);
          if (!datosBanco && varComp.baseNumerica) datosBanco = listaMaestro.find(item => item.baseNumerica === varComp.baseNumerica);

          datosBanco = datosBanco || {};

          // Calcular Tasa si existen los dos montos
          let tasaCalculada = (montoUsd > 0 && totalBs > 0) ? (totalBs / montoUsd) : "";

          resultadosConsolidadosV2.push({
            fecha: fecha,
            rif: rif || datosBanco.rifOriginal || "",
            proveedor: proveedor || datosBanco.proveedor || sheetName,
            concepto: conceptoFinal,
            banco: datosBanco.banco || "NO ENCONTRADO",
            numCuenta: datosBanco.cuenta || "NO ENCONTRADO",
            centroCosto: centroCosto,
            montoBs: typeof totalBs === 'number' ? totalBs : parseFloat(totalBs) || 0,
            montoUsd: montoUsd,
            tasa: tasaCalculada,
            semana: semana
          });
        }
      });

      // 3. Renderizar Tabla Previa en Pantalla
      const tbody = document.querySelector('tbody');
      if (tbody) {
        tbody.innerHTML = resultadosConsolidadosV2.map(item => `
          <tr class="border-b border-slate-700/50 hover:bg-slate-800/50 transition-colors">
            <td class="py-2.5 px-3 text-slate-300 text-xs">${item.fecha}</td>
            <td class="py-2.5 px-3 font-mono text-sky-400 font-semibold text-xs">${item.rif}</td>
            <td class="py-2.5 px-3 font-medium text-slate-100 text-xs">${item.proveedor}</td>
            <td class="py-2.5 px-3 text-slate-300 text-xs truncate max-w-xs" title="${item.concepto}">${item.concepto}</td>
            <td class="py-2.5 px-3 ${item.banco === 'NO ENCONTRADO' ? 'text-amber-400 font-bold' : 'text-emerald-400 font-medium'} text-xs">${item.banco}</td>
            <td class="py-2.5 px-3 font-mono text-slate-300 text-xs">${item.numCuenta}</td>
            <td class="py-2.5 px-3 font-bold text-slate-100 text-right text-xs">${item.montoBs.toLocaleString('es-VE', {minimumFractionDigits:2, maximumFractionDigits:2})} Bs.</td>
            <td class="py-2.5 px-3 font-bold text-emerald-300 text-right text-xs">$ ${item.montoUsd.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2})}</td>
            <td class="py-2.5 px-3 text-center text-xs font-semibold text-sky-300">${item.semana}</td>
          </tr>
        `).join('');
      }

      // Mostrar ambos botones de exportación
      document.getElementById('btnExportar')?.classList.remove('hidden');
      document.getElementById('btnExportarV2')?.classList.remove('hidden');

      alert(`✅ ¡Procesamiento completado! Se consolidaron ${resultadosConsolidadosV2.length} registros. Puedes exportar en cualquiera de los dos formatos.`);

    } catch (error) {
      console.error("❌ Error en procesamiento:", error);
      alert("Ocurrió un error al procesar los archivos: " + error.message);
    }
  });

  // BOTÓN 1: Exportar Formato Estándar (Original)
  document.getElementById('btnExportar')?.addEventListener('click', () => {
    if (!resultadosConsolidadosV2.length) return alert("No hay datos para exportar.");

    const datosExcel = resultadosConsolidadosV2.map(item => ({
      "FECHA": item.fecha,
      "C.I./R.I.F.": item.rif,
      "NOMBRE Y/O RAZON SOCIAL": item.proveedor,
      "CONCEPTO": item.concepto,
      "BANCO": item.banco,
      "Nro. CUENTA": item.numCuenta,
      "ANEXO DIRECT": item.banco !== 'NO ENCONTRADO' ? 'SI' : 'NO',
      "CENTRO COSTO": item.centroCosto,
      "TOTAL BS.": item.montoBs,
      "TOTAL $": item.montoUsd,
      "SEMANA": item.semana
    }));

    const ws = XLSX.utils.json_to_sheet(datosExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Solicitud_de_Pagos_V2");
    XLSX.writeFile(wb, `SOLICITUD_DE_PAGOS_${new Date().toISOString().split('T')[0]}.xlsx`);
  });

  // BOTÓN 2: Exportar Formato Procuras y Servicios (Nuevo)
  document.getElementById('btnExportarV2')?.addEventListener('click', () => {
    if (!resultadosConsolidadosV2.length) return alert("No hay datos para exportar.");

    // Mapeo estricto a las columnas solicitadas
    const datosProcuras = resultadosConsolidadosV2.map(item => ({
      "DESCRIPCIÓN DEL REQUERIMIENTO": item.concepto,
      "TOTAL C/IVA": item.montoBs > 0 ? item.montoBs : "",
      "CENTRO DE COSTO": item.centroCosto,
      "CATEGORIA": "",       // Celda en blanco si no hay coincidencia
      "SUB-CATEGORIA": "",   // Celda en blanco
      "SUB-DIVISIÓN": "",    // Celda en blanco
      "ESTATUS": "",         // Celda en blanco
      "PROVEEDOR": item.proveedor,
      "BANCO": item.banco,
      "FECHA EJECUCIÓN": item.fecha,
      "TOTAL BS": item.montoBs > 0 ? item.montoBs : "",
      "TASA": item.tasa !== "" ? Number(item.tasa.toFixed(2)) : "",
      "TOTAL  $": item.montoUsd > 0 ? item.montoUsd : ""
    }));

    // Generar una única hoja de Excel
    const ws = XLSX.utils.json_to_sheet(datosProcuras, {
      header: [
        "DESCRIPCIÓN DEL REQUERIMIENTO",
        "TOTAL C/IVA",
        "CENTRO DE COSTO",
        "CATEGORIA",
        "SUB-CATEGORIA",
        "SUB-DIVISIÓN",
        "ESTATUS",
        "PROVEEDOR",
        "BANCO",
        "FECHA EJECUCIÓN",
        "TOTAL BS",
        "TASA",
        "TOTAL  $"
      ]
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "PROCURAS Y SERVICIOS");
    XLSX.writeFile(wb, `PROCURAS_Y_SERVICIOS_${new Date().toISOString().split('T')[0]}.xlsx`);
  });
});
