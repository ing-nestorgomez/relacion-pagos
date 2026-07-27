let resultadosConsolidados = [];

// 1. Limpieza Profunda de RIF / Cédula (Remueve ceros a la izquierda y caracteres extraños)
function limpiarIdentificacion(ident) {
  if (!ident) return "";
  let str = String(ident).toUpperCase().trim();
  
  // Extraer Letra inicial (V, J, E, G, P, C)
  const letraMatch = str.match(/^[JVGEPC]/);
  const letra = letraMatch ? letraMatch[0] : "";

  // Dejar solo los números
  let numeros = str.replace(/[^0-9]/g, '');

  // Eliminar ceros a la izquierda tras la letra (ej: V-05990300 -> V5990300)
  numeros = numeros.replace(/^0+/, '');

  return letra + numeros;
}

// 2. Extraer "Base" de la cédula/RIF (sin el dígito verificador final si existe)
function obtenerBaseNum(limpio) {
  if (!limpio) return { exacto: "", base: "", soloNumeros: "" };
  
  const letra = limpio[0] && isNaN(limpio[0]) ? limpio[0] : "";
  const numeros = letra ? limpio.slice(1) : limpio;
  
  // Si tiene más de 7 dígitos, guardamos la base quitando el último dígito (dígito verificador)
  const base = numeros.length > 7 ? numeros.slice(0, -1) : numeros;

  return {
    exacto: limpio,
    baseConLetra: letra + base,
    baseNumerica: base,
    soloNumeros: numeros
  };
}

// Limpiar números de cuenta
function limpiarNumeroCuenta(cuenta) {
  if (!cuenta) return "NO ENCONTRADO";
  const limpia = String(cuenta).replace(/[^0-9]/g, '');
  return limpia.length > 0 ? limpia : String(cuenta).trim();
}

// Leer Excel
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
    console.log("🚀 Procesando con capa avanzada de coincidencias...");

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

      // 1. Cargar Maestro de Proveedores creando múltiples índices de búsqueda
      const sheetBancos = bancosWB.Sheets[bancosWB.SheetNames[0]];
      const rowsBancos = XLSX.utils.sheet_to_json(sheetBancos, { header: 1 });
      
      const listaMaestro = [];

      rowsBancos.forEach(row => {
        if (row && row.length > 0) {
          const celdaRif = String(row[0] || row[1] || '').trim();
          if (celdaRif && !celdaRif.toUpperCase().includes('RIF') && !celdaRif.toUpperCase().includes('C.I.')) {
            
            const limpio = limpiarIdentificacion(celdaRif);
            const variante = obtenerBaseNum(limpio);
            const proveedorNombre = row[1] ? String(row[1]).trim() : "";

            listaMaestro.push({
              rifOriginal: celdaRif,
              limpio: variante.exacto,
              baseConLetra: variante.baseConLetra,
              baseNumerica: variante.baseNumerica,
              soloNumeros: variante.soloNumeros,
              proveedor: proveedorNombre,
              banco: row[2] ? String(row[2]).trim() : "NO ENCONTRADO",
              cuenta: limpiarNumeroCuenta(row[3])
            });
          }
        }
      });

      resultadosConsolidados = [];

      // 2. Iterar sobre las pestañas de Comprobantes
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

        rows.forEach((row) => {
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
            if (filaTexto.toUpperCase().includes("ELABORADO POR") || filaTexto.toUpperCase().includes("DATOS DE ANTICIPOS")) {
              enBloqueConcepto = false;
            } else {
              let valCelda = "";
              if (colConcepto !== -1 && row[colConcepto]) {
                valCelda = String(row[colConcepto]).trim();
              } else {
                const candidato = row.find((c, i) => i >= 4 && String(c).trim().length > 3 && !String(c).toUpperCase().includes("OBSERVACION"));
                if (candidato) valCelda = String(candidato).trim();
              }

              if (valCelda && valCelda.toUpperCase() !== "COMPRA / SERVICIO" && valCelda.toUpperCase() !== "OBSERVACION") {
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

        let conceptoFinal = conceptoPartes.join(" ").trim();
        if (!conceptoFinal) conceptoFinal = "PAGO DE FACTURAS Y SERVICIOS";

        // Realizar Búsqueda Multinivel en el Maestro
        if (rif || proveedor || totalBs > 0) {
          const limpioComp = limpiarIdentificacion(rif);
          const varComp = obtenerBaseNum(limpioComp);
          
          let datosBanco = null;

          // Capa 1: Coincidencia Exacta Limpia (ej: V5990300 === V5990300)
          if (varComp.exacto) {
            datosBanco = listaMaestro.find(item => item.limpio === varComp.exacto);
          }

          // Capa 2: Coincidencia por Base (Ignorando dígito verificador extra, ej: V59903005 con V5990300)
          if (!datosBanco && varComp.baseConLetra) {
            datosBanco = listaMaestro.find(item => 
              item.baseConLetra === varComp.baseConLetra ||
              item.limpio === varComp.baseConLetra ||
              item.baseConLetra === varComp.exacto
            );
          }

          // Capa 3: Coincidencia Numérica Pura
          if (!datosBanco && varComp.baseNumerica) {
            datosBanco = listaMaestro.find(item => 
              item.baseNumerica === varComp.baseNumerica ||
              item.soloNumeros === varComp.soloNumeros
            );
          }

          // Capa 4: Coincidencia por Nombre de Proveedor (Respaldo)
          if (!datosBanco && proveedor) {
            const provLimpio = proveedor.toUpperCase().replace(/[^A-Z0-9]/g, '');
            datosBanco = listaMaestro.find(item => {
              const itemLimpio = item.proveedor.toUpperCase().replace(/[^A-Z0-9]/g, '');
              return itemLimpio.length > 4 && (itemLimpio.includes(provLimpio) || provLimpio.includes(itemLimpio));
            });
          }

          datosBanco = datosBanco || {};

          resultadosConsolidados.push({
            fecha: fecha,
            rif: rif || datosBanco.rifOriginal || "N/A",
            proveedor: proveedor || datosBanco.proveedor || sheetName,
            concepto: conceptoFinal,
            banco: datosBanco.banco || "NO ENCONTRADO",
            numCuenta: datosBanco.cuenta || "NO ENCONTRADO",
            anexoDirect: (datosBanco.banco && datosBanco.banco !== "NO ENCONTRADO") ? "SI" : "NO",
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
            <td class="py-2.5 px-3 ${item.banco === 'NO ENCONTRADO' ? 'text-amber-400 font-bold' : 'text-emerald-400 font-medium'} text-xs">${item.banco}</td>
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
