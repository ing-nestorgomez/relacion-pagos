// app.js

let comprobanteWB = null;
let bancosWB = null;

// 1. Lectura del Archivo 1: Comprobante
document.getElementById('fileComprobante')?.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    comprobanteWB = XLSX.read(data, { type: 'array' });
    console.log("✅ Comprobante cargado exitosamente.");
  };
  reader.readAsArrayBuffer(file);
});

// 2. Lectura del Archivo 2: Maestro de Bancos
document.getElementById('fileBancos')?.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    bancosWB = XLSX.read(data, { type: 'array' });
    console.log("✅ Maestro de Bancos cargado exitosamente.");
  };
  reader.readAsArrayBuffer(file);
});

// 3. Procesar y Cruzar Datos
document.getElementById('btnProcesar')?.addEventListener('click', () => {
  console.log("🚀 Iniciando procesamiento...");

  if (!comprobanteWB || !bancosWB) {
    alert("⚠️ Por favor, selecciona y carga ambos archivos antes de procesar.");
    return;
  }

  try {
    // Obtener hojas de trabajo
    const sheetComprobante = comprobanteWB.Sheets[comprobanteWB.SheetNames[0]];
    const sheetBancos = bancosWB.Sheets[bancosWB.SheetNames[0]];

    // Convertir a matriz de arreglos (Rows & Cols)
    const rowsComprobante = XLSX.utils.sheet_to_json(sheetComprobante, { header: 1 });
    const rowsBancos = XLSX.utils.sheet_to_json(sheetBancos, { header: 1 });

    console.log("Filas Comprobante:", rowsComprobante);
    console.log("Filas Bancos:", rowsBancos);

    let proveedor = "";
    let rif = "";
    let totalPagar = 0;
    let fecha = "23/7/2026";
    let concepto = "MATERIAL DE FERRETERIA VARIAS O/C";

    // Extraer Proveedor, RIF y Totales del Comprobante
    rowsComprobante.forEach((row) => {
      if (!row || row.length === 0) return;

      const col0 = String(row[0] || '').trim().toUpperCase();

      if (col0 === "PROVEEDOR:") proveedor = row[1] || "";
      if (col0 === "RIF:") rif = row[1] || "";
      if (col0 === "TOTAL GENERAL") {
        // El monto a pagar está en la última posición con valor
        totalPagar = row[row.length - 1] || row[12] || 0;
      }
    });

    console.log(`Datos extraídos -> Proveedor: ${proveedor}, RIF: ${rif}, Total: ${totalPagar}`);

    // Buscar coincidencia en el Maestro de Bancos por RIF
    const cuentaEncontrada = rowsBancos.find((r, idx) => idx > 0 && String(r[0]).trim() === String(rif).trim());

    const banco = cuentaEncontrada ? cuentaEncontrada[2] : "NO ENCONTRADO";
    const numCuenta = cuentaEncontrada ? cuentaEncontrada[4] : "NO ENCONTRADO";

    // Formatear monto
    const montoFormateado = typeof totalPagar === 'number' 
      ? totalPagar.toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
      : totalPagar;

    // Renderizar en el DOM
    const tbody = document.querySelector('tbody');
    if (tbody) {
      tbody.innerHTML = `
        <tr class="border-b border-slate-700/50 hover:bg-slate-800/50 transition-colors">
          <td class="py-3.5 px-4 text-slate-300">${fecha}</td>
          <td class="py-3.5 px-4 font-mono text-sky-400 font-semibold">${rif || 'N/A'}</td>
          <td class="py-3.5 px-4 font-medium text-slate-100">${proveedor || 'N/A'}</td>
          <td class="py-3.5 px-4 text-slate-300">${concepto}</td>
          <td class="py-3.5 px-4 text-emerald-400 font-medium">${banco}</td>
          <td class="py-3.5 px-4 font-mono text-slate-300">${numCuenta}</td>
          <td class="py-3.5 px-4 font-bold text-slate-100 text-right">${montoFormateado} Bs.</td>
        </tr>
      `;
    }

    // Mostrar botón de exportación
    const btnExportar = document.getElementById('btnExportar');
    if (btnExportar) btnExportar.classList.remove('hidden');

    console.log("✨ Proceso completado con éxito.");

  } catch (error) {
    console.error("❌ Error durante el procesamiento:", error);
    alert("Ocurrió un error al procesar los datos. Revisa la consola para más detalle.");
  }
});
