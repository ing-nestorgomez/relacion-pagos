// app.js
// Variable global para almacenar los datos cruzados que se exportarán a Excel
let datosProcesadosParaExportar = [];

// Función auxiliar para leer un archivo Excel y retornar su libro de trabajo
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
  console.log("🚀 Iniciando procesamiento...");

  const inputComprobante = document.getElementById('fileComprobante') || document.querySelectorAll('input[type="file"]')[0];
  const inputBancos = document.getElementById('fileBancos') || document.querySelectorAll('input[type="file"]')[1];

  try {
    const comprobanteWB = await leerExcel(inputComprobante);
    const bancosWB = await leerExcel(inputBancos);

    if (!comprobanteWB || !bancosWB) {
      alert("⚠️ Por favor, selecciona y carga ambos archivos antes de procesar.");
      return;
    }

    // Obtener hojas
    const sheetComprobante = comprobanteWB.Sheets[comprobanteWB.SheetNames[0]];
    const sheetBancos = bancosWB.Sheets[bancosWB.SheetNames[0]];

    const rowsComprobante = XLSX.utils.sheet_to_json(sheetComprobante, { header: 1 });
    const rowsBancos = XLSX.utils.sheet_to_json(sheetBancos, { header: 1 });

    let proveedor = "";
    let rif = "";
    let totalPagar = 0;
    let fecha = "23/7/2026";
    let concepto = "MATERIAL DE FERRETERIA VARIAS O/C";

    // Extraer datos de la cabecera
    rowsComprobante.forEach((row) => {
      if (!row || row.length === 0) return;
      const col0 = String(row[0] || '').trim().toUpperCase();

      if (col0 === "PROVEEDOR:") proveedor = row[1] || "";
      if (col0 === "RIF:") rif = row[1] || "";
      if (col0 === "TOTAL GENERAL") {
        totalPagar = row[row.length - 1] || row[12] || 0;
      }
    });

    // Buscar coincidencia en Maestro por RIF
    const cuentaEncontrada = rowsBancos.find((r, idx) => idx > 0 && String(r[0]).trim() === String(rif).trim());

    const banco = cuentaEncontrada ? cuentaEncontrada[2] : "NO ENCONTRADO";
    const numCuenta = cuentaEncontrada ? cuentaEncontrada[4] : "NO ENCONTRADO";

    // Formatear monto para visualización
    const montoFormateado = typeof totalPagar === 'number' 
      ? totalPagar.toLocaleString('es-VE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
      : totalPagar;

    // Renderizar en la tabla
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

    // Guardar los datos limpios para la exportación
    datosProcesadosParaExportar = [
      ["FECHA", "RIF", "PROVEEDOR", "CONCEPTO", "BANCO", "N° CUENTA", "MONTO BS."],
      [fecha, rif, proveedor, concepto, banco, numCuenta, totalPagar]
    ];

    // Mostrar el botón de exportar si estaba oculto
    const btnExportar = document.getElementById('btnExportar');
    if (btnExportar) btnExportar.classList.remove('hidden');

    console.log("✨ Proceso completado con éxito.");

  } catch (error) {
    console.error("❌ Error durante el procesamiento:", error);
    alert("Ocurrió un error al procesar los datos. Revisa la consola.");
  }
});

// 2. Botón Descargar Excel Final
document.getElementById('btnExportar')?.addEventListener('click', () => {
  if (!datosProcesadosParaExportar.length) {
    alert("No hay datos para exportar. Por favor procesa los archivos primero.");
    return;
  }

  // Crear libro y hoja de cálculo
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(datosProcesadosParaExportar);

  // Definir anchos de columnas para que el Excel descargado quede ordenado
  ws['!cols'] = [
    { wch: 12 }, // FECHA
    { wch: 16 }, // RIF
    { wch: 35 }, // PROVEEDOR
    { wch: 40 }, // CONCEPTO
    { wch: 22 }, // BANCO
    { wch: 28 }, // N° CUENTA
    { wch: 18 }  // MONTO BS.
  ];

  XLSX.utils.book_append_sheet(wb, ws, "Relacion_de_Pagos");

  // Descargar el archivo Excel
  const fechaHoy = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Relacion_Pagos_Final_${fechaHoy}.xlsx`);
});
