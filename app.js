// app.js
document.getElementById('btnProcesar').addEventListener('click', async () => {
  const inputComprobantes = document.getElementById('fileComprobantes').files[0];
  const inputBancos = document.getElementById('fileBancos').files[0];

  if (!inputComprobantes || !inputBancos) {
    alert("Por favor selecciona ambos archivos de Excel para continuar.");
    return;
  }

  console.log("Iniciando procesamiento...");
  // Aquí agregaremos la lectura con SheetJS
});
