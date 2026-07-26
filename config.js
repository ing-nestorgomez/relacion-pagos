// config.js
// Cuando recibas los archivos reales, solo ajustas las celdas aquí sin tocar app.js
const EXCEL_CONFIG = {
  // Celdas dentro de cada pestaña del Comprobante
  comprobante: {
    proveedor: 'B2',
    rif: 'A4',
    montoBs: 'J12',
    montoUSD: 'J14',
    centroCosto: 'B18',
    compraServicio: 'D22',
    // Configuración de la tabla interna de facturas
    tablaFacturas: {
      filaInicio: 6,
      colOrdenCompra: 'D'
    }
  },
  // Columnas esperadas en el Maestro de Proveedores (Bancos)
  maestroBancos: {
    colRif: 'RIF',
    colBanco: 'BANCO',
    colCuenta: 'NUMERO DE CUENTA'
  }
};
