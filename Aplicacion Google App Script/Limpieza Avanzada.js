function myFunction() {
  function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Herramientas Avanzadas")
    .addItem("Limpieza Avanzada", "limpiarDatosAvanzado")
    .addToUi();
}

function limpiarDatosAvanzado() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const datos = hoja.getDataRange().getValues();
  if (datos.length <= 1) {
    SpreadsheetApp.getUi().alert("La hoja no contiene suficientes datos.");
    return;
  }

  const encabezado = datos[0];
  let filas = datos.slice(1); // esto quitara el encabezado

  const columnasClave = ["Título", "Artista"]; // Campos que deben estar llenos
  const indiceClave = columnasClave.map(col => encabezado.indexOf(col));

  const totalOriginal = filas.length;

  // 1. Eliminar filas totalmente vacías
  let antes = filas.length;
  filas = eliminarFilasVacias(filas);
  let vaciasEliminadas = antes - filas.length;

  // 2. Eliminar filas con campos clave vacíos
  antes = filas.length;
  filas = eliminarFilasConCamposClaveVacios(filas, indiceClave);
  let claveVaciaEliminadas = antes - filas.length;

  // 3. Eliminar duplicados (por columna específica o toda la fila)
  const columnaClaveDuplicado = encabezado.indexOf("Título");
  antes = filas.length;
  filas = eliminarDuplicadosPorColumna(filas, columnaClaveDuplicado);
  let duplicadosEliminados = antes - filas.length;

  // Limpiar hoja y reescribir datos limpios
  hoja.clearContents();
  hoja.getRange(1, 1, 1, encabezado.length).setValues([encabezado]);
  hoja.getRange(2, 1, filas.length, encabezado.length).setValues(filas);

  // Registrar log
  registrarLog({
    hoja: hoja.getName(),
    totalOriginal,
    vaciasEliminadas,
    claveVaciaEliminadas,
    duplicadosEliminados,
    totalFinal: filas.length
  });

  SpreadsheetApp.getUi().alert(
    `Limpieza avanzada completada:\n\n` +
    `Total original: ${totalOriginal}\n` +
    `Filas vacías eliminadas: ${vaciasEliminadas}\n` +
    `Filas con campos clave vacíos: ${claveVaciaEliminadas}\n` +
    `Duplicados eliminados: ${duplicadosEliminados}\n` +
    `Total final: ${filas.length}`
  );
}

// Elimina filas donde todas las celdas están vacías
function eliminarFilasVacias(filas) {
  return filas.filter(fila => fila.some(celda => celda.toString().trim() !== ""));
}

//Elimina filas donde al menos un campo clave está vacío
function eliminarFilasConCamposClaveVacios(filas, indicesClave) {
  return filas.filter(fila =>
    indicesClave.every(i => fila[i] !== undefined && fila[i].toString().trim() !== "")
  );
}

// Elimina duplicados según una columna específica (ej. Título)
function eliminarDuplicadosPorColumna(filas, indice) {
  const set = new Set();
  const unicas = [];

  for (let fila of filas) {
    const clave = fila[indice].toString().toLowerCase().trim();
    if (!set.has(clave)) {
      set.add(clave);
      unicas.push(fila);
    }
  }

  return unicas;
}

//Registra un log en una hoja aparte
function registrarLog(info) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hojaLog = libro.getSheetByName("Log de Limpieza");

  if (!hojaLog) {
    hojaLog = libro.insertSheet("Log de Limpieza");
    hojaLog.appendRow(["Fecha", "Hoja", "Original", "Vacías", "Campos clave vacíos", "Duplicados", "Final"]);
  }

  hojaLog.appendRow([
    new Date(),
    info.hoja,
    info.totalOriginal,
    info.vaciasEliminadas,
    info.claveVaciaEliminadas,
    info.duplicadosEliminados,
    info.totalFinal
  ]);
}

}
