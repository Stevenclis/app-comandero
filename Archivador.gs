function archivarDatosAntiguos() {
  const dbPrincipal = SpreadsheetApp.openById('1vVwhCzNQIvIjVY9wniOWr4PyaDk0dKmMQlwaOGIBl30');
  const dbArchivo = SpreadsheetApp.openById('1qXTmDOI_t3Pvv-LTVm0yBb7WKnujz1OsCrpMZxMRBbg');

  const hComandas = dbPrincipal.getSheetByName('Comandas');
  const hDetalles = dbPrincipal.getSheetByName('Detalles');
  const hMovimientos = dbPrincipal.getSheetByName('Movimientos');

  const arcComandas = dbArchivo.getSheetByName('Comandas');
  const arcDetalles = dbArchivo.getSheetByName('Detalles');
  const arcMovimientos = dbArchivo.getSheetByName('Movimientos');

  // Si falta alguna pestaña en el archivo nuevo, se detiene para proteger tus datos
  if (!arcComandas || !arcDetalles || !arcMovimientos) return;

  const zonaHoraria = Session.getScriptTimeZone();
  const fechaHoy = Utilities.formatDate(new Date(), zonaHoraria, "yyyy-MM-dd");

  const datosC = hComandas.getDataRange().getValues();
  const datosD = hDetalles.getDataRange().getValues();
  const datosM = hMovimientos.getDataRange().getValues();

  let idsArchivados = [];
  let comConservar = [datosC[0]];
  let comMover = [];

  // 1. Separar Comandas Pagadas de días anteriores
  for (let i = 1; i < datosC.length; i++) {
    let id = datosC[i][0].toString();
    let fechaStr = datosC[i][1] instanceof Date 
      ? Utilities.formatDate(datosC[i][1], zonaHoraria, "yyyy-MM-dd") 
      : datosC[i][1].toString().trim();
    let estado = datosC[i][2].toString().trim();

    if (estado === 'Pagada' && fechaStr !== fechaHoy) {
      idsArchivados.push(id);
      comMover.push(datosC[i]);
    } else {
      comConservar.push(datosC[i]);
    }
  }

  // 2. Separar Detalles de esas comandas archivadas
  let detConservar = [datosD[0]];
  let detMover = [];

  for (let i = 1; i < datosD.length; i++) {
    let idCom = datosD[i][0].toString();
    if (idsArchivados.includes(idCom)) {
      detMover.push(datosD[i]);
    } else {
      detConservar.push(datosD[i]);
    }
  }

  // 3. Separar Movimientos de caja (Base/Gastos) de días anteriores
  let movConservar = [datosM[0]];
  let movMover = [];

  for (let i = 1; i < datosM.length; i++) {
    let fechaStr = datosM[i][0] instanceof Date 
      ? Utilities.formatDate(datosM[i][0], zonaHoraria, "yyyy-MM-dd") 
      : datosM[i][0].toString().trim();

    if (fechaStr !== fechaHoy) {
      movMover.push(datosM[i]);
    } else {
      movConservar.push(datosM[i]);
    }
  }

  // 4. Pasar al Archivo Histórico (creando encabezados si está vacío) y limpiar la base principal
  if (comMover.length > 0) {
    if (arcComandas.getLastRow() === 0 && datosC.length > 0) arcComandas.appendRow(datosC[0]);
    arcComandas.getRange(arcComandas.getLastRow() + 1, 1, comMover.length, comMover[0].length).setValues(comMover);
    hComandas.clearContents();
    hComandas.getRange(1, 1, comConservar.length, comConservar[0].length).setValues(comConservar);
  }

  if (detMover.length > 0) {
    if (arcDetalles.getLastRow() === 0 && datosD.length > 0) arcDetalles.appendRow(datosD[0]);
    arcDetalles.getRange(arcDetalles.getLastRow() + 1, 1, detMover.length, detMover[0].length).setValues(detMover);
    hDetalles.clearContents();
    hDetalles.getRange(1, 1, detConservar.length, detConservar[0].length).setValues(detConservar);
  }

  if (movMover.length > 0) {
    if (arcMovimientos.getLastRow() === 0 && datosM.length > 0) arcMovimientos.appendRow(datosM[0]);
    arcMovimientos.getRange(arcMovimientos.getLastRow() + 1, 1, movMover.length, movMover[0].length).setValues(movMover);
    hMovimientos.clearContents();
    hMovimientos.getRange(1, 1, movConservar.length, movConservar[0].length).setValues(movConservar);
  }
}
