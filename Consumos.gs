const dbHistorial = SpreadsheetApp.openById('1qXTmDOI_t3Pvv-LTVm0yBb7WKnujz1OsCrpMZxMRBbg');

function obtenerSheetConsumos() {
  let hoja = dbHistorial.getSheetByName('Consumos');
  if (!hoja) {
    hoja = dbHistorial.insertSheet('Consumos');
    hoja.appendRow(['ID', 'Fecha', 'Hora', 'Empleado', 'Producto', 'Cantidad', 'Precio', 'Subtotal', 'Registrado_Por', 'Estado', 'Fecha_Liquidacion']);
  }
  return hoja;
}

function obtenerListaConsumos() {
  const hoja = obtenerSheetConsumos();
  const datos = hoja.getDataRange().getDisplayValues();
  datos.shift(); // Quitar encabezados
  // Filtramos para enviar a la app SOLO los que dicen "Pendiente"
  return datos.map(f => ({
    id: f[0], fecha: f[1], hora: f[2], empleado: f[3],
    producto: f[4], cantidad: Number(f[5] || 0), precio: Number(f[6] || 0),
    subtotal: Number(f[7] || 0), registradoPor: f[8], estado: f[9]
  })).filter(c => c.id !== '' && c.estado === 'Pendiente');
}

function registrarConsumoEmpleado(empleado, producto, cantidad, precio, registradoPor) {
  const hoja = obtenerSheetConsumos();
  const zona = Session.getScriptTimeZone();
  const ahora = new Date();
  const fecha = Utilities.formatDate(ahora, zona, "yyyy-MM-dd");
  const hora = Utilities.formatDate(ahora, zona, "HH:mm");
  const subtotal = cantidad * precio;
  const id = String(ahora.getTime());

  hoja.appendRow([id, fecha, hora, empleado, producto, cantidad, precio, subtotal, registradoPor, 'Pendiente', '']);
  return { ok: true, consumos: obtenerListaConsumos() };
}

function eliminarConsumoItem(id) {
  const hoja = obtenerSheetConsumos();
  const datos = hoja.getDataRange().getDisplayValues();
  // No lo borra, lo marca como Cancelado para el historial
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === id.toString()) {
      hoja.getRange(i + 1, 10).setValue('Cancelado (Error)');
      break;
    }
  }
  return { ok: true, consumos: obtenerListaConsumos() };
}

function liquidarConsumoEmpleado(empleado) {
  const hoja = obtenerSheetConsumos();
  const datos = hoja.getDataRange().getDisplayValues();
  const zona = Session.getScriptTimeZone();
  const fechaLiq = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd HH:mm");

  // No borra las filas, solo les cambia el estado y les pone la fecha de pago
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][3].toString().toLowerCase() === empleado.toLowerCase() && datos[i][9] === 'Pendiente') {
      hoja.getRange(i + 1, 10).setValue('Liquidado');
      hoja.getRange(i + 1, 11).setValue(fechaLiq);
    }
  }
  return { ok: true, consumos: obtenerListaConsumos() };
}