const libro = SpreadsheetApp.openById('1vVwhCzNQIvIjVY9wniOWr4PyaDk0dKmMQlwaOGIBl30');
const hProductos = libro.getSheetByName('Productos');
const hComandas = libro.getSheetByName('Comandas');
const hDetalles = libro.getSheetByName('Detalles');
const hUsuarios = libro.getSheetByName('Usuarios');
const hCajas = libro.getSheetByName('Cajas'); 
const hMovimientos = libro.getSheetByName('Movimientos');

function doPost(e) {
  try {
    let data = JSON.parse(e.postData.contents);
    let action = data.action;
    let output = {};

    if(action === 'obtenerCargaInicial') output = obtenerCargaInicial();
    if(action === 'sincronizarFondo') output = sincronizarFondo();
    if(action === 'abrirComanda') { abrirComandaDesdeApp(data.id, data.encargado, data.referencia); output = {ok: true}; }
    if(action === 'eliminarComanda') { eliminarComanda(data.id); output = {ok: true}; }
    if(action === 'registrarMovimiento') { registrarMovimientoFondo(data.id, data.nombre, data.precio, data.cantidad); output = {ok: true}; }
    if(action === 'pagarComanda') { pagarComanda(data.id, data.total, data.efe, data.trans, data.caja); output = {ok: true}; }
    if(action === 'obtenerDetalleTicket') output = obtenerDetalleTicket(data.id);
    if(action === 'registrarMovCaja') { registrarMovCaja(data.fecha, data.caja, data.tipo, data.concepto, data.monto); output = {ok: true}; }
    if(action === 'hacerCierre') output = hacerCierre(data.fecha, data.caja);
    if(action === 'obtenerEstadoCajas') output = obtenerEstadoCajas(data.fecha);
    if(action === 'registrarBaseCaja') output = registrarBaseCaja(data.fecha, data.caja, data.monto);
    if(action === 'ajustarSaldoCajaFuerte') output = ajustarSaldoCajaFuerte(data.fecha, data.tipoAjuste, data.monto, data.concepto);
    if(action === 'ajustarSaldoBanco') output = ajustarSaldoBanco(data.fecha, data.tipoAjuste, data.monto, data.concepto);
    if(action === 'consignarACajaFuerte') output = consignarACajaFuerte(data.fecha, data.caja, data.monto);
    if(action === 'registrarGastoDetallado') output = registrarGastoDetallado(data.fecha, data.concepto, data.desglose, data.pendienteId);
    if(action === 'guardarFacturaPendiente') output = guardarFacturaPendiente(data.fecha, data.concepto, data.monto);
    if(action === 'eliminarFacturaPendiente') output = eliminarFacturaPendiente(data.id);
    if(action === 'hacerCierreDetallado') output = hacerCierreDetallado(data.fecha, data.caja);

    // ACCIONES EXCLUSIVAS SUPERADMIN
    if(action === 'guardarUsuario') output = guardarUsuario(data.pinOriginal, data.nombre, data.pin, data.rol);
    if(action === 'eliminarUsuario') output = eliminarUsuario(data.pin);
    if(action === 'guardarProducto') output = guardarProducto(data.nombreOriginal, data.categoria, data.nombre, data.precio, data.estado);
    if(action === 'eliminarProducto') output = eliminarProducto(data.nombre);
    if(action === 'guardarMenuDia') output = guardarMenuDia(data.estado, data.precio, data.descripcion);

    // RUTAS HACIA Consumos.gs
    if(action === 'registrarConsumoEmpleado') output = registrarConsumoEmpleado(data.empleado, data.producto, data.cantidad, data.precio, data.registradoPor);
    if(action === 'eliminarConsumoItem') output = eliminarConsumoItem(data.id);
    if(action === 'liquidarConsumoEmpleado') output = liquidarConsumoEmpleado(data.empleado);
    // RUTAS HACIA Nomina.gs
// RUTAS HACIA Nomina.gs
    if(action === 'obtenerDatosTurnos') output = obtenerDatosTurnos(data.token);
    if(action === 'autorizarTerminal') output = autorizarTerminal(data.pin);
    if(action === 'revocarTerminales') output = revocarTerminales();
    if(action === 'marcarEntradaTurno') output = marcarEntradaTurno(data.pin, data.token);
    if(action === 'marcarSalidaTurno') output = marcarSalidaTurno(data.pin, data.token);
    if(action === 'guardarTarifasNomina') output = guardarTarifasNomina(data.tarifas);
    if(action === 'editarTurnoAdmin') output = editarTurnoAdmin(data.id, data.horaEntrada, data.horaSalida, data.tipoTurno, data.total);
    if(action === 'eliminarTurnoAdmin') output = eliminarTurnoAdmin(data.id);
    if(action === 'liquidarNominaEmpleado') output = liquidarNominaEmpleado(data.empleado, data.descontarPlanilla);

    if(action === 'quitarMenuDia') output = quitarMenuDia();

    return ContentService.createTextOutput(JSON.stringify(output)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) { return ContentService.createTextOutput("API del Restaurante Activa ✅").setMimeType(ContentService.MimeType.TEXT); }

function obtenerSheetMenuDia() {
  let hoja = libro.getSheetByName('Menu_Dia');
  if (!hoja) {
    hoja = libro.insertSheet('Menu_Dia');
    hoja.appendRow(['Estado', 'Precio', 'Descripcion', 'Fecha_Act']);
    hoja.appendRow(['Oculto', 0, '', '']);
  }
  return hoja;
}

// === MENÚ DEL DÍA (MEMORIA FANTASMA) ===
function obtenerMenuDia() {
  const menuStr = PropertiesService.getScriptProperties().getProperty('menu_dia_fantasma');
  if (menuStr) { try { return JSON.parse(menuStr); } catch(e) {} }
  return { estado: 'Oculto', precio: 0, descripcion: '', fecha: '' };
}

function guardarMenuDia(estado, precio, descripcion) {
  const zona = Session.getScriptTimeZone();
  const fechaHoy = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");
  const menuObj = { estado: estado, precio: precio, descripcion: descripcion, fecha: fechaHoy };
  PropertiesService.getScriptProperties().setProperty('menu_dia_fantasma', JSON.stringify(menuObj));
  return { ok: true, menuDia: menuObj };
}

function quitarMenuDia() {
  const menuObj = { estado: 'Oculto', precio: 0, descripcion: '', fecha: '' };
  PropertiesService.getScriptProperties().setProperty('menu_dia_fantasma', JSON.stringify(menuObj));
  return { ok: true, menuDia: menuObj };
}

function obtenerListaUsuarios() {
  const dataUsu = hUsuarios.getDataRange().getDisplayValues(); dataUsu.shift();
  return dataUsu.map(f => ({ 
    pin: f[0].toString().trim(), 
    nombre: f[1].toString().trim(), 
    rol: f[2].toString().trim() 
  })).filter(u => u.nombre !== "");
}

function obtenerListaProductos() {
  const dataProd = hProductos.getDataRange().getDisplayValues(); dataProd.shift();
  return dataProd.map(f => ({ 
    categoria: (f[0] || "General").toString().trim(), 
    nombre: f[1].toString().trim(), 
    precio: Number(f[2].toString().replace(/\D/g, '')),
    estado: (f[3] && f[3].toString().trim() !== "") ? f[3].toString().trim() : "Activo"
  })).filter(p => p.nombre !== "");
}

function obtenerCargaInicial() {
  const dataCajas = hCajas.getDataRange().getDisplayValues(); dataCajas.shift();
  const cajas = dataCajas.map(f => f[0]).filter(c => c !== "");
  return { 
    usuarios: obtenerListaUsuarios(), 
    productos: obtenerListaProductos(), 
    cajas: cajas, 
    sync: sincronizarFondo() 
  };
}

function sincronizarFondo() {
  const datosC = hComandas.getDataRange().getDisplayValues();
  const datosD = hDetalles.getDataRange().getDisplayValues();
  datosC.shift(); datosD.shift();
  let mesasAbiertas = [], mesasCerradas = [], tempDetalles = {}; 
  const fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  for(let i = 0; i < datosC.length; i++) {
    let id = datosC[i][0], fechaComanda = datosC[i][1];
    if(datosC[i][2] === 'Abierta') {
      mesasAbiertas.push({ id: id, fecha: fechaComanda, encargado: datosC[i][3], referencia: datosC[i][4] });
      tempDetalles[id] = {};
    } else if (datosC[i][2] === 'Pagada' && fechaComanda === fechaHoy) {
      mesasCerradas.push({ id: id, fecha: fechaComanda, encargado: datosC[i][3], referencia: datosC[i][4], total: datosC[i][5], caja: datosC[i][8] || "N/A" });
    }
  }

  for(let i = 0; i < datosD.length; i++) {
    let idCom = datosD[i][0];
    if(tempDetalles[idCom] !== undefined) {
      let nombre = datosD[i][1], cant = Number(datosD[i][2]), precio = Number(datosD[i][3]), sub = Number(datosD[i][4]);
      if(!tempDetalles[idCom][nombre]) tempDetalles[idCom][nombre] = { nombre: nombre, cant: 0, precio: precio, subtotal: 0 };
      tempDetalles[idCom][nombre].cant += cant;
      tempDetalles[idCom][nombre].subtotal += sub;
    }
  }

  let detallesFinal = {};
  for(let id in tempDetalles) {
    detallesFinal[id] = [];
    for(let nom in tempDetalles[id]) {
      if(tempDetalles[id][nom].cant > 0) detallesFinal[id].push(tempDetalles[id][nom]);
    }
  }
  return { 
    mesas: mesasAbiertas.reverse(), 
    detalles: detallesFinal, 
    cerradas: mesasCerradas.reverse(),
    menuDia: obtenerMenuDia(),
    consumos: obtenerListaConsumos()
  };
}

function abrirComandaDesdeApp(id, encargado, referencia) {
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  hComandas.appendRow([id, fecha, 'Abierta', encargado, referencia, 0, 0, 0, '']);
}

function eliminarComanda(idComanda) {
  const datos = hComandas.getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === idComanda.toString()) { hComandas.deleteRow(i + 1); return true; }
  }
}

function registrarMovimientoFondo(idComanda, nombre, precio, cantidad) {
  const subtotal = precio * cantidad; 
  hDetalles.appendRow([idComanda, nombre, cantidad, precio, subtotal]);
}

function pagarComanda(idComanda, total, efectivo, transferencia, caja) {
  const datos = hComandas.getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === idComanda.toString()) {
      const fila = i + 1;
      hComandas.getRange(fila, 3).setValue('Pagada'); hComandas.getRange(fila, 6).setValue(total);
      hComandas.getRange(fila, 7).setValue(efectivo); hComandas.getRange(fila, 8).setValue(transferencia);
      hComandas.getRange(fila, 9).setValue(caja);
      const montoTrans = Number(String(transferencia || '0').replace(/\D/g, '')) || 0;
      if (montoTrans > 0 && typeof _sumarTransferenciaBanco === 'function') {
        _sumarTransferenciaBanco(montoTrans);
      }
      return true;
    }
  }
}

function obtenerDetalleTicket(idComanda) {
  const datosC = hComandas.getDataRange().getDisplayValues();
  const datosD = hDetalles.getDataRange().getDisplayValues();
  let info = {};
  for(let i=1; i<datosC.length; i++) {
    if(datosC[i][0].toString() === idComanda.toString()) {
      info = { ref: datosC[i][4], fecha: datosC[i][1], encargado: datosC[i][3], total: datosC[i][5], efe: datosC[i][6], tra: datosC[i][7], caja: datosC[i][8] };
      break;
    }
  }
  let itemsObj = {};
  for(let i=1; i<datosD.length; i++) {
    if(datosD[i][0].toString() === idComanda.toString()) {
      let nom = datosD[i][1], cant = Number(datosD[i][2]), sub = Number(datosD[i][4]);
      if(!itemsObj[nom]) itemsObj[nom] = { cant: 0, subtotal: 0 };
      itemsObj[nom].cant += cant; itemsObj[nom].subtotal += sub;
    }
  }
  let itemsLimpios = [];
  for(let k in itemsObj) { if(itemsObj[k].cant > 0) itemsLimpios.push({ nombre: k, cant: itemsObj[k].cant, subtotal: itemsObj[k].subtotal }); }
  return { info: info, items: itemsLimpios };
}

function registrarMovCaja(fecha, caja, tipo, concepto, monto) {
  hMovimientos.appendRow([fecha, caja, tipo, concepto, monto]);
}

function hacerCierre(fecha, caja) {
  const datosC = hComandas.getDataRange().getDisplayValues(); datosC.shift();
  const datosM = hMovimientos.getDataRange().getDisplayValues(); datosM.shift();
  
  let efec = 0, trans = 0;
  datosC.forEach(f => {
    if (f[2] === 'Pagada' && f[1] === fecha && f[8] && f[8].toString().toLowerCase().trim() === caja.toLowerCase().trim()) {
      efec += Number(f[6].toString().replace(/\D/g, '')); trans += Number(f[7].toString().replace(/\D/g, ''));
    }
  });

  let base = 0, gastos = 0;
  let listaGastos = [];
  
  datosM.forEach(f => {
    if (f[0] === fecha && f[1].toString().toLowerCase().trim() === caja.toLowerCase().trim()) {
      let monto = Number(f[4].toString().replace(/\D/g, ''));
      if(f[2] === 'Base') base += monto;
      if(f[2] === 'Gasto') {
        gastos += monto;
        listaGastos.push({concepto: f[3], monto: monto});
      }
    }
  });

  let efectivoEsperado = base + efec - gastos;

  return { 
    efectivo: efec, 
    transferencia: trans, 
    totalVentas: efec + trans,
    base: base,
    gastos: gastos,
    listaGastos: listaGastos,
    efectivoEsperado: efectivoEsperado
  };
}

// =========================================================================
// FUNCIONES DE GESTIÓN SUPERADMIN
// =========================================================================

function guardarUsuario(pinOriginal, nombre, pin, rol) {
  const datos = hUsuarios.getDataRange().getDisplayValues();
  const pinLimpio = pin.toString().trim();
  const origLimpio = pinOriginal ? pinOriginal.toString().trim() : "";

  for (let i = 1; i < datos.length; i++) {
    let pinFila = datos[i][0].toString().trim();
    if (pinFila === pinLimpio && pinFila !== origLimpio) {
      return { ok: false, error: "Ese PIN ya está asignado a otro trabajador." };
    }
  }

  if (origLimpio !== "") {
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === origLimpio) {
        hUsuarios.getRange(i + 1, 1).setValue("'" + pinLimpio);
        hUsuarios.getRange(i + 1, 2).setValue(nombre.trim());
        hUsuarios.getRange(i + 1, 3).setValue(rol.trim());
        return { ok: true, usuarios: obtenerListaUsuarios() };
      }
    }
  }

  hUsuarios.appendRow(["'" + pinLimpio, nombre.trim(), rol.trim()]);
  return { ok: true, usuarios: obtenerListaUsuarios() };
}

function eliminarUsuario(pin) {
  const datos = hUsuarios.getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString().trim() === pin.toString().trim()) {
      hUsuarios.deleteRow(i + 1);
      break;
    }
  }
  return { ok: true, usuarios: obtenerListaUsuarios() };
}

function guardarProducto(nombreOriginal, categoria, nombre, precio, estado) {
  const datos = hProductos.getDataRange().getDisplayValues();
  const nomLimpio = nombre.toString().trim();
  const origLimpio = nombreOriginal ? nombreOriginal.toString().trim() : "";

  for (let i = 1; i < datos.length; i++) {
    let nomFila = datos[i][1].toString().trim();
    if (nomFila.toLowerCase() === nomLimpio.toLowerCase() && nomFila.toLowerCase() !== origLimpio.toLowerCase()) {
      return { ok: false, error: "Ya existe un producto con ese nombre en la carta." };
    }
  }

  if (origLimpio !== "") {
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][1].toString().trim() === origLimpio) {
        hProductos.getRange(i + 1, 1).setValue(categoria.trim());
        hProductos.getRange(i + 1, 2).setValue(nomLimpio);
        hProductos.getRange(i + 1, 3).setValue(Number(precio));
        hProductos.getRange(i + 1, 4).setValue(estado || "Activo");
        return { ok: true, productos: obtenerListaProductos() };
      }
    }
  }

  hProductos.appendRow([categoria.trim(), nomLimpio, Number(precio), estado || "Activo"]);
  return { ok: true, productos: obtenerListaProductos() };
}

function eliminarProducto(nombre) {
  const datos = hProductos.getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1].toString().trim() === nombre.toString().trim()) {
      hProductos.deleteRow(i + 1);
      break;
    }
  }
  return { ok: true, productos: obtenerListaProductos() };
}