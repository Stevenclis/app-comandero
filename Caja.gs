// =========================================================
// MÓDULO DE CONTROL DE CAJA, CAJA FUERTE, BANCO Y FACTURAS
// =========================================================

function _getSheetMovimientos() {
  return (typeof hMovimientos !== 'undefined' && hMovimientos)
    ? hMovimientos
    : SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Movimientos');
}

function _getSheetComandas() {
  return (typeof hComandas !== 'undefined' && hComandas)
    ? hComandas
    : SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Comandas');
}

function _leerSaldoCajaFuerte() {
  const val = PropertiesService.getScriptProperties().getProperty('CAJA_FUERTE_SALDO');
  return val !== null ? Number(val) : 0;
}

function _guardarSaldoCajaFuerte(nuevoSaldo) {
  PropertiesService.getScriptProperties().setProperty('CAJA_FUERTE_SALDO', String(Math.max(0, Number(nuevoSaldo) || 0)));
}

function _leerSaldoBanco() {
  const val = PropertiesService.getScriptProperties().getProperty('CUENTA_BANCARIA_SALDO');
  return val !== null ? Number(val) : 0;
}

function _guardarSaldoBanco(nuevoSaldo) {
  PropertiesService.getScriptProperties().setProperty('CUENTA_BANCARIA_SALDO', String(Math.max(0, Number(nuevoSaldo) || 0)));
}

function _sumarTransferenciaBanco(monto) {
  const val = Number(monto) || 0;
  if (val > 0) {
    _guardarSaldoBanco(_leerSaldoBanco() + val);
  }
}

function _leerFacturasPendientes() {
  const raw = PropertiesService.getScriptProperties().getProperty('FACTURAS_PENDIENTES');
  if (!raw) return [];
  try { return JSON.parse(raw); } catch (e) { return []; }
}

function _guardarFacturasPendientes(lista) {
  PropertiesService.getScriptProperties().setProperty('FACTURAS_PENDIENTES', JSON.stringify(lista || []));
}

function _mismaFechaCaja(valDisplay, valRaw, fechaObjetivo) {
  let obj = String(fechaObjetivo || '').trim();
  let anio = 0, mes = 0, dia = 0;

  if (obj.indexOf('-') !== -1) {
    let p = obj.split(' ')[0].split('-');
    if (p.length === 3) { anio = Number(p[0]); mes = Number(p[1]); dia = Number(p[2]); }
  } else if (obj.indexOf('/') !== -1) {
    let p = obj.split(' ')[0].split('/');
    if (p.length === 3) { dia = Number(p[0]); mes = Number(p[1]); anio = Number(p[2]); }
  }

  if (valRaw instanceof Date && !isNaN(valRaw.getTime()) && anio > 0) {
    return valRaw.getFullYear() === anio && (valRaw.getMonth() + 1) === mes && valRaw.getDate() === dia;
  }

  let s = String(valDisplay || '').trim();
  if (!s) return false;
  if (s === obj) return true;

  let soloFecha = s.split(' ')[0].trim();
  if (soloFecha === obj) return true;

  if (anio > 0) {
    if (soloFecha.indexOf('-') !== -1) {
      let p = soloFecha.split('-');
      if (p.length === 3) {
        if (Number(p[0]) === anio && Number(p[1]) === mes && Number(p[2]) === dia) return true;
        if (Number(p[2]) === anio && Number(p[1]) === mes && Number(p[0]) === dia) return true;
      }
    }
    if (soloFecha.indexOf('/') !== -1) {
      let p = soloFecha.split('/');
      if (p.length === 3) {
        if (Number(p[2]) === anio && Number(p[1]) === mes && Number(p[0]) === dia) return true;
        if (Number(p[2]) === anio && Number(p[0]) === mes && Number(p[1]) === dia && Number(p[1]) > 12) return true;
      }
    }
  }
  return false;
}

function _mismaCajaNombre(valCelda, nombreCaja) {
  let c1 = String(valCelda || '').toLowerCase().trim();
  let c2 = String(nombreCaja || '').toLowerCase().trim();
  if (!c1 || !c2) return false;
  if (c1 === c2) return true;
  if (c2 === 'caja piscina' && c1 === 'piscina') return true;
  if (c2 === 'caja restaurante' && c1 === 'restaurante') return true;
  return false;
}

// Calcula el estado vivo de una caja específica en una fecha
function _calcularEstadoUnaCaja(fecha, nombreCaja, datosC, datosM, rawC, rawM) {
  let ventasEfectivo = 0;
  let ventasTransferencia = 0;

  datosC.forEach((f, idx) => {
    let r = rawC ? rawC[idx] : null;
    if (
      String(f[2] || '').trim().toLowerCase() === 'pagada' &&
      _mismaFechaCaja(f[1], r ? r[1] : null, fecha) &&
      _mismaCajaNombre(f[8], nombreCaja)
    ) {
      ventasEfectivo += Number(f[6].toString().replace(/\D/g, '')) || 0;
      ventasTransferencia += Number(f[7].toString().replace(/\D/g, '')) || 0;
    }
  });

  let baseInicial = 0;
  let gastosBase = 0;
  let gastosVentas = 0;
  let consignadoDesdeVentas = 0;
  let consignadoDesdeBase = 0;

  let listaGastosBase = [];
  let listaGastosVentas = [];
  let listaConsignaciones = [];

  datosM.forEach((f, idx) => {
    let r = rawM ? rawM[idx] : null;
    if (_mismaFechaCaja(f[0], r ? r[0] : null, fecha) && _mismaCajaNombre(f[1], nombreCaja)) {
      const tipo = (f[2] || '').toString().trim();
      const concepto = (f[3] || '').toString().trim();
      const monto = Number((f[4] || '0').toString().replace(/\D/g, '')) || 0;
      const origen = (f[5] || '').toString().trim();

      if (tipo === 'Base') {
        baseInicial += monto;
      } else if (tipo === 'Gasto') {
        if (origen === 'Base') {
          gastosBase += monto;
          listaGastosBase.push({ concepto: concepto, monto: monto });
        } else {
          gastosVentas += monto;
          listaGastosVentas.push({ concepto: concepto, monto: monto });
        }
      } else if (tipo === 'Consignacion CF' || tipo === 'Consignacion Banco') {
        listaConsignaciones.push({ concepto: concepto, monto: monto, origen: origen });
        if (origen.indexOf('Ventas:') !== -1) {
          const partes = origen.split('|');
          const mV = Number((partes[0] || '').replace(/\D/g, '')) || 0;
          const mB = Number((partes[1] || '').replace(/\D/g, '')) || 0;
          consignadoDesdeVentas += mV;
          consignadoDesdeBase += mB;
        } else if (origen === 'Base') {
          consignadoDesdeBase += monto;
        } else {
          consignadoDesdeVentas += monto;
        }
      }
    }
  });

  const baseDisponible = Math.max(0, baseInicial - gastosBase - consignadoDesdeBase);
  const ventasDisponible = Math.max(0, ventasEfectivo - gastosVentas - consignadoDesdeVentas);
  const totalFisico = baseDisponible + ventasDisponible;
  const totalConsignadoHoy = consignadoDesdeVentas + consignadoDesdeBase;
  const totalGastosHoy = gastosBase + gastosVentas;

  return {
    caja: nombreCaja,
    baseInicial: baseInicial,
    gastosBase: gastosBase,
    consignadoDesdeBase: consignadoDesdeBase,
    baseDisponible: baseDisponible,
    ventasEfectivo: ventasEfectivo,
    gastosVentas: gastosVentas,
    consignadoDesdeVentas: consignadoDesdeVentas,
    ventasDisponible: ventasDisponible,
    consignadoCF: totalConsignadoHoy,
    gastos: totalGastosHoy,
    efectivoEsperado: totalFisico,
    transferencia: ventasTransferencia,
    alertaMillon: totalFisico >= 1000000,
    listaGastosBase: listaGastosBase,
    listaGastosVentas: listaGastosVentas,
    listaConsignaciones: listaConsignaciones
  };
}

// Estado completo para la pantalla de Control de Caja y alertas
function obtenerEstadoCajas(fecha) {
  const sheetC = _getSheetComandas();
  const sheetM = _getSheetMovimientos();
  const datosC = sheetC.getDataRange().getDisplayValues(); datosC.shift();
  const datosM = sheetM.getDataRange().getDisplayValues(); datosM.shift();

  const rest = _calcularEstadoUnaCaja(fecha, 'Caja Restaurante', datosC, datosM);
  const pisc = _calcularEstadoUnaCaja(fecha, 'Caja Piscina', datosC, datosM);
  const saldoCF = _leerSaldoCajaFuerte();
  const saldoBanco = _leerSaldoBanco();
  const pendientes = _leerFacturasPendientes();

  return {
    ok: true,
    cajaFuerte: saldoCF,
    banco: saldoBanco,
    cajas: {
      'Caja Restaurante': rest,
      'Caja Piscina': pisc
    },
    pendientes: pendientes
  };
}

// Registrar Base Inicial con tope máximo estricto de $500.000
function registrarBaseCaja(fecha, caja, monto) {
  const val = Number(monto) || 0;
  if (val <= 0) return { error: 'Ingresa un monto válido mayor a $0.' };
  if (val > 500000) {
    return { error: 'El monto máximo permitido de Base en caja es de $500.000. El excedente debe guardarse en Caja Fuerte.' };
  }

  const estado = obtenerEstadoCajas(fecha);
  const infoCaja = estado.cajas[caja];
  if (infoCaja && infoCaja.baseInicial > 0) {
    return { error: 'Ya existe una base registrada hoy para ' + caja + '.' };
  }

  const sheetM = _getSheetMovimientos();
  sheetM.appendRow([fecha, caja, 'Base', 'Ingreso de Base Inicial', val, 'Base', 'Pagada']);
  return { ok: true };
}

// Ajustar o ingresar saldo inicial de Caja Fuerte (Superadmin)
function ajustarSaldoCajaFuerte(fecha, tipoAjuste, monto, concepto) {
  const val = Number(monto) || 0;
  if (val < 0) return { error: 'El monto no puede ser negativo.' };

  let saldoActual = _leerSaldoCajaFuerte();
  let nuevoSaldo = saldoActual;
  let desc = concepto || 'Ajuste de Caja Fuerte';

  if (tipoAjuste === 'fijar') {
    nuevoSaldo = val;
    desc = 'Saldo inicial / arqueo Caja Fuerte (' + desc + ')';
  } else if (tipoAjuste === 'sumar') {
    if (val <= 0) return { error: 'Ingresa un monto mayor a $0.' };
    nuevoSaldo = saldoActual + val;
    desc = 'Ingreso a Caja Fuerte: ' + desc;
  } else if (tipoAjuste === 'restar') {
    if (val <= 0) return { error: 'Ingresa un monto mayor a $0.' };
    if (val > saldoActual) return { error: 'No hay suficiente saldo en Caja Fuerte ($' + saldoActual.toLocaleString('es-CO') + ').' };
    nuevoSaldo = saldoActual - val;
    desc = 'Retiro de Caja Fuerte: ' + desc;
  }

  _guardarSaldoCajaFuerte(nuevoSaldo);
  const sheetM = _getSheetMovimientos();
  sheetM.appendRow([fecha, 'Caja Fuerte', 'Ajuste CF', desc, val, 'Caja Fuerte', 'Pagada']);
  return { ok: true, cajaFuerte: nuevoSaldo };
}

// Ajustar o ingresar saldo inicial de Cuenta Bancaria (Superadmin)
function ajustarSaldoBanco(fecha, tipoAjuste, monto, concepto) {
  const val = Number(monto) || 0;
  if (val < 0) return { error: 'El monto no puede ser negativo.' };

  let saldoActual = _leerSaldoBanco();
  let nuevoSaldo = saldoActual;
  let desc = concepto || 'Ajuste de Cuenta Bancaria';

  if (tipoAjuste === 'fijar') {
    nuevoSaldo = val;
    desc = 'Saldo inicial / arqueo Banco (' + desc + ')';
  } else if (tipoAjuste === 'sumar') {
    if (val <= 0) return { error: 'Ingresa un monto mayor a $0.' };
    nuevoSaldo = saldoActual + val;
    desc = 'Ingreso a Cuenta Bancaria: ' + desc;
  } else if (tipoAjuste === 'restar') {
    if (val <= 0) return { error: 'Ingresa un monto mayor a $0.' };
    if (val > saldoActual) return { error: 'No hay suficiente saldo en Cuenta Bancaria ($' + saldoActual.toLocaleString('es-CO') + ').' };
    nuevoSaldo = saldoActual - val;
    desc = 'Retiro / Transferencia personal Banco: ' + desc;
  }

  _guardarSaldoBanco(nuevoSaldo);
  const sheetM = _getSheetMovimientos();
  sheetM.appendRow([fecha, 'Cuenta Bancaria', 'Ajuste Banco', desc, val, 'Cuenta Bancaria', 'Pagada']);
  return { ok: true, banco: nuevoSaldo };
}

// Consignar efectivo desde una caja hacia la Caja Fuerte (Liberar caja)
function consignarACajaFuerte(fecha, caja, monto) {
  const val = Number(monto) || 0;
  if (val <= 0) return { error: 'Ingresa un monto válido para consignar.' };

  const estado = obtenerEstadoCajas(fecha);
  const info = estado.cajas[caja];
  if (!info) return { error: 'Caja no válida.' };

  if (val > info.efectivoEsperado) {
    return {
      error: 'No puedes consignar $' + val.toLocaleString('es-CO') + ' porque en ' + caja + ' solo hay $' + info.efectivoEsperado.toLocaleString('es-CO') + ' en efectivo.'
    };
  }

  let desdeVentas = Math.min(val, info.ventasDisponible);
  let desdeBase = Math.max(0, val - desdeVentas);
  let origenTexto = 'Ventas:' + desdeVentas + '|Base:' + desdeBase;

  const sheetM = _getSheetMovimientos();
  sheetM.appendRow([fecha, caja, 'Consignacion CF', 'Liberación de efectivo a Caja Fuerte', val, origenTexto, 'Pagada']);

  const nuevoSaldoCF = _leerSaldoCajaFuerte() + val;
  _guardarSaldoCajaFuerte(nuevoSaldoCF);

  return { ok: true, cajaFuerte: nuevoSaldoCF };
}

// Registrar Gasto o Factura Pagada (bolsillo único o Mixto, incluyendo Cuenta Bancaria)
function registrarGastoDetallado(fecha, concepto, desglose, pendienteId) {
  if (!concepto || !concepto.trim()) return { error: 'Escribe el concepto o proveedor de la factura/gasto.' };
  if (!Array.isArray(desglose) || desglose.length === 0) return { error: 'Debes indicar de dónde sale el dinero.' };

  const estado = obtenerEstadoCajas(fecha);
  let saldoCF = estado.cajaFuerte;
  let saldoBanco = estado.banco;

  let req = {
    'Caja Restaurante - Base': 0,
    'Caja Restaurante - Ventas': 0,
    'Caja Piscina - Base': 0,
    'Caja Piscina - Ventas': 0,
    'Caja Fuerte': 0,
    'Cuenta Bancaria': 0
  };

  let totalPago = 0;
  desglose.forEach(item => {
    const m = Number(item.monto) || 0;
    if (m > 0 && req.hasOwnProperty(item.bolsillo)) {
      req[item.bolsillo] += m;
      totalPago += m;
    }
  });

  if (totalPago <= 0) return { error: 'El monto total a pagar debe ser mayor a $0.' };

  if (req['Caja Restaurante - Base'] > estado.cajas['Caja Restaurante'].baseDisponible) {
    return { error: 'Saldo insuficiente en Base de Caja Restaurante (Disponible: $' + estado.cajas['Caja Restaurante'].baseDisponible.toLocaleString('es-CO') + ').' };
  }
  if (req['Caja Restaurante - Ventas'] > estado.cajas['Caja Restaurante'].ventasDisponible) {
    return { error: 'Saldo insuficiente en Ventas de Caja Restaurante (Disponible: $' + estado.cajas['Caja Restaurante'].ventasDisponible.toLocaleString('es-CO') + ').' };
  }
  if (req['Caja Piscina - Base'] > estado.cajas['Caja Piscina'].baseDisponible) {
    return { error: 'Saldo insuficiente en Base de Caja Piscina (Disponible: $' + estado.cajas['Caja Piscina'].baseDisponible.toLocaleString('es-CO') + ').' };
  }
  if (req['Caja Piscina - Ventas'] > estado.cajas['Caja Piscina'].ventasDisponible) {
    return { error: 'Saldo insuficiente en Ventas de Caja Piscina (Disponible: $' + estado.cajas['Caja Piscina'].ventasDisponible.toLocaleString('es-CO') + ').' };
  }
  if (req['Caja Fuerte'] > saldoCF) {
    return { error: 'Saldo insuficiente en Caja Fuerte (Disponible: $' + saldoCF.toLocaleString('es-CO') + ').' };
  }
  if (req['Cuenta Bancaria'] > saldoBanco) {
    return { error: 'Saldo insuficiente en Cuenta Bancaria (Disponible: $' + saldoBanco.toLocaleString('es-CO') + ').' };
  }

  const sheetM = _getSheetMovimientos();
  const esMixto = Object.values(req).filter(v => v > 0).length > 1;
  const etiquetaConcepto = esMixto ? concepto.trim() + ' [Pago Mixto]' : concepto.trim();

  if (req['Caja Restaurante - Base'] > 0) {
    sheetM.appendRow([fecha, 'Caja Restaurante', 'Gasto', etiquetaConcepto, req['Caja Restaurante - Base'], 'Base', 'Pagada']);
  }
  if (req['Caja Restaurante - Ventas'] > 0) {
    sheetM.appendRow([fecha, 'Caja Restaurante', 'Gasto', etiquetaConcepto, req['Caja Restaurante - Ventas'], 'Ventas', 'Pagada']);
  }
  if (req['Caja Piscina - Base'] > 0) {
    sheetM.appendRow([fecha, 'Caja Piscina', 'Gasto', etiquetaConcepto, req['Caja Piscina - Base'], 'Base', 'Pagada']);
  }
  if (req['Caja Piscina - Ventas'] > 0) {
    sheetM.appendRow([fecha, 'Caja Piscina', 'Gasto', etiquetaConcepto, req['Caja Piscina - Ventas'], 'Ventas', 'Pagada']);
  }
  if (req['Caja Fuerte'] > 0) {
    sheetM.appendRow([fecha, 'Caja Fuerte', 'Gasto', etiquetaConcepto, req['Caja Fuerte'], 'Caja Fuerte', 'Pagada']);
    _guardarSaldoCajaFuerte(saldoCF - req['Caja Fuerte']);
  }
  if (req['Cuenta Bancaria'] > 0) {
    sheetM.appendRow([fecha, 'Cuenta Bancaria', 'Gasto', etiquetaConcepto, req['Cuenta Bancaria'], 'Cuenta Bancaria', 'Pagada']);
    _guardarSaldoBanco(saldoBanco - req['Cuenta Bancaria']);
  }

  if (pendienteId) {
    const pendientes = _leerFacturasPendientes().filter(p => String(p.id) !== String(pendienteId));
    _guardarFacturasPendientes(pendientes);
  }

  return { ok: true };
}

// Guardar factura pendiente (no descuenta saldo hasta que se pague; sobrevive a las 5:00 AM)
function guardarFacturaPendiente(fecha, concepto, monto) {
  const val = Number(monto) || 0;
  if (!concepto || !concepto.trim()) return { error: 'Escribe el proveedor o concepto de la factura.' };
  if (val <= 0) return { error: 'Ingresa un monto válido mayor a $0.' };

  const lista = _leerFacturasPendientes();
  lista.push({
    id: String(Date.now()),
    fecha: fecha,
    concepto: concepto.trim(),
    monto: val
  });
  _guardarFacturasPendientes(lista);
  return { ok: true, pendientes: lista };
}

function eliminarFacturaPendiente(id) {
  const lista = _leerFacturasPendientes().filter(p => String(p.id) !== String(id));
  _guardarFacturasPendientes(lista);
  return { ok: true, pendientes: lista };
}

// Cierre detallado por caja (incluyendo desglose de Base, Ventas, Productos Unificados, Caja Fuerte, Banco y Pendientes)
function hacerCierreDetallado(fecha, caja) {
  const sheetC = _getSheetComandas();
  const sheetM = _getSheetMovimientos();
  const datosC = sheetC.getDataRange().getDisplayValues(); datosC.shift();
  const rawC = sheetC.getDataRange().getValues(); rawC.shift();
  const datosM = sheetM.getDataRange().getDisplayValues(); datosM.shift();
  const rawM = sheetM.getDataRange().getValues(); rawM.shift();

  const info = _calcularEstadoUnaCaja(fecha, caja, datosC, datosM, rawC, rawM);

  // Recolectar IDs de las comandas pagadas hoy en esta caja para unificar productos vendidos
  let idsComandasCaja = {};
  datosC.forEach((f, idx) => {
    let r = rawC[idx];
    if (
      String(f[2] || '').trim().toLowerCase() === 'pagada' &&
      _mismaFechaCaja(f[1], r ? r[1] : null, fecha) &&
      _mismaCajaNombre(f[8], caja)
    ) {
      let idCom = String(f[0] || '').trim();
      if (idCom) idsComandasCaja[idCom] = true;
    }
  });

  let productosVendidos = [];
  const sheetD = (typeof hDetalles !== 'undefined' && hDetalles) ? hDetalles : libro.getSheetByName('Detalles');
  if (sheetD && Object.keys(idsComandasCaja).length > 0) {
    const datosD = sheetD.getDataRange().getValues();
    datosD.shift();
    let porComandaYProd = {};
    for (let i = 0; i < datosD.length; i++) {
      let idCom = String(datosD[i][0] || '').trim();
      if (idsComandasCaja[idCom]) {
        let nom = String(datosD[i][1] || '').trim();
        let cant = Number(datosD[i][2]) || 0;
        let precio = Number(datosD[i][3]) || 0;
        let sub = Number(datosD[i][4]) || 0;
        if (!nom) continue;
        let key = idCom + '||' + nom;
        if (!porComandaYProd[key]) {
          porComandaYProd[key] = { nombre: nom, cant: 0, precio: precio, subtotal: 0 };
        }
        porComandaYProd[key].cant += cant;
        porComandaYProd[key].subtotal += sub;
        if (precio > 0) porComandaYProd[key].precio = precio;
      }
    }

    let unificado = {};
    for (let key in porComandaYProd) {
      let item = porComandaYProd[key];
      if (item.cant > 0) {
        if (!unificado[item.nombre]) {
          unificado[item.nombre] = { nombre: item.nombre, cant: 0, precio: item.precio, subtotal: 0 };
        }
        unificado[item.nombre].cant += item.cant;
        unificado[item.nombre].subtotal += item.subtotal;
        if (item.precio > 0) unificado[item.nombre].precio = item.precio;
      }
    }

    productosVendidos = Object.values(unificado).sort((a, b) => b.subtotal - a.subtotal);
  }

  let gastosCFHoy = 0;
  let listaGastosCF = [];
  let gastosBancoHoy = 0;
  let listaGastosBanco = [];

  datosM.forEach((f, idx) => {
    let r = rawM[idx];
    if (_mismaFechaCaja(f[0], r ? r[0] : null, fecha) && (f[2] || '').toString().trim() === 'Gasto') {
      const m = Number((f[4] || '0').toString().replace(/\D/g, '')) || 0;
      const lugar = (f[1] || '').toString().trim();
      if (lugar === 'Caja Fuerte') {
        gastosCFHoy += m;
        listaGastosCF.push({ concepto: (f[3] || '').toString(), monto: m });
      } else if (lugar === 'Cuenta Bancaria') {
        gastosBancoHoy += m;
        listaGastosBanco.push({ concepto: (f[3] || '').toString(), monto: m });
      }
    }
  });

  return {
    base: info.baseInicial,
    baseInicial: info.baseInicial,
    gastosBase: info.gastosBase,
    consignadoDesdeBase: info.consignadoDesdeBase,
    baseRestante: info.baseDisponible,

    efectivo: info.ventasEfectivo,
    ventasEfectivo: info.ventasEfectivo,
    gastosVentas: info.gastosVentas,
    consignadoDesdeVentas: info.consignadoDesdeVentas,
    ventasRestante: info.ventasDisponible,

    consignadoCF: info.consignadoCF,
    gastos: info.gastos,
    efectivoEsperado: info.efectivoEsperado,
    transferencia: info.transferencia,

    listaGastosBase: info.listaGastosBase,
    listaGastosVentas: info.listaGastosVentas,
    listaConsignaciones: info.listaConsignaciones,
    productosVendidos: productosVendidos,

    gastosCajaFuerteHoy: gastosCFHoy,
    listaGastosCajaFuerte: listaGastosCF,
    saldoCajaFuerteActual: _leerSaldoCajaFuerte(),

    gastosBancoHoy: gastosBancoHoy,
    listaGastosBanco: listaGastosBanco,
    saldoBancoActual: _leerSaldoBanco(),

    facturasPendientes: _leerFacturasPendientes()
  };
}
