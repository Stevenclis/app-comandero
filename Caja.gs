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

// Calcula el estado vivo de una caja específica en una fecha
function _calcularEstadoUnaCaja(fecha, nombreCaja, datosC, datosM) {
  let ventasEfectivo = 0;
  let ventasTransferencia = 0;

  datosC.forEach(f => {
    if (
      f[2] === 'Pagada' &&
      f[1] === fecha &&
      f[8] &&
      f[8].toString().toLowerCase().trim() === nombreCaja.toLowerCase().trim()
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

  datosM.forEach(f => {
    if (f[0] === fecha && f[1] && f[1].toString().toLowerCase().trim() === nombreCaja.toLowerCase().trim()) {
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

// Cierre detallado por caja (incluyendo desglose de Base, Ventas, Caja Fuerte, Banco y Pendientes)
function hacerCierreDetallado(fecha, caja) {
  const sheetC = _getSheetComandas();
  const sheetM = _getSheetMovimientos();
  const datosC = sheetC.getDataRange().getDisplayValues(); datosC.shift();
  const datosM = sheetM.getDataRange().getDisplayValues(); datosM.shift();

  const info = _calcularEstadoUnaCaja(fecha, caja, datosC, datosM);

  let gastosCFHoy = 0;
  let listaGastosCF = [];
  let gastosBancoHoy = 0;
  let listaGastosBanco = [];

  datosM.forEach(f => {
    if (f[0] === fecha && (f[2] || '').toString().trim() === 'Gasto') {
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

    gastosCajaFuerteHoy: gastosCFHoy,
    listaGastosCajaFuerte: listaGastosCF,
    saldoCajaFuerteActual: _leerSaldoCajaFuerte(),

    gastosBancoHoy: gastosBancoHoy,
    listaGastosBanco: listaGastosBanco,
    saldoBancoActual: _leerSaldoBanco(),

    facturasPendientes: _leerFacturasPendientes()
  };
}