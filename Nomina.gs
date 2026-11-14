const dbHistorialNomina = SpreadsheetApp.openById('1qXTmDOI_t3Pvv-LTVm0yBb7WKnujz1OsCrpMZxMRBbg');

function obtenerSheetTurnos() {
  let hoja = dbHistorialNomina.getSheetByName('Turnos');
  if (!hoja) {
    hoja = dbHistorialNomina.insertSheet('Turnos');
    hoja.appendRow([
      'ID', 'Fecha', 'Empleado', 'Rol', 'Tipo_Dia',
      'Hora_Entrada', 'Hora_Salida', 'Minutos_Totales',
      'Tipo_Turno', 'Minutos_Extra', 'Pago_Base', 'Pago_Extra',
      'Total_Pagar', 'Estado', 'Entrada_MS', 'Fecha_Pago'
    ]);
  }
  return hoja;
}

// === MEMORIA FANTASMA PARA SEGURIDAD ===
function obtenerTokenTerminal() {
  return PropertiesService.getScriptProperties().getProperty('token_terminal') || '';
}

function guardarTokenTerminal(nuevoToken) {
  PropertiesService.getScriptProperties().setProperty('token_terminal', nuevoToken);
}

function validarTokenTerminal(token) {
  const tokenGuardado = obtenerTokenTerminal();
  if (!tokenGuardado || tokenGuardado === '') return false;
  return token && token.toString().trim() === tokenGuardado;
}

function autorizarTerminal(pin) {
  const usuarios = obtenerListaUsuarios();
  const admin = usuarios.find(u => u.pin === (pin || '').toString().trim() && u.rol.toLowerCase().trim() === 'superadmin');
  if (!admin) return { ok: false, error: 'PIN incorrecto o no corresponde a un Superadmin.' };
  
  const nuevoToken = 'ALT-' + new Date().getTime() + '-' + Math.floor(Math.random() * 100000);
  guardarTokenTerminal(nuevoToken);
  return { ok: true, token: nuevoToken, mensaje: 'Dispositivo autorizado como reloj oficial de Caja.' };
}

function revocarTerminales() {
  guardarTokenTerminal('');
  return { ok: true };
}

// === MEMORIA FANTASMA PARA TARIFAS ===
function obtenerTarifasNomina() {
  const props = PropertiesService.getScriptProperties();
  const defaults = {
    mc_medio_sem: 25000, mc_comp_sem: 60000, mc_medio_sab: 30000, mc_comp_sab: 60000, mc_medio_dom: 40000, mc_comp_dom: 80000,
    coc_medio_sem: 30000, coc_comp_sem: 70000, coc_medio_sab: 40000, coc_comp_sab: 80000, coc_medio_dom: 50000, coc_comp_dom: 100000,
    extra_sem: 5000, extra_fin: 7000, extra_coc_dom: 8000, tolerancia_min: 15
  };
  let tarifas = {};
  for (let k in defaults) {
    let val = props.getProperty(k);
    tarifas[k] = val !== null ? Number(val) : defaults[k];
  }
  return tarifas;
}

function guardarTarifasNomina(nuevasTarifas) {
  const props = PropertiesService.getScriptProperties();
  for (let k in nuevasTarifas) {
    props.setProperty(k, nuevasTarifas[k].toString());
  }
  return { ok: true, tarifas: obtenerTarifasNomina() };
}

// === FERIADOS Y CÁLCULOS ===
function obtenerFeriadosColombia(anio) {
  const a = anio % 19, b = Math.floor(anio / 100), c = anio % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mesPascua = Math.floor((h + l - 7 * m + 114) / 31), diaPascua = ((h + l - 7 * m + 114) % 31) + 1;
  const pascua = new Date(anio, mesPascua - 1, diaPascua, 12, 0, 0);

  function sumarDias(fb, d) { const r = new Date(fb.getTime()); r.setDate(r.getDate() + d); return r; }
  function siguienteLunes(a, m, d) {
    const f = new Date(a, m - 1, d, 12, 0, 0);
    const ds = f.getDay();
    if (ds === 1) return f;
    return sumarDias(f, ds === 0 ? 1 : (8 - ds));
  }
  function fmt(fe) { return `${fe.getFullYear()}-${String(fe.getMonth() + 1).padStart(2, '0')}-${String(fe.getDate()).padStart(2, '0')}`; }

  const lista = [
    { fecha: `${anio}-01-01`, nombre: 'Año Nuevo' }, { fecha: `${anio}-05-01`, nombre: 'Día del Trabajo' },
    { fecha: `${anio}-07-20`, nombre: 'Día de la Independencia' }, { fecha: `${anio}-08-07`, nombre: 'Batalla de Boyacá' },
    { fecha: `${anio}-12-08`, nombre: 'Inmaculada Concepción' }, { fecha: `${anio}-12-25`, nombre: 'Navidad' },
    { fecha: fmt(sumarDias(pascua, -3)), nombre: 'Jueves Santo' }, { fecha: fmt(sumarDias(pascua, -2)), nombre: 'Viernes Santo' },
    { fecha: fmt(siguienteLunes(anio, 1, 6)), nombre: 'Reyes Magos' }, { fecha: fmt(siguienteLunes(anio, 3, 19)), nombre: 'San José' },
    { fecha: fmt(siguienteLunes(anio, 6, 29)), nombre: 'San Pedro y San Pablo' }, { fecha: fmt(siguienteLunes(anio, 8, 15)), nombre: 'Asunción de la Virgen' },
    { fecha: fmt(siguienteLunes(anio, 10, 12)), nombre: 'Día de la Raza' }, { fecha: fmt(siguienteLunes(anio, 11, 1)), nombre: 'Todos los Santos' },
    { fecha: fmt(siguienteLunes(anio, 11, 11)), nombre: 'Independencia de Cartagena' }, { fecha: fmt(sumarDias(pascua, 43)), nombre: 'Ascensión del Señor' },
    { fecha: fmt(sumarDias(pascua, 64)), nombre: 'Corpus Christi' }, { fecha: fmt(sumarDias(pascua, 71)), nombre: 'Sagrado Corazón' }
  ];
  return lista.sort((x, y) => x.fecha.localeCompare(y.fecha));
}

function clasificarDiaColombia(fechaStr) {
  const partes = fechaStr.split('-');
  const fechaObj = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]), 12, 0, 0);
  const feriados = obtenerFeriadosColombia(Number(partes[0]));
  const fer = feriados.find(f => f.fecha === fechaStr);

  if (fer) return { codigo: 'DOM_FEST', etiqueta: `Festivo (${fer.nombre})` };
  if (fechaObj.getDay() === 0) return { codigo: 'DOM_FEST', etiqueta: 'Domingo' };
  if (fechaObj.getDay() === 6) return { codigo: 'SABADO', etiqueta: 'Sábado' };
  return { codigo: 'SEMANA', etiqueta: 'Entre Semana' };
}

function calcularPagoTurno(rol, codigoDia, horaEntradaStr, minutosTotales, tarifas) {
  const esCocina = rol.toLowerCase().includes('cocina');
  const prefijo = esCocina ? 'coc' : 'mc';
  let sufijo = 'sem', tarifaExtra = Number(tarifas.extra_sem || 5000);

  if (codigoDia === 'SABADO') { sufijo = 'sab'; } 
  else if (codigoDia === 'DOM_FEST') { 
    sufijo = 'dom'; 
    tarifaExtra = esCocina ? Number(tarifas.extra_coc_dom || 8000) : Number(tarifas.extra_fin || 7000); 
  }

  const tarifaMedio = Number(tarifas[`${prefijo}_medio_${sufijo}`] || 0);
  const tarifaCompleto = Number(tarifas[`${prefijo}_comp_${sufijo}`] || 0);
  const tolerancia = Number(tarifas.tolerancia_min || 15);
  const horaEntrada = Number((horaEntradaStr || '09:00').split(':')[0]);

  let tipoTurno = 'Medio Turno', minutosBase = 360, pagoBase = tarifaMedio;

  if (horaEntrada < 11 && minutosTotales >= 600) {
    tipoTurno = 'Turno Completo'; minutosBase = 720; pagoBase = tarifaCompleto;
  } else if (minutosTotales < (360 - tolerancia)) {
    tipoTurno = 'Fracción Turno'; minutosBase = minutosTotales;
    pagoBase = Math.round(((minutosTotales / 360) * tarifaMedio) / 100) * 100;
  }

  let minutosExtra = 0, pagoExtra = 0, exceso = minutosTotales - minutosBase;
  if (exceso > tolerancia) {
    minutosExtra = exceso;
    pagoExtra = Math.round(((minutosExtra / 60) * tarifaExtra) / 100) * 100;
  }
  return { tipoTurno: tipoTurno, minutosExtra: minutosExtra, pagoBase: pagoBase, pagoExtra: pagoExtra, totalPagar: pagoBase + pagoExtra };
}

function obtenerListaTurnos() {
  const hoja = obtenerSheetTurnos();
  const datos = hoja.getDataRange().getDisplayValues();
  datos.shift();
  return datos.map(f => ({
    id: f[0], fecha: f[1], empleado: f[2], rol: f[3], tipoDia: f[4],
    horaEntrada: f[5], horaSalida: f[6], minutosTotales: Number(f[7] || 0),
    tipoTurno: f[8], minutosExtra: Number(f[9] || 0), pagoBase: Number((f[10] || '0').toString().replace(/\D/g, '')),
    pagoExtra: Number((f[11] || '0').toString().replace(/\D/g, '')), totalPagar: Number((f[12] || '0').toString().replace(/\D/g, '')),
    estado: f[13], entradaMs: Number(f[14] || 0)
  })).filter(t => t.id !== '' && t.estado !== 'Pagado'); // Solo devuelve a la app los que falten por pagar
}

function obtenerDatosTurnos(token) {
  const autorizado = validarTokenTerminal(token);
  const zona = Session.getScriptTimeZone();
  const hoyStr = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");
  return {
    ok: true, autorizado: autorizado, usuarios: obtenerListaUsuarios(),
    turnos: obtenerListaTurnos(), tarifas: obtenerTarifasNomina(), consumos: obtenerListaConsumos(),
    infoHoy: clasificarDiaColombia(hoyStr), feriadosAnio: obtenerFeriadosColombia(Number(hoyStr.split('-')[0])),
    feriadosProximo: obtenerFeriadosColombia(Number(hoyStr.split('-')[0]) + 1)
  };
}

function marcarEntradaTurno(pin, token) {
  if (!validarTokenTerminal(token)) return { ok: false, noAutorizado: true, error: 'Dispositivo no autorizado.' };
  const user = obtenerListaUsuarios().find(u => u.pin === pin.toString().trim());
  if (!user) return { ok: false, error: 'PIN incorrecto.' };

  const turnos = obtenerListaTurnos();
  if (turnos.find(t => t.empleado.toLowerCase() === user.nombre.toLowerCase() && t.estado === 'Abierto')) {
    return { ok: false, error: `${user.nombre} ya tiene un turno abierto.` };
  }

  const hoja = obtenerSheetTurnos();
  const zona = Session.getScriptTimeZone();
  const ahora = new Date();
  const id = String(ahora.getTime()), fecha = Utilities.formatDate(ahora, zona, "yyyy-MM-dd");
  const hora = Utilities.formatDate(ahora, zona, "HH:mm");
  const diaInfo = clasificarDiaColombia(fecha);

  hoja.appendRow([id, fecha, user.nombre, user.rol, diaInfo.etiqueta, hora, '', 0, 'En curso', 0, 0, 0, 0, 'Abierto', ahora.getTime(), '']);
  return { ok: true, mensaje: `Entrada: ${hora} (${diaInfo.etiqueta})`, turnos: obtenerListaTurnos() };
}

function marcarSalidaTurno(pin, token) {
  if (!validarTokenTerminal(token)) return { ok: false, noAutorizado: true, error: 'Dispositivo no autorizado.' };
  const user = obtenerListaUsuarios().find(u => u.pin === pin.toString().trim());
  if (!user) return { ok: false, error: 'PIN incorrecto.' };

  const hoja = obtenerSheetTurnos();
  const datos = hoja.getDataRange().getDisplayValues();
  const zona = Session.getScriptTimeZone();
  const ahora = new Date();
  const horaSalida = Utilities.formatDate(ahora, zona, "HH:mm");

  for (let i = datos.length - 1; i >= 1; i--) {
    if (datos[i][2].toString().toLowerCase() === user.nombre.toLowerCase() && datos[i][13].toString().trim() === 'Abierto') {
      const fila = i + 1;
      const entradaMs = Number(datos[i][14] || 0);
      let minutosTotales = entradaMs > 0 ? Math.max(1, Math.round((ahora.getTime() - entradaMs) / 60000)) : 1;
      
      const calc = calcularPagoTurno(datos[i][3].toString(), clasificarDiaColombia(datos[i][1].toString()).codigo, datos[i][5].toString(), minutosTotales, obtenerTarifasNomina());

      hoja.getRange(fila, 7).setValue(horaSalida); hoja.getRange(fila, 8).setValue(minutosTotales);
      hoja.getRange(fila, 9).setValue(calc.tipoTurno); hoja.getRange(fila, 10).setValue(calc.minutosExtra);
      hoja.getRange(fila, 11).setValue(calc.pagoBase); hoja.getRange(fila, 12).setValue(calc.pagoExtra);
      hoja.getRange(fila, 13).setValue(calc.totalPagar); hoja.getRange(fila, 14).setValue('Pendiente');

      return { ok: true, resumen: { empleado: user.nombre, horaEntrada: datos[i][5].toString(), horaSalida: horaSalida, minutosTotales: minutosTotales, tipoTurno: calc.tipoTurno, minutosExtra: calc.minutosExtra, totalPagar: calc.totalPagar }, turnos: obtenerListaTurnos() };
    }
  }
  return { ok: false, error: `${user.nombre} no tiene turno abierto.` };
}

function editarTurnoAdmin(id, horaEntrada, horaSalida, tipoTurnoManual, totalManual) {
  const hoja = obtenerSheetTurnos();
  const datos = hoja.getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === id.toString()) {
      const pIn = horaEntrada.split(':').map(Number), pOut = horaSalida.split(':').map(Number);
      let mins = (pOut[0] * 60 + pOut[1]) - (pIn[0] * 60 + pIn[1]);
      if (mins < 0) mins += 1440;
      hoja.getRange(i+1, 6).setValue(horaEntrada); hoja.getRange(i+1, 7).setValue(horaSalida);
      hoja.getRange(i+1, 8).setValue(mins); hoja.getRange(i+1, 9).setValue(tipoTurnoManual);
      hoja.getRange(i+1, 11).setValue(Number(totalManual)); hoja.getRange(i+1, 12).setValue(0);
      hoja.getRange(i+1, 13).setValue(Number(totalManual)); hoja.getRange(i+1, 14).setValue('Pendiente');
      break;
    }
  }
  return { ok: true, turnos: obtenerListaTurnos() };
}

function eliminarTurnoAdmin(id) {
  const hoja = obtenerSheetTurnos();
  const datos = hoja.getDataRange().getDisplayValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0].toString() === id.toString()) { hoja.deleteRow(i + 1); break; }
  }
  return { ok: true, turnos: obtenerListaTurnos() };
}

function liquidarNominaEmpleado(empleado, descontarPlanilla) {
  const hoja = obtenerSheetTurnos();
  const datos = hoja.getDataRange().getDisplayValues();
  const zona = Session.getScriptTimeZone();
  const fechaPago = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd HH:mm");

  // No borra el turno, le pone Pagado y la fecha de pago para el historial
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][2].toString().toLowerCase() === empleado.toLowerCase() && datos[i][13].toString().trim() === 'Pendiente') {
      hoja.getRange(i + 1, 14).setValue('Pagado');
      hoja.getRange(i + 1, 16).setValue(fechaPago);
    }
  }

  let consumosActualizados = obtenerListaConsumos();
  if (descontarPlanilla) {
    const resCons = liquidarConsumoEmpleado(empleado);
    if (resCons && resCons.consumos) consumosActualizados = resCons.consumos;
  }
  return { ok: true, turnos: obtenerListaTurnos(), consumos: consumosActualizados };
}