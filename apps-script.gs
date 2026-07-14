// ═══════════════════════════════════════════════════════════
//  SISTEMA RxH — Apps Script completo (con blindaje de seguridad)
//  Pega este código completo en Apps Script, reemplazando todo.
//  Luego: Implementar → Administrar implementaciones → editar
//  la implementación existente → Nueva versión → Guardar.
//  (Editar el deploy existente conserva la misma URL)
//
//  SEGURIDAD: cada petición debe traer el token de Google del
//  usuario logueado (param `token`). verificar_() confirma que sea
//  una cuenta @ttaudit.com válida; si no, no ejecuta nada.
//
//  NUEVO: columna AL "DIAS DETALLE" para el calendario de asistencia
//  (patrón de días trabajados que llena el supervisor).
// ═══════════════════════════════════════════════════════════

var CLIENT_ID = '545979370697-jg1delh4ce63cbcgfnb1dv6ppepg9i3v.apps.googleusercontent.com';
var DOMINIO   = 'ttaudit.com';

// Devuelve el email (minúsculas) si el token es de un usuario @ttaudit.com válido; si no, null.
function verificar_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var key = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)
  );
  var cached = cache.get(key);
  if (cached) return cached;
  try {
    var resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
      { muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) return null;
    var info = JSON.parse(resp.getContentText());
    if (info.aud !== CLIENT_ID) return null;
    if (String(info.email_verified) !== 'true') return null;
    var email = String(info.email || '').toLowerCase();
    if (email.slice(-(DOMINIO.length + 1)) !== '@' + DOMINIO) return null;
    if (info.exp && (parseInt(info.exp, 10) * 1000) < Date.now()) return null;
    cache.put(key, email, 300); // cachear 5 min para bajar latencia
    return email;
  } catch (err) {
    return null;
  }
}


function doPost(e) {
  var data = JSON.parse(e.parameter.data || e.postData.contents);

  // ── SEGURIDAD ──
  if (!verificar_(e.parameter.token || data.token)) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'no-autorizado' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── AGREGAR PERSONA A BBDD (vía POST) ──
  if (data.accion === 'agregarPersona') {
    var bbdd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BBDD');
    bbdd.appendRow([
      data.nombre,       // A: APELLIDOS Y NOMBRE
      data.dni,          // B: DNI
      data.ciudad,       // C: CIUDAD
      data.departamento, // D: DEPARTAMENTO
      data.banco,        // E: BANCO
      data.cuenta,       // F: # CUENTA
      data.cci,          // G: # CCI
      'ACTIVO'           // H: STATUS
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── GUARDAR REGISTRO ──
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTROS');

  sheet.appendRow([
    new Date(),                    // A: FECHA DE REGISTRO
    data.quincena,                 // B: QUINCENA
    data.cliente,                  // C: CLIENTE
    data.proyecto,                 // D: PROYECTO
    data.supervisor,               // E: SUPERVISOR
    data.ejCuentas,                // F: EJECUTIVO DE CUENTAS
    data.nombre,                   // G: APELLIDOS Y NOMBRE
    data.dni,                      // H: DNI
    data.departamento,             // I: DEPARTAMENTO
    data.ciudad,                   // J: CIUDAD
    data.banco,                    // K: BANCO
    data.cuenta,                   // L: # CUENTA
    data.cci,                      // M: # CCI
    data.dias,                     // N: DÍAS TRABAJADOS
    data.pagoDia,                  // O: PAGO POR DÍA
    data.pagoDias,                 // P: PAGO POR DÍAS TRABAJADOS (N*O)
    data.movilidadDia,             // Q: PAGO MOVILIDAD DIARIA
    data.movilidad,                // R: MOVILIDAD (N*Q)
    data.movAdicional,             // S: MOVILIDAD ADICIONAL
    data.movViajera,               // T: MOVILIDAD VIAJERA
    data.megas,                    // U: DATOS INTERNET
    data.otros,                    // V: OTROS CONCEPTOS
    data.adelanto,                 // W: ADELANTO MOV VIAJERA
    parseFloat(data.total) || 0,   // X: TOTAL A PAGAR (número)
    data.rxh,                      // Y: EMITIÓ RXH
    data.fechaRxh,                 // Z: FECHA EMISIÓN
    data.numRxh,                   // AA: # RECIBO
    data.emisor,                   // AB: NOMBRE DEL EMISOR / detalle
    data.fechaAbono,               // AC: FECHA DE ABONO
    'EN REVISIÓN CUENTAS',         // AD: ESTADO
    '',                            // AE: OBSERVACION CUENTAS
    '',                            // AF: FECHA APROBACION CUENTAS
    '',                            // AG: APROBADO POR GERENTE
    '',                            // AH: OBSERVACION CONTABILIDAD
    '',                            // AI: FECHA APROBACION CONTABILIDAD
    'No',                          // AJ: APROBADO CONTABILIDAD
    'No',                          // AK: REPORTE DE GASTOS EMITIDO
    data.diasDetalle || ''         // AL: DIAS DETALLE (patrón de asistencia)
  ]);

  // Forzar columnas L y M como texto para preservar ceros iniciales
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 12).setNumberFormat('@'); // L: # CUENTA
  sheet.getRange(lastRow, 13).setNumberFormat('@'); // M: # CCI

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}


function doGet(e) {
  var accion   = e.parameter.accion || 'dni';
  var callback = e.parameter.callback || '';

  function respond(obj) {
    var json = JSON.stringify(obj);
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── SEGURIDAD ──
  if (!verificar_(e.parameter.token)) {
    return respond({ error: 'no-autorizado' });
  }

  // ── AGREGAR PERSONA A BBDD (vía GET/JSONP) ──
  if (accion === 'agregarPersona') {
    var bbdd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BBDD');
    bbdd.appendRow([
      e.parameter.nombre,       // A: APELLIDOS Y NOMBRE
      e.parameter.dni,          // B: DNI
      e.parameter.ciudad,       // C: CIUDAD
      e.parameter.departamento, // D: DEPARTAMENTO
      e.parameter.banco,        // E: BANCO
      e.parameter.cuenta,       // F: # CUENTA
      e.parameter.cci,          // G: # CCI
      'ACTIVO'                  // H: STATUS
    ]);
    return respond({ status: 'ok' });
  }

  // ── BUSCAR POR DNI ──
  if (accion === 'dni') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BBDD');
    var data  = sheet.getDataRange().getValues();
    var dni   = e.parameter.dni;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).replace(/\.0$/, '') === String(dni)) {
        return respond({
          found:        true,
          nombre:       data[i][0],
          ciudad:       data[i][2],
          departamento: data[i][3],
          ubigeo:       data[i][3],
          banco:        data[i][4],
          cuenta:       String(data[i][5]),
          cci:          String(data[i][6])
        });
      }
    }
    return respond({ found: false });
  }

  // ── OBTENER BBDD COMPLETA (para carga masiva y modal nuevo personal) ──
  if (accion === 'bbdd') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BBDD');
    var data  = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      result.push({
        nombre:       row[0],
        dni:          String(row[1]).replace(/\.0$/, ''),
        ciudad:       row[2],
        departamento: row[3],
        banco:        row[4],
        cuenta:       String(row[5]),
        cci:          String(row[6])
      });
    }
    return respond({ personas: result });
  }

  // ── OBTENER PROYECTOS ──
  if (accion === 'proyectos') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PROYECTOS');
    var rows  = sheet.getDataRange().getValues();
    var hdrs  = rows[0];
    var idxP  = {};
    for (var h = 0; h < hdrs.length; h++) {
      idxP[String(hdrs[h]).trim().toUpperCase()] = h;
    }
    function colP(row, nombre) {
      var i = idxP[String(nombre).trim().toUpperCase()];
      return (i !== undefined && row[i] !== undefined && row[i] !== null && row[i] !== '')
        ? String(row[i]) : '';
    }
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row[1] && !row[2]) continue;
      var fi = colP(row, 'FECHA INICIO') || colP(row, 'FECHA DE INICIO') || '';
      var fechaStr = '';
      if (fi) {
        try { fechaStr = new Date(fi).toLocaleDateString('es-PE'); }
        catch (err) { fechaStr = fi; }
      }
      result.push({
        gerente:     colP(row, 'GERENTE'),
        cliente:     colP(row, 'CLIENTE'),
        proyecto:    colP(row, 'PROYECTO'),
        supervisor:  colP(row, 'SUPERVISOR'),
        ejCuentas:   colP(row, 'EJECUTIVO DE CUENTAS') || colP(row, 'EJ. CUENTAS') || colP(row, 'EJECUTIVO CUENTAS') || colP(row, 'GERENTE') || '',
        fechaInicio: fechaStr,
        status:      colP(row, 'STATUS') || colP(row, 'ESTADO') || 'ACTIVO'
      });
    }
    return respond({ proyectos: result });
  }

  // ── OBTENER REGISTROS ──
  if (accion === 'registros') {
    var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTROS');
    var rows    = sheet.getDataRange().getValues();
    var headers = rows[0];
    var result  = [];

    var idx = {};
    for (var h = 0; h < headers.length; h++) {
      idx[String(headers[h]).trim().toUpperCase()] = h;
    }

    function col(row, nombre, fallback) {
      var i = idx[String(nombre).trim().toUpperCase()];
      return (i !== undefined && row[i] !== undefined && row[i] !== null && row[i] !== '')
        ? row[i]
        : (fallback !== undefined ? fallback : '');
    }

    function colStr(row, nombres) {
      for (var n = 0; n < nombres.length; n++) {
        var val = col(row, nombres[n]);
        if (val !== '') return String(val);
      }
      return '';
    }

    function normalizarRxh(raw) {
      raw = String(raw || '').trim();
      if (!raw) return '';
      if (raw.indexOf('::') !== -1) {
        return raw.split('|').map(function(s) {
          return s.split('::')[0].trim();
        }).filter(Boolean).join('|');
      }
      if (raw.indexOf(',') !== -1) {
        return raw.split(',').map(function(s) { return s.trim(); }).filter(Boolean).join('|');
      }
      return raw;
    }

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row[0] && !row[1]) continue;

      var totalRaw = col(row, 'TOTAL A PAGAR');
      var totalStr = '';
      if (totalRaw !== '' && totalRaw !== null && totalRaw !== undefined) {
        totalStr = String(totalRaw).replace(/[^0-9.\-]/g, '');
      }

      result.push({
        fila:         i + 1,
        fecha:        row[0] ? new Date(row[0]).toLocaleDateString('es-PE') : '',
        quincena:     col(row, 'QUINCENA'),
        cliente:      col(row, 'CLIENTE'),
        proyecto:     col(row, 'PROYECTO'),
        supervisor:   col(row, 'SUPERVISOR'),
        ejCuentas:    colStr(row, ['EJECUTIVO DE CUENTAS', 'EJ. CUENTAS', 'EJECUTIVO CUENTAS', 'GERENTE']),
        nombre:       col(row, 'APELLIDOS Y NOMBRE'),
        dni:          String(col(row, 'DNI')),
        departamento: col(row, 'DEPARTAMENTO'),
        ciudad:       col(row, 'CIUDAD'),
        banco:        col(row, 'BANCO'),
        cuenta:       colStr(row, ['# CUENTA', '#CUENTA', 'N° CUENTA', 'CUENTA']),
        cci:          colStr(row, ['# CCI', '#CCI', 'N° CCI', 'CCI']),
        dias:         col(row, 'DÍAS TRABAJADOS'),
        pagoDia:      col(row, 'PAGO POR DÍA'),
        pagoDias:     col(row, 'PAGO POR DÍAS TRABAJADOS'),
        movilidadDia: col(row, 'PAGO MOVILIDAD DIARIA'),
        movilidad:    col(row, 'MOVILIDAD'),
        movAdicional: col(row, 'MOVILIDAD ADICIONAL'),
        movViajera:   col(row, 'MOVILIDAD VIAJERA'),
        megas:        col(row, 'DATOS INTERNET'),
        otros:        col(row, 'OTROS CONCEPTOS'),
        adelanto:     col(row, 'ADELANTO MOV VIAJERA'),
        total:        totalStr,
        rxh:          col(row, 'EMITIÓ RXH'),
        numRxh:       normalizarRxh(colStr(row, ['# RECIBO', '#RECIBO', 'N° RECIBO', 'RECIBO', 'NUM RECIBO', 'NUMERO RECIBO'])),
        // Detalle completo del RxH (num::fecha::monto::emisor|...) desde columna AB (28)
        rxhDetalle:   String(row[27] || ''),
        emisor:       col(row, 'NOMBRE EMISOR'),
        estado:       col(row, 'ESTADO'),
        obsCuentas:   col(row, 'OBS CUENTAS'),
        aprobadoPor:  col(row, 'APROBADO POR GERENTE'),
        obsCont:      col(row, 'OBS CONTABILIDAD'),
        obsContabilidad: col(row, 'OBS CONTABILIDAD'),
        // Patrón de asistencia (calendario) desde columna AL (38)
        diasDetalle:  String(row[37] || '')
      });
    }
    return respond({ registros: result });
  }

  // ── ACTUALIZAR ESTADO ──
  if (accion === 'actualizar') {
    var sheet       = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTROS');
    var fila        = parseInt(e.parameter.fila);
    var tipo        = e.parameter.tipo;
    var nuevoEstado = e.parameter.estado;
    var obs         = e.parameter.obs   || '';
    var quien       = e.parameter.quien || '';

    if (tipo === 'cuentas') {
      sheet.getRange(fila, 30).setValue(nuevoEstado); // AD: ESTADO
      sheet.getRange(fila, 31).setValue(obs);         // AE: OBS CUENTAS
      sheet.getRange(fila, 32).setValue(new Date());  // AF: FECHA APROBACION
      sheet.getRange(fila, 33).setValue(quien);       // AG: APROBADO POR

    } else if (tipo === 'correccion-monto') {
      var dias      = parseFloat(e.parameter.dias)     || 0;
      var pagoDia   = parseFloat(e.parameter.pagoDia)  || 0;
      var movDia    = parseFloat(e.parameter.movDia)   || 0;
      var movAd     = parseFloat(e.parameter.movAd)    || 0;
      var movVi     = parseFloat(e.parameter.movVi)    || 0;
      var adelanto  = parseFloat(e.parameter.adelanto) || 0;
      var megas     = parseFloat(e.parameter.megas)    || 0;
      var otros     = parseFloat(e.parameter.otros)    || 0;
      var pagoDias  = dias * pagoDia;
      var movilidad = dias * movDia;
      var total     = parseFloat(e.parameter.total)    || 0;
      sheet.getRange(fila, 14).setValue(dias);        // N: DÍAS TRABAJADOS
      sheet.getRange(fila, 15).setValue(pagoDia);     // O: PAGO POR DÍA
      sheet.getRange(fila, 16).setValue(pagoDias);    // P: PAGO POR DÍAS TRABAJADOS
      sheet.getRange(fila, 17).setValue(movDia);      // Q: PAGO MOVILIDAD DIARIA
      sheet.getRange(fila, 18).setValue(movilidad);   // R: MOVILIDAD
      sheet.getRange(fila, 19).setValue(movAd);       // S: MOVILIDAD ADICIONAL
      sheet.getRange(fila, 20).setValue(movVi);       // T: MOVILIDAD VIAJERA
      sheet.getRange(fila, 21).setValue(megas);       // U: DATOS INTERNET
      sheet.getRange(fila, 22).setValue(otros);       // V: OTROS CONCEPTOS
      sheet.getRange(fila, 23).setValue(adelanto);    // W: ADELANTO MOV VIAJERA
      sheet.getRange(fila, 24).setValue(total);       // X: TOTAL A PAGAR
      sheet.getRange(fila, 30).setValue(nuevoEstado); // AD: ESTADO
      sheet.getRange(fila, 31).setValue('');          // AE: limpiar obs cuentas
      // Si el supervisor corrigió también la asistencia, actualizar AL
      if (e.parameter.diasDetalle !== undefined && e.parameter.diasDetalle !== '') {
        sheet.getRange(fila, 38).setValue(e.parameter.diasDetalle); // AL: DIAS DETALLE
      }

    } else if (tipo === 'rxh') {
      sheet.getRange(fila, 25).setValue('Sí');                       // Y: EMITIÓ RXH
      sheet.getRange(fila, 27).setValue(e.parameter.numRxh || '');   // AA: # RECIBO (legible)
      sheet.getRange(fila, 28).setValue(e.parameter.rxh    || '');   // AB: Detalle completo
      sheet.getRange(fila, 30).setValue(nuevoEstado);                // AD: ESTADO
      sheet.getRange(fila, 35).setValue('');                         // AI: limpiar obs contabilidad

    } else if (tipo === 'revertir-contabilidad') {
      sheet.getRange(fila, 30).setValue('OBSERVADO-CONTABILIDAD');   // AD: ESTADO
      sheet.getRange(fila, 34).setValue(obs);                        // AH: OBS CONTABILIDAD
      sheet.getRange(fila, 35).setValue(new Date());                 // AI: FECHA
      sheet.getRange(fila, 36).setValue('No');                       // AJ: APROBADO CONT

    } else {
      // contabilidad (aprobación o rechazo)
      sheet.getRange(fila, 30).setValue(nuevoEstado);                // AD: ESTADO
      sheet.getRange(fila, 34).setValue(obs);                        // AH: OBS CONTABILIDAD
      sheet.getRange(fila, 35).setValue(new Date());                 // AI: FECHA
      sheet.getRange(fila, 36).setValue(nuevoEstado === 'PAGADO' || nuevoEstado === 'APROBADO CONTABILIDAD' ? 'Si' : 'No'); // AJ
    }
    return respond({ status: 'ok' });
  }

  // ── ELIMINAR FILA ──
  if (accion === 'eliminar') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTROS');
    var fila  = parseInt(e.parameter.fila);
    if (fila > 1) {
      sheet.deleteRow(fila);
    }
    return respond({ status: 'ok' });
  }

  return respond({ error: 'accion no reconocida' });
}
