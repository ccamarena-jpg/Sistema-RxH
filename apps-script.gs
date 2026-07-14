/*************************************************************************
 * Sistema RxH — Backend (Google Apps Script)
 * ----------------------------------------------------------------------
 * BLINDAJE DE SEGURIDAD (Ruta A ligera):
 * Cada petición (doGet y doPost) debe traer el token de identidad de
 * Google (JWT) del usuario logueado. verificar_() lo valida contra Google
 * y confirma que sea una cuenta @ttaudit.com verificada y no expirada.
 * Si no es válido, se responde {error:'no-autorizado'} y NO se ejecuta
 * ninguna acción (ni lectura ni escritura).
 *
 * IMPORTANTE: tras pegar este código, vuelve a DESPLEGAR (Implementar →
 * Gestionar implementaciones → editar → Nueva versión). Mantén "Ejecutar
 * como: Yo" y "Quién tiene acceso: Cualquiera".
 *************************************************************************/

// El client_id de Google del login (debe coincidir con el data-client_id del index.html)
var CLIENT_ID = '545979370697-jg1delh4ce63cbcgfnb1dv6ppepg9i3v.apps.googleusercontent.com';
var DOMINIO   = 'ttaudit.com';

/**
 * Verifica el token de Google. Devuelve el email (minúsculas) si es un
 * usuario @ttaudit.com válido, o null si no lo es.
 */
function verificar_(token) {
  if (!token) return null;

  var cache = CacheService.getScriptCache();
  var key = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)
  );
  var cached = cache.get(key);
  if (cached) return cached;   // ya verificado hace poco → evita latencia

  try {
    var resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
      { muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) return null;

    var info = JSON.parse(resp.getContentText());

    // 1) El token debe haber sido emitido para NUESTRA app
    if (info.aud !== CLIENT_ID) return null;
    // 2) Email verificado
    if (String(info.email_verified) !== 'true') return null;
    // 3) Del dominio de la empresa
    var email = String(info.email || '').toLowerCase();
    if (email.slice(-(DOMINIO.length + 1)) !== '@' + DOMINIO) return null;
    // 4) No expirado (tokeninfo ya rechaza expirados; doble chequeo)
    if (info.exp && (parseInt(info.exp, 10) * 1000) < Date.now()) return null;

    cache.put(key, email, 300);  // cachear 5 min
    return email;
  } catch (err) {
    return null;
  }
}

function noAutorizado_(callback) {
  var json = JSON.stringify({ error: 'no-autorizado' });
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}


function doPost(e) {
  var data = JSON.parse(e.parameter.data || e.postData.contents);

  // ── SEGURIDAD: exigir usuario @ttaudit.com válido ──
  if (!verificar_(e.parameter.token || data.token)) {
    return noAutorizado_(null);
  }

  // ── AGREGAR PERSONA A BBDD ──
  if (data.accion === 'agregarPersona') {
    var bbdd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BBDD");
    bbdd.appendRow([
      data.nombre, data.dni, data.ciudad, data.departamento,
      data.banco, data.cuenta, data.cci, 'ACTIVO'
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({status:"ok"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── GUARDAR REGISTRO ──
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
              .getSheetByName("REGISTROS");

  sheet.appendRow([
    new Date(),            // A: FECHA DE REGISTRO
    data.quincena,         // B: QUINCENA
    data.cliente,          // C: CLIENTE
    data.proyecto,         // D: PROYECTO
    data.supervisor,       // E: SUPERVISOR
    data.ejCuentas,        // F: EJECUTIVO DE CUENTAS
    data.nombre,           // G: APELLIDOS Y NOMBRE
    data.dni,              // H: DNI
    data.departamento,     // I: DEPARTAMENTO
    data.ciudad,           // J: CIUDAD
    data.banco,            // K: BANCO
    data.cuenta,           // L: # CUENTA
    data.cci,              // M: # CCI
    data.dias,             // N: DÍAS TRABAJADOS
    data.pagoDia,          // O: PAGO POR DÍA
    data.pagoDias,         // P: PAGO POR DÍAS TRABAJADOS (N*O)
    data.movilidadDia,     // Q: PAGO MOVILIDAD DIARIA
    data.movilidad,        // R: MOVILIDAD (N*Q)
    data.movAdicional,     // S: MOVILIDAD ADICIONAL
    data.movViajera,       // T: MOVILIDAD VIAJERA
    data.megas,            // U: DATOS INTERNET
    data.otros,            // V: OTROS CONCEPTOS
    data.adelanto,         // W: ADELANTO MOV VIAJERA
    parseFloat(data.total) || 0, // X: TOTAL A PAGAR (número)
    data.rxh,              // Y: EMITIÓ RXH
    data.fechaRxh,         // Z: FECHA EMISIÓN
    data.numRxh,           // AA: # RECIBO
    data.emisor,           // AB: NOMBRE DEL EMISOR
    data.fechaAbono,       // AC: FECHA DE ABONO
    "EN REVISIÓN CUENTAS", // AD: ESTADO
    "",                    // AE: OBSERVACION CUENTAS
    "",                    // AF: FECHA APROBACION CUENTAS
    "",                    // AG: APROBADO POR GERENTE
    "",                    // AH: OBSERVACION CONTABILIDAD
    "",                    // AI: FECHA APROBACION CONTABILIDAD
    "No",                  // AJ: APROBADO CONTABILIDAD
    "No"                   // AK: REPORTE DE GASTOS EMITIDO
  ]);

  // Forzar columnas L (# CUENTA) y M (# CCI) como texto para preservar ceros iniciales
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 12).setNumberFormat('@');  // L: # CUENTA
  sheet.getRange(lastRow, 13).setNumberFormat('@');  // M: # CCI

  return ContentService
    .createTextOutput(JSON.stringify({status:"ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var accion   = e.parameter.accion || 'dni';
  var callback = e.parameter.callback || ''; // JSONP support

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

  // ── SEGURIDAD: exigir usuario @ttaudit.com válido ──
  if (!verificar_(e.parameter.token)) {
    return respond({ error: 'no-autorizado' });
  }

  // ── AGREGAR PERSONA A BBDD (via GET/JSONP) ──
  if (accion === 'agregarPersona') {
    var bbdd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BBDD");
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
    return respond({status: 'ok'});
  }

  // ── BUSCAR POR DNI ──
  if (accion === 'dni') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BBDD");
    var data  = sheet.getDataRange().getValues();
    var dni   = e.parameter.dni;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).replace(/\.0$/,'') === String(dni)) {
        return respond({
          found: true, nombre: data[i][0], ciudad: data[i][2],
          ubigeo: data[i][3], departamento: data[i][3],
          banco: data[i][4], cuenta: String(data[i][5]), cci: String(data[i][6])
        });
      }
    }
    return respond({found:false});
  }

  // ── OBTENER BBDD COMPLETA (para carga masiva) ──
  if (accion === 'bbdd') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BBDD");
    var data  = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      result.push({
        nombre: row[0], dni: String(row[1]).replace(/\.0$/,''),
        ciudad: row[2], departamento: row[3],
        banco: row[4], cuenta: String(row[5]), cci: String(row[6])
      });
    }
    return respond({ personas: result });
  }

  // ── OBTENER PROYECTOS ──
  if (accion === 'proyectos') {
    var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROYECTOS");
    var rows   = sheet.getDataRange().getValues();
    var hdrs   = rows[0];
    var idxP   = {};
    for (var h = 0; h < hdrs.length; h++) {
      idxP[String(hdrs[h]).trim().toUpperCase()] = h;
    }
    function colP(row, nombre) {
      var i = idxP[String(nombre).trim().toUpperCase()];
      return (i !== undefined && row[i] !== undefined && row[i] !== null && row[i] !== '') ? String(row[i]) : '';
    }
    var result = [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row[1] && !row[2]) continue; // saltar filas vacías (cliente y proyecto vacíos)
      var fi = colP(row, 'FECHA INICIO') || colP(row, 'FECHA DE INICIO') || '';
      var fechaStr = '';
      if (fi) { try { fechaStr = new Date(fi).toLocaleDateString('es-PE'); } catch(err) { fechaStr = fi; } }
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
    return respond({proyectos: result});
  }

  // ── OBTENER REGISTROS ──
  if (accion === 'registros') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REGISTROS");
    var rows    = sheet.getDataRange().getValues();
    var headers = rows[0];
    var result  = [];

    // Mapear headers a índices dinámicamente
    var idx = {};
    for (var h = 0; h < headers.length; h++) {
      idx[String(headers[h]).trim().toUpperCase()] = h;
    }

    function col(row, nombre, fallback) {
      var i = idx[String(nombre).trim().toUpperCase()];
      return (i !== undefined && row[i] !== undefined && row[i] !== null && row[i] !== '') ? row[i] : (fallback !== undefined ? fallback : '');
    }

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row[0] && !row[1]) continue;
      result.push({
        fila:            i + 1,
        fecha:           row[0] ? new Date(row[0]).toLocaleDateString('es-PE') : '',
        quincena:        col(row, 'QUINCENA'),
        cliente:         col(row, 'CLIENTE'),
        proyecto:        col(row, 'PROYECTO'),
        supervisor:      col(row, 'SUPERVISOR'),
        ejCuentas:       col(row, 'EJECUTIVO DE CUENTAS') || col(row, 'EJ. CUENTAS') || col(row, 'EJECUTIVO CUENTAS') || col(row, 'GERENTE') || '',
        nombre:          col(row, 'APELLIDOS Y NOMBRE'),
        dni:             String(col(row, 'DNI')),
        departamento:    col(row, 'DEPARTAMENTO'),
        ciudad:          col(row, 'CIUDAD'),
        banco:           col(row, 'BANCO'),
        cuenta:          String(col(row, '# CUENTA') || col(row, '#CUENTA') || col(row, 'N° CUENTA') || col(row, 'CUENTA') || ''),
        cci:             String(col(row, '# CCI') || col(row, '#CCI') || col(row, 'N° CCI') || col(row, 'CCI') || ''),
        dias:            col(row, 'DÍAS TRABAJADOS'),
        pagoDia:         col(row, 'PAGO POR DÍA'),
        total:           String(col(row, 'TOTAL A PAGAR')),
        movAdicional:    col(row, 'MOVILIDAD ADICIONAL'),
        movViajera:      col(row, 'MOVILIDAD VIAJERA'),
        megas:           col(row, 'DATOS INTERNET'),
        otros:           col(row, 'OTROS CONCEPTOS'),
        adelanto:        col(row, 'ADELANTO MOV VIAJERA'),
        rxh:             col(row, 'EMITIÓ RXH'),
        numRxh:          (function() {
          // Probar todas las variantes posibles del header
          var raw = String(
            col(row, '# RECIBO') ||
            col(row, '#RECIBO')  ||
            col(row, 'N° RECIBO') ||
            col(row, 'RECIBO')   ||
            col(row, 'NUM RECIBO') ||
            col(row, 'NUMERO RECIBO') ||
            ''
          ).trim();
          if (!raw) return '';
          if (raw.indexOf('::') !== -1) {
            return raw.split('|').map(function(s){ return s.split('::')[0].trim(); }).filter(Boolean).join('|');
          }
          if (raw.indexOf(',') !== -1) {
            return raw.split(',').map(function(s){ return s.trim(); }).filter(Boolean).join('|');
          }
          return raw;
        })(),
        emisor:          col(row, 'NOMBRE EMISOR'),
        estado:          col(row, 'ESTADO'),
        obsCuentas:      col(row, 'OBS CUENTAS'),
        aprobadoPor:     col(row, 'APROBADO POR GERENTE'),
        obsCont:         col(row, 'OBS CONTABILIDAD'),
        obsContabilidad: col(row, 'OBS CONTABILIDAD')
      });
    }
    return respond({registros: result});
  }

  // ── ACTUALIZAR ESTADO ──
  if (accion === 'actualizar') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REGISTROS");
    var fila  = parseInt(e.parameter.fila);
    var tipo  = e.parameter.tipo; // 'cuentas' o 'contabilidad'
    var nuevoEstado = e.parameter.estado;
    var obs   = e.parameter.obs || '';
    var quien = e.parameter.quien || '';
    if (tipo === 'cuentas') {
      sheet.getRange(fila, 30).setValue(nuevoEstado);    // AD: estado
      sheet.getRange(fila, 31).setValue(obs);            // AE: obs cuentas
      sheet.getRange(fila, 32).setValue(new Date());     // AF: fecha aprobacion
      sheet.getRange(fila, 33).setValue(quien);          // AG: aprobado por
    } else if (tipo === 'correccion-monto') {
      // Supervisor corrige los conceptos tras observación de Cuentas
      var dias      = parseFloat(e.parameter.dias)      || 0;
      var pagoDia   = parseFloat(e.parameter.pagoDia)   || 0;
      var movDia    = parseFloat(e.parameter.movDia)    || 0;
      var movAd     = parseFloat(e.parameter.movAd)     || 0;
      var movVi     = parseFloat(e.parameter.movVi)     || 0;
      var adelanto  = parseFloat(e.parameter.adelanto)  || 0;
      var megas     = parseFloat(e.parameter.megas)     || 0;
      var otros     = parseFloat(e.parameter.otros)     || 0;
      var pagoDias  = dias * pagoDia;
      var movilidad = dias * movDia;
      var total     = parseFloat(e.parameter.total)     || 0;
      sheet.getRange(fila, 13).setValue(dias);           // M: DÍAS TRABAJADOS
      sheet.getRange(fila, 14).setValue(pagoDia);        // N: PAGO POR DÍA
      sheet.getRange(fila, 15).setValue(pagoDias);       // O: PAGO POR DÍAS TRABAJADOS
      sheet.getRange(fila, 16).setValue(movDia);         // P: PAGO MOVILIDAD DIARIA
      sheet.getRange(fila, 17).setValue(movilidad);      // Q: MOVILIDAD
      sheet.getRange(fila, 18).setValue(movAd);          // R: MOVILIDAD ADICIONAL
      sheet.getRange(fila, 19).setValue(movVi);          // S: MOVILIDAD VIAJERA
      sheet.getRange(fila, 20).setValue(megas);          // T: DATOS INTERNET
      sheet.getRange(fila, 21).setValue(otros);          // U: OTROS CONCEPTOS
      sheet.getRange(fila, 22).setValue(adelanto);       // V: ADELANTO MOV VIAJERA
      sheet.getRange(fila, 23).setValue(total);          // W: TOTAL A PAGAR
      sheet.getRange(fila, 30).setValue(nuevoEstado);    // AD: ESTADO
      sheet.getRange(fila, 31).setValue('');             // AE: limpiar obs cuentas
    } else if (tipo === 'rxh') {
      // X=emitió(24) Z=#recibo(26) AA=detalle(27) AC=estado(29)
      sheet.getRange(fila, 25).setValue('Sí');                      // Y: Emitió RxH
      sheet.getRange(fila, 27).setValue(e.parameter.numRxh || '');  // AA: # Recibo (legible)
      sheet.getRange(fila, 28).setValue(e.parameter.rxh || '');     // AB: Detalle completo
      sheet.getRange(fila, 30).setValue(nuevoEstado);               // AD: Estado
    } else if (tipo === 'revertir-contabilidad') {
      // Revertir aprobación de Contabilidad → vuelve a OBSERVADO-CONTABILIDAD
      sheet.getRange(fila, 30).setValue('OBSERVADO-CONTABILIDAD');  // AD: estado
      sheet.getRange(fila, 34).setValue(obs);                       // AH: obs contabilidad
      sheet.getRange(fila, 35).setValue(new Date());                // AI: fecha reversion
      sheet.getRange(fila, 36).setValue('No');                      // AJ: aprobado cont → No
    } else {
      sheet.getRange(fila, 30).setValue(nuevoEstado);               // AD: estado
      sheet.getRange(fila, 34).setValue(obs);                       // AH: obs contabilidad
      sheet.getRange(fila, 35).setValue(new Date());                // AI: fecha aprobacion cont
      sheet.getRange(fila, 36).setValue('Si');                      // AJ: aprobado cont
    }
    return respond({status:'ok'});
  }

  // ── ELIMINAR FILA ──────────────────────────────────────────
  if (accion === 'eliminar') {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTROS');
    var fila  = parseInt(e.parameter.fila);
    if (fila > 1) {
      sheet.deleteRow(fila);
    }
    return respond({status:'ok'});
  }

  return respond({error:'accion no reconocida'});
}
