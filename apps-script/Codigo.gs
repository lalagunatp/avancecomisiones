/* ==================================================================
   MI AVANCE DE INSTALACIONES — BACKEND (Apps Script)
   Todo vive en un único archivo: BASE LA LAGUNA 2026 (pestañas
   PLANTILLA, COMISIONES, RANKING, RANKING ENTRENAMIENTO, Accesos).
   El cliente (index.html) nunca conoce ese ID: solo conoce esta URL
   /exec. Todas las lecturas se hacen con SpreadsheetApp (identidad
   del propio script), así que el archivo no necesita estar
   compartido por enlace — nada se puede leer sin pasar por la app
   y sin PIN.
   ================================================================== */

const ID_BASE = '1Ph5T-m-Lkbdw1LBq-9wIIMW6C8bljOG1t5GfZQhNZ2o';   // BASE LA LAGUNA 2026 — único archivo
const HOJA_ACCESOS = 'Accesos';   // pestaña dentro de BASE LA LAGUNA 2026 para el log de accesos

/* 'comisiones'/'ranking'/'rankingEnt' -> pestaña real dentro de ID_BASE.
   Ya no se llama a ningún endpoint público: todo se lee con SpreadsheetApp,
   con la identidad del propio script, así que ningún archivo necesita estar
   compartido por enlace. */
const HOJAS_PERMITIDAS = {
  comisiones: {hoja: 'COMISIONES'},
  ranking:    {hoja: 'RANKING'},
  rankingEnt: {hoja: 'RANKING ENTRENAMIENTO'}
};

const PUESTOS_ACCESO = ['VENDEDOR','PROMOVENDEDOR PUNTO DE VENTA',
  'COACH VENTAS','COACH PROMOVENDEDOR PUNTO DE VENTA',
  'LIDER DE VENTAS','LIDER VENTAS','DIRECTOR DISTRITAL'];

const PIN_PRECARGADO   = '8888';
const VIGENCIA_TOKEN_MS = 4 * 60 * 60 * 1000;   // 4 horas

/* ---------------- entrada HTTP ---------------- */
function doGet(e){ return manejar(e); }
function doPost(e){ return manejar(e); }

function manejar(e){
  const p = (e && e.parameter) || {};
  let resultado;
  try{
    switch(p.accion){
      case 'login':      resultado = accLogin(p); break;
      case 'cambiarPin':  resultado = accCambiarPin(p); break;
      case 'directorio':  resultado = accDirectorio(p); break;
      case 'relay':       resultado = accRelay(p); break;
      case 'registrar':   resultado = accRegistrar(p); break;
      case 'consultar':   resultado = accConsultarAccesos(p); break;
      default: resultado = {ok:false, error:'acción desconocida'};
    }
  }catch(err){
    resultado = {ok:false, error:String((err && err.message) || err)};
  }
  return salida(resultado, p.callback);
}

function salida(obj, callback){
  const json = JSON.stringify(obj);
  const texto = callback ? (callback + '(' + json + ')') : json;
  return ContentService.createTextOutput(texto)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

/* ---------------- utilerías ---------------- */
const norm = s => String(s==null?'':s).toLowerCase()
  .replace(/[áàäâ]/g,'a').replace(/[éèëê]/g,'e').replace(/[íìïî]/g,'i')
  .replace(/[óòöô]/g,'o').replace(/[úùüû]/g,'u').replace(/ñ/g,'n')
  .replace(/\s+/g,' ').trim();
const limpio = v => String(v==null?'':v).replace(/\.0+$/,'').trim();
const listaTiene = (lista, v) => lista.some(x => norm(x) === norm(v));

function secreto(){
  const s = PropertiesService.getScriptProperties().getProperty('SESSION_SECRET');
  if(!s) throw new Error('Falta configurar SESSION_SECRET en Script Properties');
  return s;
}

/* Token firmado tipo JWT simplificado: nadie sin el secreto puede fabricar
   uno válido ni estirarle la fecha de vencimiento. */
function firmar(payload){
  const b64 = Utilities.base64EncodeWebSafe(JSON.stringify(payload));
  const firma = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(b64, secreto()));
  return b64 + '.' + firma;
}

function verificarToken(token){
  if(!token || token.indexOf('.') < 0) throw new Error('sesión inválida, vuelve a entrar');
  const partes = token.split('.');
  const b64 = partes[0], firma = partes[1];
  const esperada = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(b64, secreto()));
  if(firma !== esperada) throw new Error('sesión inválida, vuelve a entrar');
  const payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(b64)).getDataAsString());
  if(!payload.exp || Date.now() > payload.exp) throw new Error('sesión vencida, vuelve a entrar');
  return payload;
}

/* ---------------- directorio (hoja PLANTILLA) ---------------- */
let _plantillaCache = null;
function leerPlantilla(){
  if(_plantillaCache) return _plantillaCache;
  const hoja = SpreadsheetApp.openById(ID_BASE).getSheetByName('PLANTILLA');
  const valores = hoja.getDataRange().getValues();
  const filas = [];
  for(let i = 1; i < valores.length; i++){        // fila 1 = encabezados
    const r = valores[i];
    filas.push({
      empSF:      limpio(r[1]),   // B — NUMERO EMPLEADO SF
      idPos:      limpio(r[2]),   // C — ID POSICIONES
      empTGS:     limpio(r[3]),   // D — NUMERO EMPLEADO TGS
      nombre:     limpio(r[4]),   // E — NOMBRE DEL EMPLEADO
      puesto:     limpio(r[5]),   // F — POSICIÓN
      distrito:   limpio(r[7]),   // H — DISTRITO
      lider:      limpio(r[8]),   // I — NUMERO LR
      pin:        limpio(r[27]),  // AB — PIN
      homologado: limpio(r[30]),  // AE — HOMOLOGAD
      _fila: i + 1
    });
  }
  _plantillaCache = filas;
  return filas;
}

function idDe(f){ return f.homologado || f.empSF || f.empTGS || f.idPos; }

function buscarPorNumero(numero){
  const n = norm(numero);
  if(!n) return null;
  return leerPlantilla().find(f =>
    norm(f.empSF) === n || norm(f.idPos) === n || norm(f.empTGS) === n || norm(f.homologado) === n
  ) || null;
}

/* ---------------- reglas del PIN nuevo ---------------- */
function pinInvalido(pin, numeroEmpleado){
  if(!/^\d{4}$/.test(pin)) return 'El PIN debe ser de 4 dígitos.';
  if(/^(\d)\1{3}$/.test(pin)) return 'El PIN no puede ser el mismo dígito repetido.';
  const d = pin.split('').map(Number);
  let asc = true, desc = true;
  for(let i = 1; i < d.length; i++){
    if(d[i] !== d[i-1] + 1) asc = false;
    if(d[i] !== d[i-1] - 1) desc = false;
  }
  if(asc || desc) return 'El PIN no puede ser una secuencia consecutiva (ej. 1234).';
  const emp = limpio(numeroEmpleado).replace(/\D/g,'');
  if(pin === limpio(numeroEmpleado) || (emp && pin === emp.slice(-4)))
    return 'El PIN no puede ser igual a tu número de empleado.';
  if(pin === PIN_PRECARGADO) return 'Elige un PIN distinto al precargado.';
  return '';
}

/* ---------------- acciones ---------------- */
function accLogin(p){
  const fila = buscarPorNumero(p.numero || '');
  if(!fila || norm(fila.nombre) === 'vacante') return {ok:false, error:'noAutorizado'};
  if(!listaTiene(PUESTOS_ACCESO, fila.puesto)) return {ok:false, error:'sinPuesto'};
  const pinEnviado = limpio(p.pin || '');
  if(!fila.pin || pinEnviado !== fila.pin) return {ok:false, error:'pinIncorrecto'};

  const numero = idDe(fila);
  const payload = {
    numero, homologado: fila.homologado, posicion: fila.puesto,
    distrito: fila.distrito, exp: Date.now() + VIGENCIA_TOKEN_MS
  };
  return {
    ok: true,
    token: firmar(payload),
    debeCambiarPin: fila.pin === PIN_PRECARGADO,
    numero, nombre: fila.nombre, distrito: fila.distrito,
    posicion: fila.puesto, homologado: fila.homologado
  };
}

function accCambiarPin(p){
  const payload = verificarToken(p.token);
  const fila = buscarPorNumero(payload.numero);
  if(!fila) return {ok:false, error:'noAutorizado'};
  const nuevo = limpio(p.pinNuevo || '');
  const err = pinInvalido(nuevo, payload.numero);
  if(err) return {ok:false, error:err};

  const hoja = SpreadsheetApp.openById(ID_BASE).getSheetByName('PLANTILLA');
  hoja.getRange(fila._fila, 28).setValue(nuevo);   // columna AB
  _plantillaCache = null;

  const fresco = Object.assign({}, payload, {exp: Date.now() + VIGENCIA_TOKEN_MS});
  return {ok:true, token: firmar(fresco)};
}

function accDirectorio(p){
  verificarToken(p.token);
  const datos = leerPlantilla()
    .filter(f => idDe(f) && norm(f.nombre) !== 'vacante')
    .map(f => [idDe(f), f.nombre, f.distrito, f.puesto, f.lider, '', f.homologado]);
  return {ok:true, datos};
}

/* Igual que armarArbol()/descendencia() del cliente, pero server-side:
   determina qué números puede pedir un token (él mismo + su línea hacia
   abajo), para que el relay no entregue cuentas ajenas aunque el cliente
   mande una consulta manipulada. */
function descendenciaAutorizada(numeroRaiz){
  const filas = leerPlantilla().filter(f => listaTiene(PUESTOS_ACCESO, f.puesto));
  const hijos = new Map();
  filas.forEach(f => {
    if(!f.lider) return;
    const k = norm(f.lider);
    if(!hijos.has(k)) hijos.set(k, []);
    hijos.get(k).push(f);
  });
  const raiz = norm(numeroRaiz);
  const vistos = new Set([raiz]);
  const pila = [raiz];
  while(pila.length){
    const actual = pila.pop();
    (hijos.get(actual) || []).forEach(f => {
      const id = norm(idDe(f));
      if(!id || vistos.has(id)) return;
      vistos.add(id); pila.push(id);
    });
  }
  return vistos;
}

/* ==================================================================
   INTÉRPRETE DE CONSULTAS TIPO GVIZ
   El cliente sigue armando cadenas 'select B,C where C = 123 limit 5'
   (igual que cuando se leía Google Sheets directo); en vez de
   reenviarlas al endpoint público de GViz, se resuelven aquí mismo
   contra lo que ya leyó SpreadsheetApp. Son solo 9 variantes fijas en
   todo el cliente, todas con esta forma:
     select <col,col,...|*|count(col)> [where <cond>] [limit N] [offset N]
   ================================================================== */
function letraAIndice(letra){
  let n = 0;
  const s = String(letra).toUpperCase();
  for(let i = 0; i < s.length; i++) n = n*26 + (s.charCodeAt(i) - 64);
  return n - 1;
}
function indiceALetra(indice){
  let n = indice + 1, s = '';
  while(n > 0){ const r = (n-1) % 26; s = String.fromCharCode(65+r) + s; n = Math.floor((n-1)/26); }
  return s;
}

/* Corta desde la derecha: en las 9 variantes que arma el cliente, cuando
   aparecen 'where'/'limit'/'offset' siempre van en ese orden, así que
   cortar por la derecha evita depender de una regex con grupos opcionales
   frágiles. */
function parseTq(tqOriginal){
  const tq = String(tqOriginal || '').trim();
  if(tq.slice(0,7).toLowerCase() !== 'select ') throw new Error('consulta no reconocida: '+tqOriginal);
  let resto = tq.slice(7);
  let restoBajo = resto.toLowerCase();
  let limit = null, offset = null, where = null;

  const cortar = (palabra) => {
    const i = restoBajo.lastIndexOf(' '+palabra+' ');
    if(i < 0) return null;
    const valor = resto.slice(i + palabra.length + 2).trim();
    resto = resto.slice(0, i);
    restoBajo = restoBajo.slice(0, i);
    return valor;
  };

  const vOffset = cortar('offset'); if(vOffset != null) offset = parseInt(vOffset, 10);
  const vLimit  = cortar('limit');  if(vLimit  != null) limit  = parseInt(vLimit, 10);
  const vWhere  = cortar('where');  if(vWhere  != null) where  = vWhere.trim();

  return {select: resto.trim(), where, limit, offset};
}

function parseSelect(sel){
  sel = sel.trim();
  if(sel === '*') return {tipo:'todas'};
  const mCount = /^count\(([A-Za-z]+)\)$/i.exec(sel);
  if(mCount) return {tipo:'count', col: mCount[1].toUpperCase()};
  return {tipo:'columnas', cols: sel.split(',').map(c => c.trim().toUpperCase())};
}

/* Las 5 formas de condición atómica que arma el cliente */
function parseCondicionSimple(cond){
  cond = cond.trim();
  let m;
  if((m = /^upper\(([A-Za-z]+)\)\s+contains\s+'(.*)'$/i.exec(cond)))
    return {col:m[1].toUpperCase(), op:'contains', val:m[2], mayus:true};
  if((m = /^upper\(([A-Za-z]+)\)\s*=\s*'(.*)'$/i.exec(cond)))
    return {col:m[1].toUpperCase(), op:'=', val:m[2], mayus:true};
  if((m = /^([A-Za-z]+)\s+contains\s+'(.*)'$/i.exec(cond)))
    return {col:m[1].toUpperCase(), op:'contains', val:m[2]};
  if((m = /^([A-Za-z]+)\s*=\s*'(.*)'$/i.exec(cond)))
    return {col:m[1].toUpperCase(), op:'=', val:m[2]};
  if((m = /^([A-Za-z]+)\s*=\s*(-?\d+(?:\.\d+)?)$/.exec(cond)))
    return {col:m[1].toUpperCase(), op:'=', val:m[2]};
  throw new Error('condición no reconocida: '+cond);
}

function evaluarWhere(whereStr, fila){
  let s = whereStr.trim();
  if(s[0] === '(' && s[s.length-1] === ')') s = s.slice(1,-1);
  const partes = s.split(/\s+or\s+/i);
  return partes.some(parte => {
    const c = parseCondicionSimple(parte);
    const v = limpio(fila[letraAIndice(c.col)]);
    if(c.op === '=') return norm(v) === norm(c.val) || v === c.val;
    if(c.op === 'contains') return v.toUpperCase().indexOf(c.val.toUpperCase()) > -1;
    return false;
  });
}

/* GViz representa una fecha como el texto "Date(2026,0,15)", que
   aFecha() (index.html) ya sabe parsear. SpreadsheetApp da objetos Date
   de verdad: si se mandaran tal cual, JSON los convierte a ISO-UTC y se
   puede correr un día según la zona horaria del teléfono. Se arma el
   mismo formato de texto con los getters normales de Date (misma zona
   horaria del proyecto), para que el cliente no note la diferencia. */
function formatearValor(v){
  if(v instanceof Date) return 'Date('+v.getFullYear()+','+v.getMonth()+','+v.getDate()+')';
  return (v === '' ? null : v);
}

let _hojaCache = {};
function leerHoja(nombre){
  if(_hojaCache[nombre]) return _hojaCache[nombre];
  const hoja = SpreadsheetApp.openById(ID_BASE).getSheetByName(nombre);
  if(!hoja) throw new Error('no existe la pestaña "'+nombre+'"');
  const valores = hoja.getDataRange().getValues();
  _hojaCache[nombre] = valores;
  return valores;
}

function ejecutarConsulta(tq, nombreHoja, headers){
  const datos = leerHoja(nombreHoja);
  const {select, where, limit, offset} = parseTq(tq);
  const sel = parseSelect(select);

  const conEncabezado = !(headers == 0 || headers === '0');
  const encabezados = conEncabezado ? datos[0] : null;
  let filas = datos.slice(conEncabezado ? 1 : 0);

  if(where) filas = filas.filter(f => evaluarWhere(where, f));

  filas = filas.slice(offset || 0);
  if(limit != null) filas = filas.slice(0, limit);

  if(sel.tipo === 'count'){
    const idx = letraAIndice(sel.col);
    const n = filas.filter(f => limpio(f[idx]) !== '').length;
    return {cols:[{id:'count', label:'', type:'number'}], rows:[{c:[{v:n}]}]};
  }

  let letras = sel.tipo === 'todas'
    ? Array.from({length: datos.reduce((m,f) => Math.max(m, f.length), 0)}, (_, i) => indiceALetra(i))
    : sel.cols;

  const cols = letras.map(l => ({
    id: l,
    label: encabezados ? limpio(encabezados[letraAIndice(l)]) : '',
    type: 'string'
  }));
  const rows = filas.map(f => ({
    c: letras.map(l => ({v: formatearValor(f[letraAIndice(l)])}))
  }));
  return {cols, rows};
}

/* Relay: interpreta la MISMA consulta que ya arma el cliente, pero contra
   lo que lee SpreadsheetApp con la identidad del propio script — el
   archivo no necesita estar compartido por enlace — y solo si el token
   trae permiso sobre los números que la consulta menciona. */
function accRelay(p){
  const payload = verificarToken(p.token);
  const cfg = HOJAS_PERMITIDAS[p.hoja];
  if(!cfg) return {ok:false, error:'hoja no permitida'};

  const tq = String(p.tq || '');
  if(!autorizarConsulta(tq, payload)) return {ok:false, error:'no autorizado para esa consulta'};

  try{
    const tabla = ejecutarConsulta(tq, cfg.hoja, p.headers == null ? 1 : p.headers);
    return {ok:true, tabla};
  }catch(err){
    return {ok:false, error:'CONSULTA: ' + String((err && err.message) || err)};
  }
}

/* Si el WHERE trae números de 4+ dígitos (claves de empleado), deben estar
   dentro de "yo mismo o mi descendencia". Consultas sin números (ej. por
   distrito, o "select count(...)") se dejan pasar: no traen datos de una
   persona puntual fuera de lo que la propia app ya filtra. */
function autorizarConsulta(tq, payload){
  const m = tq.match(/where\s+(.+)$/i);
  if(!m) return true;
  const numeros = (m[1].match(/-?\d+/g) || []).filter(n => n.length >= 4);
  if(!numeros.length) return true;
  const permitidos = descendenciaAutorizada(payload.numero);
  return numeros.every(n => permitidos.has(norm(n)));
}

/* ---------------- registro de accesos ---------------- */
function accRegistrar(p){
  hojaAccesos().appendRow([new Date(), p.num || '', p.nombre || '', p.puesto || '', p.distrito || '']);
  return {ok:true};
}

function accConsultarAccesos(){
  const valores = hojaAccesos().getDataRange().getValues();
  const porNumero = new Map();
  for(let i = 1; i < valores.length; i++){
    const [fecha, num] = valores[i];
    if(!num) continue;
    const k = String(num);
    const prev = porNumero.get(k);
    if(!prev) porNumero.set(k, {primera: fecha, ultima: fecha, total: 1});
    else{
      prev.total++;
      if(fecha < prev.primera) prev.primera = fecha;
      if(fecha > prev.ultima) prev.ultima = fecha;
    }
  }
  const datos = [];
  porNumero.forEach((v, k) => datos.push([k, v.ultima, v.total, v.primera]));
  return {ok:true, datos};
}

function hojaAccesos(){
  const libro = SpreadsheetApp.openById(ID_BASE);
  let hoja = libro.getSheetByName(HOJA_ACCESOS);
  if(!hoja){
    hoja = libro.insertSheet(HOJA_ACCESOS);
    hoja.appendRow(['Fecha','Número','Nombre','Puesto','Distrito']);
  }
  return hoja;
}
