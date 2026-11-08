'use strict';
/*********************************************************************************************
 * WENTEK STORE · Servidor en Vercel
 * -------------------------------------------------------------------------------------------
 * Hace lo mismo que Code.gs, pero hablando DIRECTO con tu hoja de Google (API de Sheets v4)
 * mediante una cuenta de servicio. Sin Apps Script en medio, sin dependencias externas.
 *
 * Variables de entorno (Vercel → Settings → Environment Variables):
 *   SHEET_ID             ID de la hoja (lo que va entre /d/ y /edit en su URL)
 *   GOOGLE_CLIENT_EMAIL  "client_email" del JSON de la cuenta de servicio
 *   GOOGLE_PRIVATE_KEY   "private_key" del JSON de la cuenta de servicio
 *   APP_TZ               (opcional) zona horaria, por defecto America/Bogota
 *********************************************************************************************/
const crypto = require('crypto');

const TZ   = process.env.APP_TZ || 'America/Bogota';
const API  = 'https://sheets.googleapis.com/v4/spreadsheets';
const USER = 'web';
const PAGOS = ['Efectivo', 'Datafono', 'Transferencia', 'Dañado'];


/* ===========================================================================================
 * 1 · AUTENTICACIÓN (cuenta de servicio → token de acceso)
 * =========================================================================================== */
let _tok = null, _tokExp = 0;
const b64u = b => Buffer.from(b).toString('base64').replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_');

async function token() {
  const now = Math.floor(Date.now() / 1000);
  if (_tok && now < _tokExp - 60) return _tok;

  const email = String(process.env.GOOGLE_CLIENT_EMAIL || '').trim();
  const key = String(process.env.GOOGLE_PRIVATE_KEY || '').trim()
    .replace(/^["']|["']$/g, '').replace(/\\n/g, '\n');
  if (!email || !key || !process.env.SHEET_ID) {
    throw new Error('Faltan variables de entorno en Vercel: SHEET_ID, GOOGLE_CLIENT_EMAIL y GOOGLE_PRIVATE_KEY.');
  }

  const head  = b64u(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const claim = b64u(JSON.stringify({
    iss: email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now, exp: now + 3600
  }));
  let sig;
  try { sig = crypto.createSign('RSA-SHA256').update(head + '.' + claim).sign(key); }
  catch (e) { throw new Error('La clave privada (GOOGLE_PRIVATE_KEY) no es válida. Pégala completa, tal cual está en el JSON.'); }

  const r = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: 'grant_type=' + encodeURIComponent('urn:ietf:params:oauth:grant-type:jwt-bearer') +
          '&assertion=' + head + '.' + claim + '.' + b64u(sig)
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error('Google rechazó la cuenta de servicio: ' + (j.error_description || j.error || r.status));
  _tok = j.access_token;
  _tokExp = now + (j.expires_in || 3600);
  return _tok;
}


/* ===========================================================================================
 * 2 · CAPA DE ACCESO A LA HOJA
 * =========================================================================================== */
const sleep = ms => new Promise(r => setTimeout(r, ms));
const q   = name => "'" + name + "'";
const enc = encodeURIComponent;

async function sheets(path, method, body, intento) {
  intento = intento || 0;
  const t = await token();
  const r = await fetch(API + '/' + process.env.SHEET_ID + path, {
    method: method || 'GET',
    headers: { Authorization: 'Bearer ' + t, 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined
  });
  const j = await r.json().catch(() => ({}));
  if (r.ok) return j;

  if ((r.status === 429 || r.status >= 500) && intento < 2) {
    await sleep(500 * (intento + 1));
    return sheets(path, method, body, intento + 1);
  }
  const m = (j.error && j.error.message) || ('HTTP ' + r.status);
  if (r.status === 403) throw new Error('Sin permiso sobre la hoja. Compártela como Editor con ' + process.env.GOOGLE_CLIENT_EMAIL + '. (' + m + ')');
  if (r.status === 404) throw new Error('No se encontró la hoja. Revisa SHEET_ID. (' + m + ')');
  if (/Unable to parse range/i.test(m)) throw new Error('Falta una pestaña en la hoja o tiene otro nombre. (' + m + ')');
  throw new Error('Error de Google Sheets: ' + m);
}

async function readRange(range) {
  const j = await sheets('/values/' + enc(range) + '?valueRenderOption=UNFORMATTED_VALUE');
  return j.values || [];
}

/** Lee varias zonas en UNA sola llamada (clave para la velocidad). */
async function batchGet(ranges) {
  const j = await sheets('/values:batchGet?valueRenderOption=UNFORMATTED_VALUE&' +
    ranges.map(r => 'ranges=' + enc(r)).join('&'));
  return (j.valueRanges || []).map(v => v.values || []);
}

const limpia = row => row.map(c => (c === undefined || c === null) ? '' : c);

async function writeRange(range, row) {
  await sheets('/values/' + enc(range) + '?valueInputOption=RAW', 'PUT',
    { range, majorDimension: 'ROWS', values: [limpia(row)] });
}

async function appendRow(range, row) {
  await sheets('/values/' + enc(range) + ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS', 'POST',
    { majorDimension: 'ROWS', values: [limpia(row)] });
}

async function batchWrite(data) {
  if (!data.length) return;
  await sheets('/values:batchUpdate', 'POST', {
    valueInputOption: 'RAW',
    data: data.map(d => ({ range: d.range, majorDimension: 'ROWS', values: d.values }))
  });
}

let _ids = null;
async function sheetId(name) {
  if (!_ids || _ids[name] === undefined) {
    const j = await sheets('?fields=sheets.properties(sheetId,title)');
    _ids = {};
    (j.sheets || []).forEach(s => { _ids[s.properties.title] = s.properties.sheetId; });
  }
  if (_ids[name] === undefined) throw new Error('No existe la pestaña "' + name + '".');
  return _ids[name];
}

async function deleteRow(name, row) {
  const id = await sheetId(name);
  await sheets(':batchUpdate', 'POST', { requests: [{
    deleteDimension: { range: { sheetId: id, dimension: 'ROWS', startIndex: row - 1, endIndex: row } }
  }] });
}

/** Fila (número de hoja) cuyo primer valor coincide; `col` viene de leer 'A2:A'. */
function rowOf(col, val) {
  for (let i = 0; i < col.length; i++) {
    if (col[i] && col[i][0] !== undefined && String(col[i][0]) === String(val)) return i + 2;
  }
  return 0;
}


/* ===========================================================================================
 * 3 · UTILIDADES
 * =========================================================================================== */
const hoyStr = () => new Date().toLocaleDateString('en-CA', { timeZone: TZ });

/** Filas de datos (sin encabezado y sin filas vacías). */
const rows = vals => vals.slice(1).filter(r => r.some(c => c !== '' && c !== undefined && c !== null));

/** Normaliza a yyyy-MM-dd. Acepta texto ISO, "yyyy-MM-dd HH:mm" o número de serie de Sheets. */
function fechaISO(v) {
  if (typeof v === 'number' && isFinite(v)) {
    return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  }
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  const m = s.match(/(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  const d = new Date(s);
  return isNaN(d) ? s : d.toLocaleDateString('en-CA', { timeZone: TZ });
}

const ventaObj = r => ({
  id: r[0],
  f:  fechaISO(r[1]),
  c:  String(r[2] ?? ''),
  q:  +r[3] || 0,
  p:  +r[4] || 0,
  t:  +r[5] || 0,
  m:  r[6],
  r:  r[7] === true || r[7] === 'TRUE',
  fa: +r[8] || 0
});

const clave = (k, n) => String((k[n] && k[n][0]) ?? '').trim();
const keyOK = (k, n, pass) => { const s = clave(k, n); return s !== '' && String(pass ?? '').trim() === s; };

const num = (x, nombre) => {
  const n = Number(x);
  if (!isFinite(n)) throw new Error('Valor inválido en ' + nombre + '.');
  return n;
};


/* ===========================================================================================
 * 4 · LECTURA
 * =========================================================================================== */

/** Unidades vendidas (s) y dañadas (d) por código desde su fecha de corte. */
function sold(inv, ventas) {
  const hoy = hoyStr();
  const corte = {};
  inv.forEach(p => { corte[p.c] = p.corte || hoy; });
  const s = {}, d = {};
  ventas.forEach(v => {
    if (corte[v.c] !== undefined && v.f >= corte[v.c]) {
      if (v.m === 'Dañado') d[v.c] = (d[v.c] || 0) + v.q;
      else                  s[v.c] = (s[v.c] || 0) + v.q;
    }
  });
  return { s, d };
}

async function getData(kind, val) {
  const esMes = (kind === 'ms');
  const key = esMes ? String(val || '').slice(0, 7) : fechaISO(val);
  const inRange = f => esMes ? f.slice(0, 7) === key : f === key;

  const [vd, ga, fa, iv, cf] = await batchGet([
    q('VENTA DIARIA'), q('Gastos'), q('Facturas'), q('Inventario'), q('Config') + '!A2:D2'
  ]);

  const todas  = rows(vd).map(ventaObj);
  const ventas = todas.filter(v => inRange(v.f));

  const gastos = rows(ga)
    .filter(r => inRange(fechaISO(r[1])))
    .map(r => ({ id: r[0], f: fechaISO(r[1]), d: r[2], v: +r[3] || 0 }));

  const facs = rows(fa)
    .filter(r => inRange(fechaISO(r[1])))
    .map(r => {
      let items = [];
      try { items = JSON.parse(r[7] || '[]'); } catch (e) {}
      return {
        n: +r[0], f: fechaISO(r[1]), cli: r[2], nit: String(r[3] ?? ''), mail: r[4],
        total: +r[5] || 0, pdf: r[6] || '', items, emp: r[8], enit: String(r[9] ?? '')
      };
    });

  const inv = rows(iv).map(r => ({
    c: String(r[0] ?? ''), n: r[1], q: +r[2] || 0, pm: +r[3] || 0, pv: +r[4] || 0,
    d: +r[5] || 0, corte: fechaISO(r[6])
  }));

  const c = cf[0] || [];
  const sd = sold(inv, todas);
  return {
    ventas, gastos, facs, inv,
    cfg: { base: +c[0] || 0, nf: +c[1] || 1, emp: String(c[2] ?? ''), nit: String(c[3] ?? '') },
    sold: sd.s,
    dam: sd.d
  };
}


/* ===========================================================================================
 * 5 · ESCRITURA
 * =========================================================================================== */

async function addVenta(v) {
  if (!v || !String(v.id || '') || !String(v.c || '').trim()) throw new Error('Venta incompleta.');
  const m = PAGOS.includes(v.m) ? v.m : 'Efectivo';
  const dn = m === 'Dañado';
  const qn = num(v.q, 'cantidad');
  if (!(qn > 0)) throw new Error('La cantidad debe ser mayor que cero.');
  const p = dn ? 0 : num(v.p, 'precio');
  const t = dn ? 0 : num(v.t, 'total');
  await appendRow(q('VENTA DIARIA') + '!A:J',
    [String(v.id), fechaISO(v.f) || hoyStr(), String(v.c).trim(), qn, p, t, m, false, '', USER]);
  return 1;
}

async function delVenta(f, id) {
  const [col] = await batchGet([q('VENTA DIARIA') + '!A2:A']);
  const r = rowOf(col, id);
  if (r) await deleteRow('VENTA DIARIA', r);
  return 1;
}

async function setRev(f, id) {
  const [col, rev] = await batchGet([q('VENTA DIARIA') + '!A2:A', q('VENTA DIARIA') + '!H2:H']);
  const r = rowOf(col, id);
  if (r) {
    const actual = rev[r - 2] && rev[r - 2][0];
    await writeRange(q('VENTA DIARIA') + '!H' + r, [!(actual === true || actual === 'TRUE')]);
  }
  return 1;
}

async function addGasto(g) {
  if (!g || !String(g.id || '') || !String(g.d || '').trim()) throw new Error('Gasto incompleto.');
  await appendRow(q('Gastos') + '!A:E',
    [String(g.id), fechaISO(g.f) || hoyStr(), String(g.d).trim(), num(g.v, 'valor'), USER]);
  return 1;
}

async function delGasto(id) {
  const [col] = await batchGet([q('Gastos') + '!A2:A']);
  const r = rowOf(col, id);
  if (r) await deleteRow('Gastos', r);
  return 1;
}

async function saveProd(p, pass) {
  const [k, col] = await batchGet([q('KEY') + '!A1:A2', q('Inventario') + '!A2:A']);
  if (!keyOK(k, 1, pass)) throw new Error('Contraseña de edición incorrecta.');
  if (!p || !String(p.c || '').trim()) throw new Error('Falta el código del producto.');
  const vals = [String(p.c).trim(), String(p.n ?? ''), num(p.q, 'existencias'), num(p.pm, 'precio de compra'),
                num(p.pv, 'precio de venta'), num(p.d, 'dañados'), fechaISO(p.corte) || hoyStr()];
  const r = rowOf(col, p.c);
  if (r) await writeRange(q('Inventario') + '!A' + r + ':G' + r, vals);
  else   await appendRow(q('Inventario') + '!A:G', vals);
  return 1;
}

async function delProd(c, pass) {
  const [k, col] = await batchGet([q('KEY') + '!A1:A2', q('Inventario') + '!A2:A']);
  if (!keyOK(k, 1, pass)) throw new Error('Contraseña de edición incorrecta.');
  const r = rowOf(col, c);
  if (r) await deleteRow('Inventario', r);
  return 1;
}

async function saveCfg(c) {
  if (!c) throw new Error('Configuración vacía.');
  await writeRange(q('Config') + '!A2:D2',
    [num(c.base, 'base'), num(c.nf, 'número de factura'), String(c.emp ?? ''), String(c.nit ?? '')]);
  return 1;
}


/* ===========================================================================================
 * 6 · FACTURACIÓN
 * (El PDF ya no se guarda en Drive: la factura queda registrada en la hoja y se imprime o se
 *  guarda como PDF desde el navegador con el botón "Imprimir".)
 * =========================================================================================== */

async function genFactura(d) {
  if (!d || typeof d !== 'object') throw new Error('No se recibieron los datos de la factura.');
  if (!String(d.cli || '').trim()) throw new Error('Debes indicar el nombre del cliente.');
  if (!Array.isArray(d.items) || !d.items.length) throw new Error('La factura no tiene productos.');
  if (!Array.isArray(d.ids) || !d.ids.length) throw new Error('La factura no tiene ventas seleccionadas.');

  const [cf, vcol] = await batchGet([q('Config') + '!A2:D2', q('VENTA DIARIA') + '!A2:A']);
  const cfg = cf[0] || [];
  let n = Number(cfg[1]);
  if (!isFinite(n) || n < 1) n = 1;

  const hoy     = hoyStr();
  const total   = Number(d.total || 0);
  const cliente = String(d.cli).trim();
  const nit     = String(d.nit || '').trim();
  const mail    = String(d.mail || '').trim();
  const emp     = String(cfg[2] ?? '');
  const enit    = String(cfg[3] ?? '');

  const items = d.items.map(i => ({
    c: String(i.c ?? ''), d: String(i.d ?? ''),
    q: Number(i.q || 0), p: Number(i.p || 0), t: Number(i.t || 0)
  }));

  await appendRow(q('Facturas') + '!A:J', [n, hoy, cliente, nit, mail, total, '', JSON.stringify(items), emp, enit]);

  // Marca las ventas incluidas y sube el contador, todo en una sola llamada
  const data = [];
  d.ids.forEach(x => {
    const id = x && x.id != null ? x.id : x;
    const r = rowOf(vcol, id);
    if (r) data.push({ range: q('VENTA DIARIA') + '!I' + r, values: [[n]] });
  });
  data.push({ range: q('Config') + '!B2', values: [[n + 1]] });
  await batchWrite(data);

  return { n, f: hoy, cli: cliente, nit, mail, total, pdf: '', items, emp, enit };
}

/** Quita la factura del registro y libera sus ventas. NO reutiliza el número. */
async function delFactura(n) {
  const [fcol, icol] = await batchGet([q('Facturas') + '!A2:A', q('VENTA DIARIA') + '!I2:I']);
  const r = rowOf(fcol, n);
  if (!r) return 0;

  const data = [];
  icol.forEach((c, i) => {
    if (c && c[0] !== undefined && String(c[0]) === String(n)) {
      data.push({ range: q('VENTA DIARIA') + '!I' + (i + 2), values: [['']] });
    }
  });
  await batchWrite(data);
  await deleteRow('Facturas', r);
  return 1;
}


/* ===========================================================================================
 * 7 · CONTRASEÑAS E INVERSIÓN
 * =========================================================================================== */

async function checkKey(pass)  { return keyOK(await readRange(q('KEY') + '!A1:A2'), 0, pass); }
async function checkKey2(pass) { return keyOK(await readRange(q('KEY') + '!A1:A2'), 1, pass); }

async function resumenInversion(pass) {
  const [k, iv, vd] = await batchGet([q('KEY') + '!A1:A2', q('Inventario'), q('VENTA DIARIA')]);
  if (!keyOK(k, 0, pass)) throw new Error('Contraseña incorrecta.');

  const costo = {};
  rows(iv).forEach(r => { costo[String(r[0] ?? '').toLowerCase()] = Number(r[3]) || 0; });

  const meses = {};
  rows(vd).forEach(r => {
    const ym = fechaISO(r[1]).slice(0, 7);
    if (!ym) return;
    const qn = Number(r[3]) || 0, p = Number(r[4]) || 0, t = Number(r[5]) || 0;
    const m = meses[ym] || (meses[ym] = { ym, venta: 0, ganancia: 0, sinCosto: 0, perdida: 0 });
    const c = String(r[2] ?? '').toLowerCase();
    if (String(r[6]) === 'Dañado') {            // dañado: pérdida al costo, no suma a ventas ni ganancia
      if (costo[c] !== undefined) m.perdida += qn * costo[c];
      return;
    }
    m.venta += t;
    if (costo[c] !== undefined) m.ganancia += qn * (p - costo[c]);
    else m.sinCosto += t;
  });

  return Object.keys(meses).sort().reverse().map(k2 => meses[k2]);
}


/* ===========================================================================================
 * 8 · PUNTO DE ENTRADA (POST /api/rpc  con  { fn, args })
 * =========================================================================================== */
const FN = {
  getData, addVenta, delVenta, setRev, addGasto, delGasto,
  saveProd, delProd, saveCfg, genFactura, delFactura,
  checkKey, checkKey2, resumenInversion
};

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method === 'GET') return res.status(200).json({ ok: true });
  if (req.method !== 'POST') return res.status(405).json({ error: 'Método no permitido.' });

  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    const fn = body.fn;
    if (typeof fn !== 'string' || !Object.prototype.hasOwnProperty.call(FN, fn)) {
      return res.status(400).json({ error: 'Operación desconocida.' });
    }
    const out = await FN[fn](...(Array.isArray(body.args) ? body.args : []));
    return res.status(200).json(out === undefined ? null : out);
  } catch (e) {
    return res.status(400).json({ error: (e && e.message) || String(e) });
  }
};

module.exports._test = { FN, fechaISO };
