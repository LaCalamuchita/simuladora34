/**
 * Simulador Area 34 · Registro de contactos
 * ------------------------------------------
 * Pegá este código en el Google Sheet del simulador:
 *   Extensiones → Apps Script → borrar lo que haya → pegar → Guardar.
 * Después: Implementar → Nueva implementación → Tipo: Aplicación web
 *   · Ejecutar como: Yo
 *   · Quién tiene acceso: Cualquier usuario
 * Copiá la URL que termina en /exec y pegala en index.html, en CONEXION.LEADS_URL.
 *
 * Qué hace cada vez que alguien toca "Recibir mi simulación":
 *   1. Agrega una fila en la pestaña "Leads".
 *   2. Si eligió email, le manda la simulación y las otras unidades que coinciden.
 *   3. Avisa por email al equipo comercial.
 * Cuenta gratuita de Google: hasta 100 emails por día.
 */

const EMAIL_COMERCIAL = 'comercial@area34.com';
const NOMBRE_EMPRESA = 'Area 34';
const HOJA_LEADS = 'Leads';
const COLUMNAS = ['fecha', 'nombre', 'canal', 'email', 'tipologia', 'zona', 'uso', 'proyecto', 'unidad',
                  'precio', 'anticipo', 'plazo', 'cuota', 'moneda', 'origen', 'codigo', 'link', 'evento'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(HOJA_LEADS) || ss.insertSheet(HOJA_LEADS);
    if (hoja.getLastRow() === 0) hoja.appendRow(COLUMNAS);

    if (d.tipo === 'evento') {            // por ahora: clics en "Agendar una llamada"
      hoja.appendRow(COLUMNAS.map(c => c === 'fecha' ? new Date() : c === 'evento' ? limpiar(d.evento) :
                                        c === 'nombre' ? limpiar(d.nombre) : c === 'origen' ? limpiar(d.origen) : ''));
      return ok();
    }
    if (!d.nombre) return ok();

    hoja.appendRow(COLUMNAS.map(c => c === 'fecha' ? new Date() : c === 'evento' ? 'simulacion' : limpiar(d[c])));

    const emailValido = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(d.email || ''));
    if (d.canal === 'email' && emailValido) {
      MailApp.sendEmail({
        to: d.email,
        replyTo: EMAIL_COMERCIAL,
        name: NOMBRE_EMPRESA,
        subject: `${d.proyecto} · ${d.unidad}: tu simulación y la ficha de la obra`,
        htmlBody: emailCliente(d),
      });
    }
    MailApp.sendEmail({
      to: EMAIL_COMERCIAL,
      name: 'Simulador web',
      subject: `Nuevo contacto: ${d.nombre} · ${d.tipologia} · ${d.canal}`,
      htmlBody: `<p><b>${esc(d.nombre)}</b> pidió su simulación por <b>${esc(d.canal)}</b>${d.email ? ' (' + esc(d.email) + ')' : ''}.</p>
        <p>${esc(d.proyecto)} · ${esc(d.unidad)} — anticipo ${dinero(d.anticipo, d.moneda)}, ${d.plazo} cuotas de ${dinero(d.cuota, d.moneda)}.</p>
        <p>Busca: ${esc(d.tipologia)} · Zona: ${esc(d.zona)} · Para: ${esc(d.uso)} · Origen: ${esc(d.origen || 'directo')}</p>
        ${d.link ? `<p><a href="${esc(d.link)}">Abrir su simulación</a></p>` : ''}`,
    });
    return ok();
  } catch (err) {
    console.error(err);
    return ok();
  } finally {
    lock.releaseLock();
  }
}

function emailCliente(d) {
  const filas = (d.opciones || []).map(o => `
    <tr><td style="padding:10px 0;border-top:1px solid #E2E0DA">${esc(o.nombre)}<br><span style="color:#6E6C66">${esc(o.tipologia)} · ${dinero(o.precio, d.moneda)}</span></td>
        <td style="padding:10px 0;border-top:1px solid #E2E0DA;text-align:right"><b>${dinero(o.cuota, d.moneda)}</b><br><span style="color:#6E6C66">${o.plazo} cuotas</span></td></tr>`).join('');
  return `<div style="font-family:Futura,Jost,Helvetica,Arial,sans-serif;color:#111;max-width:520px;margin:auto;padding:24px;background:#F6F5F2">
    <p style="letter-spacing:.3em;font-size:13px">AREA 34</p>
    <h1 style="font-weight:500;font-size:26px">Hola ${esc(d.nombre)}, esta es tu simulación.</h1>
    <div style="background:#fff;border:1px solid #E2E0DA;border-radius:16px;padding:20px">
      <p style="margin:0;color:#6E6C66">${esc(d.proyecto)} · ${esc(d.unidad)}</p>
      <p style="margin:8px 0;font-size:40px;color:#C65A3A">${dinero(d.cuota, d.moneda)} <span style="font-size:15px;color:#6E6C66">por mes</span></p>
      <p style="margin:0">Anticipo ${dinero(d.anticipo, d.moneda)} · ${d.plazo} cuotas fijas</p>
    </div>
    ${filas ? `<h2 style="font-weight:500;font-size:18px;margin-top:28px">Unidades que coinciden con tu búsqueda</h2><table style="width:100%;border-collapse:collapse;font-size:15px">${filas}</table>` : ''}
    ${d.ficha ? `<p style="margin-top:24px"><a href="${esc(d.ficha)}" style="display:inline-block;padding:14px 22px;border-radius:999px;background:#111;color:#fff;text-decoration:none">Ver la ficha de ${esc(d.proyecto)} (PDF)</a></p>` : ''}
    ${d.link ? `<p style="margin-top:16px"><a href="${esc(d.link)}" style="color:#111">Volver a abrir tu simulación</a></p>` : ''}
    <p style="margin-top:24px">Si querés, respondé este email y coordinamos una llamada de 15 minutos.</p>
    <p style="font-size:12px;color:#6E6C66">Simulación orientativa. No constituye una oferta. Precios y disponibilidad sujetos a cambios.</p>
  </div>`;
}

function ok() { return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function dinero(v, m) { return (m || 'USD') + ' ' + Number(v || 0).toLocaleString('es-AR', { maximumFractionDigits: 0 }); }
// Evita que un texto que empieza con = + - @ se interprete como fórmula en el Sheet
function limpiar(v) { if (v == null) return ''; if (typeof v === 'number') return v; const s = String(v).slice(0, 300); return /^[=+\-@]/.test(s) ? "'" + s : s; }
