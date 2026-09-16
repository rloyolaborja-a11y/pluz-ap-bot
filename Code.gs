/* ============================================================================
   Puente hacia GitHub (Reporte de Pendientes/Atendidas AP - Pluz)
   ============================================================================
   Esto recibe archivos desde tus herramientas locales (PUBLICAR-WEB-LOCAL,
   PENDIENTES-LOCAL, ATENDIDAS-LOCAL) y desde la propia página ("Procesar y
   actualizar"), y los publica en tu repositorio de GitHub, a través de la
   Apps Script de Google (el "puente" hacia GitHub) — quien le habla a GitHub
   es Google (este script), no tu laptop ni el navegador de quien visita el
   sitio, así ninguna restricción de red de tu empresa afecta el proceso.

   IMPORTANTE (2026-08-27): todos los archivos que llegan en una sola llamada
   se publican como UN SOLO commit (usando la Git Data API de GitHub en vez
   de la Contents API), en vez de un commit por archivo. Esto es a propósito:
   cada commit en la rama "main" dispara un despliegue nuevo en Vercel, y
   Vercel tiene un límite diario de cuántos despliegues puede hacer — subir
   38 (o 10) archivos como 38 (o 10) commits separados agotaba esa cuota en
   un par de publicaciones. Con un solo commit por publicación, cada corrida
   de Publicar.bat / Ejecutar.bat / "Procesar y actualizar" cuenta como UN
   solo despliegue para Vercel, sin importar cuántos archivos cambien.

   CONFIGURACIÓN (una sola vez, ver LEEME-APPS-SCRIPT.txt):
   1. Project Settings (el ícono de engranaje, a la izquierda) → Script
      Properties → agrega dos propiedades:
        GITHUB_TOKEN = tu token personal de GitHub (permiso Contents: Read
                        and write sobre el repositorio)
        SECRET       = una clave que tú inventes (o la que te dé Claude) —
                        sirve para que nadie más pueda usar esta URL.
   2. Deploy → New deployment → tipo "Web app" → Execute as: Me (tu cuenta)
      → Who has access: Anyone → Deploy.
   3. Copia la URL que termina en /exec — esa va en el "script_url" de cada
      config.json de tus 3 carpetas locales (y en assets/js/gh-token.js del
      sitio, para "Procesar y actualizar" desde la página).

   El token de GitHub queda SOLO acá, guardado en Script Properties — nunca
   más viaja a ninguna laptop.

   2026-09-05: se agregó accion:"correccion" + doGet ?path=correcciones para
   la pestaña "Correcciones" de Atendidas AP — guarda/lee las correcciones
   manuales de Actividad/Rubro en una Google Sheet (ver bloque al final).
   ========================================================================== */

var OWNER = "rloyolaborja-a11y";
var REPO = "REPORTE-PENDIENTES-AP";
var BRANCH = "main";

// Carpeta de Drive donde se guardan los DATOS (bd_actual.json, bd_completa.json,
// etc.) cuando se publican con accion:"datos". Esto es lo que reemplaza a
// GitHub/Vercel para los datos (2026-08-29): la página los lee en vivo desde
// acá vía doGet, así que actualizar datos ya NO dispara ningún despliegue en
// Vercel — solo los cambios de código del sitio (accion:"sitio", el mismo
// flujo de siempre) siguen yendo por GitHub.
var CARPETA_DATOS_NOMBRE = "PortalAP_Datos";

// 2026-09-05: Google Sheet donde viven las correcciones manuales de
// Actividad/Rubro (pestaña "Correcciones" de Atendidas AP). Se usa una hoja
// de cálculo y NO un archivo JSON en Drive porque acá editan varias personas
// (cada contratista + Pluz) y Google serializa las escrituras a la hoja (con
// LockService) — un archivo JSON se pisaría entre dos que guardan a la vez.
var SHEET_CORRECCIONES_ID = "1-PQdHiRPvOPww3aJveAN8Yhm8oiAmuHZRaC5Yo286tQ";
var SHEET_CORRECCIONES_HOJA = "correcciones";
var COLS_CORRECCION = [
  "reclamo", "contratista", "distrito", "odm", "observacion",
  "actividad", "rubro", "modificado_por", "fecha_modificacion",
  "validado", "validado_por", "fecha_validacion"
];

function doGet(e) {
  var path = e && e.parameter && (e.parameter.path || e.parameter.tipo);
  if (!path) {
    return respond({ ok: false, error: "Falta el parámetro 'path' (ej: ?path=atendidas/data/bd_actual.json)." }, 400);
  }

  // 2026-09-05: lectura de las correcciones manuales (pestaña "Correcciones"
  // de Atendidas y el pipeline). Devuelve { rows: [...] } igual que los otros
  // datos, para que leerDatoRemoto() lo consuma igual.
  if (path === "correcciones") {
    try {
      return ContentService.createTextOutput(JSON.stringify(leerCorrecciones_()))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return respond({ ok: false, error: String(err) }, 500);
    }
  }

  try {
    var archivo = buscarArchivoDatos_(path, false);
    if (!archivo) {
      return respond({ ok: false, error: "No hay datos guardados todavía para: " + path }, 404);
    }
    // Se devuelve el JSON tal cual se guardó (no envuelto), para que el
    // fetch() de la página lo pueda usar directo, igual que hoy usa los
    // archivos estáticos data/*.json.
    return ContentService.createTextOutput(archivo.getBlob().getDataAsString("UTF-8")).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return respond({ ok: false, error: String(err) }, 500);
  }
}

function doPost(e) {
  var props = PropertiesService.getScriptProperties();
  var SECRET = props.getProperty("SECRET");

  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond({ ok: false, error: "Cuerpo de la petición inválido (no es JSON)." }, 400);
  }

  if (!SECRET) {
    return respond({ ok: false, error: "Falta configurar la propiedad SECRET en este script." }, 500);
  }
  if (body.secret !== SECRET) {
    return respond({ ok: false, error: "Clave incorrecta." }, 401);
  }

  var mensaje = body.message || ("Actualizar (Apps Script) - " + new Date().toISOString());
  var accion = body.accion || "sitio"; // "sitio" (GitHub) | "datos" (Drive) | "correccion" (Google Sheet)

  // 2026-09-05: guardar/actualizar correcciones manuales de Actividad/Rubro
  // en la Google Sheet. Cada item: { reclamo, contratista, distrito, odm,
  // observacion, actividad, rubro, modificado_por, validado (bool),
  // validado_por }. Upsert por "reclamo" (clave única).
  if (accion === "correccion") {
    try {
      // Las correcciones vienen en base64 (correcciones_b64) para que las
      // tildes/Ñ no se corrompan — igual que content_base64 de los datos.
      // Se acepta también "correcciones" (array plano) por compatibilidad.
      var items;
      if (body.correcciones_b64) {
        var txt = Utilities.newBlob(Utilities.base64Decode(body.correcciones_b64)).getDataAsString("UTF-8");
        items = JSON.parse(txt);
      } else {
        items = body.correcciones || [];
      }
      if (!items || !items.length) {
        return respond({ ok: false, error: "No se enviaron correcciones." }, 400);
      }
      var n = guardarCorrecciones_(items);
      return respond({ ok: true, guardados: n }, 200);
    } catch (err) {
      return respond({ ok: false, error: String(err) }, 200);
    }
  }

  var archivos = body.files || [];
  if (!archivos.length) {
    return respond({ ok: false, error: "No se enviaron archivos (files vacío)." }, 400);
  }

  if (accion === "datos") {
    try {
      archivos.forEach(function (f) {
        if (!f.path || !f.content_base64) throw new Error("Falta path o content_base64 en un archivo.");
        guardarArchivoDatos_(f.path, f.content_base64);
      });
      var resultadosDatos = archivos.map(function (f) { return { path: f.path, status: "OK" }; });
      return respond({ ok: true, publicados: archivos.length, total: archivos.length, resultados: resultadosDatos, modo: "datos (Drive, sin despliegue)" }, 200);
    } catch (err) {
      return respond({ ok: false, error: String(err) }, 200);
    }
  }

  var GITHUB_TOKEN = props.getProperty("GITHUB_TOKEN");
  if (!GITHUB_TOKEN) {
    return respond({ ok: false, error: "Falta configurar la propiedad GITHUB_TOKEN en este script." }, 500);
  }

  try {
    var commitSha = publicarUnSoloCommit(GITHUB_TOKEN, archivos, mensaje);
    var resultados = archivos.map(function (f) { return { path: f.path, status: "OK" }; });
    return respond({ ok: true, publicados: archivos.length, total: archivos.length, resultados: resultados, commit: commitSha }, 200);
  } catch (err) {
    // Si algo falla a mitad de camino (blob, tree o commit), NINGÚN archivo
    // queda a medias en el repo — no se movió la rama todavía en ese punto.
    var resultados = archivos.map(function (f) { return { path: f.path, status: "FALLO", error: String(err) }; });
    return respond({ ok: false, error: String(err), resultados: resultados }, 200);
  }
}

// ---------------------------------------------------------------------------
// Almacén de DATOS en Drive (reemplaza a GitHub solo para los archivos de
// datos: bd_actual.json, bd_completa.json y sus variantes por contratista).
// Un archivo de Drive por "path" (ej. "atendidas/data/bd_actual.json"), con
// el nombre sanitizado (los "/" se cambian por "__"), siempre en la misma
// carpeta. Se SOBRESCRIBE el contenido en cada publicación — no se acumulan
// versiones viejas.
// ---------------------------------------------------------------------------

function nombreSanitizado_(path) {
  return path.replace(/\//g, "__");
}

function carpetaDatos_() {
  var carpetas = DriveApp.getFoldersByName(CARPETA_DATOS_NOMBRE);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(CARPETA_DATOS_NOMBRE);
}

function buscarArchivoDatos_(path) {
  var nombre = nombreSanitizado_(path);
  var archivos = carpetaDatos_().getFilesByName(nombre);
  return archivos.hasNext() ? archivos.next() : null;
}

function guardarArchivoDatos_(path, contentBase64) {
  var nombre = nombreSanitizado_(path);
  var carpeta = carpetaDatos_();
  var bytes = Utilities.base64Decode(contentBase64);
  var blob = Utilities.newBlob(bytes, "application/json", nombre);
  var existente = buscarArchivoDatos_(path);
  if (existente) {
    existente.setContent(blob.getDataAsString("UTF-8"));
  } else {
    carpeta.createFile(blob);
  }
}

function respond(obj, httpStatusHint) {
  // Apps Script Web Apps siempre responden 200 al llamador (no se puede fijar
  // el código HTTP real de la respuesta) — por eso el estado real va DENTRO
  // del JSON (campo "ok" y, en los items, "status"). httpStatusHint queda
  // solo de referencia/documentación.
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ghFetch(url, options) {
  options = options || {};
  options.headers = options.headers || {};
  options.headers.Authorization = "Bearer " + options.token;
  options.headers.Accept = "application/vnd.github+json";
  options.muteHttpExceptions = true;
  delete options.token;
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  if (code >= 400) {
    throw new Error("GitHub " + (options.method || "get").toUpperCase() + " " + url.replace("https://api.github.com", "") + " -> " + code + ": " + resp.getContentText().slice(0, 300));
  }
  return JSON.parse(resp.getContentText());
}

// Publica TODOS los archivos recibidos como un único commit en la rama
// BRANCH, usando la Git Data API (blobs + tree + commit + mover la rama) en
// vez de la Contents API (que crea un commit por archivo). Devuelve el SHA
// del commit nuevo.
function publicarUnSoloCommit(token, archivos, mensaje) {
  var base = "https://api.github.com/repos/" + OWNER + "/" + REPO;

  // 1) SHA del commit actual en la punta de la rama.
  var ref = ghFetch(base + "/git/refs/heads/" + BRANCH, { token: token, method: "get" });
  var commitActualSha = ref.object.sha;

  // 2) SHA del árbol (tree) de ese commit, como base para el árbol nuevo —
  // así los archivos que NO cambian en esta publicación se mantienen tal
  // cual, sin tener que reenviarlos.
  var commitActual = ghFetch(base + "/git/commits/" + commitActualSha, { token: token, method: "get" });
  var treeBaseSha = commitActual.tree.sha;

  // 3) Un blob por archivo (no dispara ningún commit/despliegue por sí solo).
  var entradasArbol = archivos.map(function (f) {
    if (!f.path || !f.content_base64) throw new Error("Falta path o content_base64 en un archivo.");
    var blob = ghFetch(base + "/git/blobs", {
      token: token,
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ content: f.content_base64, encoding: "base64" }),
    });
    return { path: f.path, mode: "100644", type: "blob", sha: blob.sha };
  });

  // 4) Árbol nuevo = árbol base + los blobs de arriba (reemplazando esos
  // paths puntuales; el resto del árbol queda igual).
  var treeNuevo = ghFetch(base + "/git/trees", {
    token: token,
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ base_tree: treeBaseSha, tree: entradasArbol }),
  });

  // 5) Un solo commit nuevo, con TODOS los archivos de esta publicación.
  var commitNuevo = ghFetch(base + "/git/commits", {
    token: token,
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ message: mensaje, tree: treeNuevo.sha, parents: [commitActualSha] }),
  });

  // 6) Mover la rama "main" a ese commit — ESTE es el único paso que
  // dispara un despliegue nuevo en Vercel (uno solo, sin importar cuántos
  // archivos se hayan tocado arriba).
  ghFetch(base + "/git/refs/heads/" + BRANCH, {
    token: token,
    method: "patch",
    contentType: "application/json",
    payload: JSON.stringify({ sha: commitNuevo.sha, force: false }),
  });

  return commitNuevo.sha;
}

// ===========================================================================
// CORRECCIONES MANUALES DE ACTIVIDAD/RUBRO (Atendidas AP) — Google Sheet
// ===========================================================================
// La hoja "Correcciones AP" (pestaña "correcciones") tiene los encabezados de
// COLS_CORRECCION en la fila 1. Cada fila = una corrección, clave = "reclamo".
//  - doGet ?path=correcciones           -> devuelve { rows: [...] }
//  - doPost accion:"correccion"          -> upsert por "reclamo" (con Lock)
// ---------------------------------------------------------------------------

function hojaCorrecciones_() {
  var ss = SpreadsheetApp.openById(SHEET_CORRECCIONES_ID);
  var hoja = ss.getSheetByName(SHEET_CORRECCIONES_HOJA);
  if (!hoja) {
    throw new Error("No existe la pestaña '" + SHEET_CORRECCIONES_HOJA + "' en la hoja de correcciones.");
  }
  return hoja;
}

function leerCorrecciones_() {
  var hoja = hojaCorrecciones_();
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { rows: [] };
  var valores = hoja.getRange(2, 1, ultimaFila - 1, COLS_CORRECCION.length).getValues();
  var rows = [];
  valores.forEach(function (fila) {
    if (String(fila[0]).trim() === "") return; // sin reclamo -> fila basura
    var o = {};
    COLS_CORRECCION.forEach(function (c, i) {
      var v = fila[i];
      // Las fechas de la hoja pueden venir como Date -> a texto ISO.
      o[c] = (v instanceof Date) ? v.toISOString() : String(v == null ? "" : v);
    });
    rows.push(o);
  });
  return { rows: rows };
}

function guardarCorrecciones_(items) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // espera hasta 30s a que otra escritura termine
  try {
    var hoja = hojaCorrecciones_();
    var ultimaFila = hoja.getLastRow();

    // Mapa clave(reclamo) -> nro de fila real en la hoja.
    var filaDe = {};
    if (ultimaFila >= 2) {
      var col = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
      for (var i = 0; i < col.length; i++) {
        var k = String(col[i][0]).trim();
        if (k) filaDe[k] = i + 2;
      }
    }

    // fecha_modificacion / fecha_validacion actuales, para no re-pisarlas si
    // no hace falta (cuando Pluz re-valida algo ya validado, se mantiene la
    // primera fecha de validación).
    var ahora = new Date().toISOString();
    var previas = {};
    if (ultimaFila >= 2) {
      var todo = hoja.getRange(2, 1, ultimaFila - 1, COLS_CORRECCION.length).getValues();
      todo.forEach(function (f) {
        var k = String(f[0]).trim();
        if (k) previas[k] = { validado: String(f[9]).trim().toUpperCase(), fechaValid: f[11] };
      });
    }

    var guardados = 0;
    items.forEach(function (it) {
      var clave = String(it.reclamo == null ? "" : it.reclamo).trim();
      if (!clave) return;

      var validado = !!it.validado;
      var fechaValid = "";
      if (validado) {
        var prev = previas[clave];
        fechaValid = (prev && prev.validado === "SI" && prev.fechaValid)
          ? prev.fechaValid   // ya estaba validado -> conservar la fecha original
          : ahora;
      }

      var fila = [
        clave,
        it.contratista || "",
        it.distrito || "",
        it.odm || "",
        it.observacion || "",
        it.actividad || "",
        it.rubro || "",
        it.modificado_por || "",
        ahora,
        validado ? "SI" : "NO",
        validado ? (it.validado_por || "") : "",
        fechaValid
      ];

      if (filaDe[clave]) {
        hoja.getRange(filaDe[clave], 1, 1, COLS_CORRECCION.length).setValues([fila]);
      } else {
        hoja.appendRow(fila);
        filaDe[clave] = hoja.getLastRow();
      }
      guardados++;
    });

    SpreadsheetApp.flush();
    return guardados;
  } finally {
    lock.releaseLock();
  }
}
