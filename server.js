// server.js
// Webhook de WhatsApp Cloud API. Soporta dos flujos:
//
//  1) "credencial <cedula>"  -> responde con el PDF de la credencial del profesor
//     (un archivo <cedula>.pdf dentro de DRIVE_FOLDER_ID)
//
//  2) "plantel <codigo_dea>" -> responde con todos los archivos del plantel
//     (informe .docx, materiales .xlsx, ficha .pptx dentro de una SUBCARPETA
//     llamada <codigo_dea> ubicada dentro de PLANTELES_FOLDER_ID)
//
// Requisitos: Node.js 18+
// Instalar dependencias:   npm install express axios dotenv csv-parse googleapis exifr
// Ejecutar:                node server.js

require('dotenv').config();
const express = require('express');
const fs = require('fs');
const os = require('os');
const path = require('path');
const axios = require('axios');
const { parse } = require('csv-parse/sync');
const { google } = require('googleapis');
const exifr = require('exifr');

const app = express();

// Middleware de diagnóstico: loguea TODA solicitud entrante (método + ruta),
// sin importar si existe una ruta definida para ella. Útil para detectar si
// Meta está llegando a una URL distinta a la esperada.
app.use((req, res, next) => {
  console.log(`>>> Solicitud entrante: ${req.method} ${req.originalUrl}`);
  next();
});

app.use(express.json());

// ---------- Configuración (poner estos valores en un archivo .env) ----------
const VERIFY_TOKEN = process.env.VERIFY_TOKEN;           // palabra secreta que tú inventas
const WHATSAPP_TOKEN = process.env.WHATSAPP_TOKEN;       // token permanente de Meta
const PHONE_NUMBER_ID = process.env.PHONE_NUMBER_ID;     // ID del número de WhatsApp (lo da Meta)
const DRIVE_FOLDER_ID = process.env.DRIVE_FOLDER_ID;     // ID de la carpeta "Credenciales Profesores" en Drive
const PLANTELES_FOLDER_ID = process.env.PLANTELES_FOLDER_ID; // ID de la carpeta que contiene una subcarpeta por cada plantel (nombrada con el código DEA)
const GOOGLE_SERVICE_ACCOUNT_KEY = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH; // ruta al JSON de la cuenta de servicio
const PROFESORES_CSV = path.join(__dirname, 'profesores.csv'); // columnas: cedula,telefono,nombre
const INGENIEROS_CSV = path.join(__dirname, 'ingenieros.csv'); // columnas: telefono,nombre (lista blanca para solicitar info de planteles)
const PLANTELES_CSV = path.join(__dirname, 'planteles.csv'); // columnas: codigo_dea,nombre
const SHEETS_COORDENADAS_ID = process.env.SHEETS_COORDENADAS_ID; // ID de la hoja "Coordenadas Planteles"
const NOMINA_RAC_CSV = path.join(__dirname, 'nomina_rac.csv'); // columnas: cedula,nombre,cargo,codigo_rac,codigo_dea,nombre_plantel,municipio,parroquia,horas_adm,fecha_ingreso,situacion,...
const PANEL_USER = process.env.PANEL_USER;         // usuario para entrar al panel interno
const PANEL_PASSWORD = process.env.PANEL_PASSWORD; // contraseña para entrar al panel interno

// Guarda, por número de teléfono, el código DEA que está esperando confirmación
// ("¿es este el plantel correcto?"). Se pierde si el servidor se reinicia,
// lo cual está bien: en ese caso el ingeniero simplemente vuelve a escribir.
const confirmacionesPendientes = new Map();

// Guarda, por número de teléfono, el código DEA al que corresponderá la
// PRÓXIMA foto (enviada como documento) que llegue de ese número. Se usa
// porque WhatsApp no siempre permite escribir una descripción al enviar un
// documento, así que el código se indica antes, como mensaje de texto aparte.
const fotoPendientePorTelefono = new Map();

const driveAuth = new google.auth.GoogleAuth({
  keyFile: GOOGLE_SERVICE_ACCOUNT_KEY,
  scopes: [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
  ],
});
const drive = google.drive({ version: 'v3', auth: driveAuth });
const sheets = google.sheets({ version: 'v4', auth: driveAuth });

// ---------- 1. Verificación del webhook (Meta la llama una sola vez al configurar) ----------
app.get('/webhook', (req, res) => {
  const mode = req.query['hub.mode'];
  const token = req.query['hub.verify_token'];
  const challenge = req.query['hub.challenge'];

  console.log('GET /webhook recibido. mode:', mode, '| token recibido:', token, '| token esperado:', VERIFY_TOKEN, '| coinciden:', token === VERIFY_TOKEN);

  if (mode === 'subscribe' && token === VERIFY_TOKEN) {
    console.log('Verificación de webhook exitosa.');
    return res.status(200).send(challenge);
  }
  console.log('Verificación de webhook FALLIDA.');
  return res.sendStatus(403);
});

// ---------- 2. Recepción de mensajes entrantes ----------
app.post('/webhook', async (req, res) => {
  // Responder rápido a Meta para que no reintente la entrega
  res.sendStatus(200);
  console.log('POST /webhook recibido:', JSON.stringify(req.body));

  try {
    const entry = req.body.entry?.[0];
    const change = entry?.changes?.[0];
    const message = change?.value?.messages?.[0];

    if (!message) {
      console.log('No hay mensaje en el payload, se ignora.');
      return;
    }

    if (message.type === 'document') {
      await manejarDocumentoRecibido(message);
      return;
    }

    if (message.type === 'image') {
      // Foto enviada como "foto" normal (comprimida): WhatsApp le borra el GPS,
      // así que no sirve para registrar coordenadas. Se avisa cómo enviarla bien.
      console.log(`Imagen (foto normal) recibida de ${message.from}, se avisa que debe enviarse como documento.`);
      await enviarTexto(
        message.from,
        'Para registrar la ubicación de un plantel: primero escribe "foto <código DEA>", y luego envía la foto como *Documento* (📎 → Documento), no como foto normal — así no se pierde la ubicación GPS.'
      );
      return;
    }

    if (message.type !== 'text') {
      console.log(`Mensaje de tipo "${message.type}" recibido, se ignora (solo se procesan texto, imagen y documento).`);
      return;
    }

    const from = message.from; // número de WhatsApp de quien escribe (formato: 58414XXXXXXX)
    const textoOriginal = message.text.body.trim();
    const texto = textoOriginal.toLowerCase();
    console.log(`Mensaje de ${from}: "${texto}"`);

    // Si este número tiene una confirmación de plantel pendiente, este mensaje
    // se interpreta como respuesta a esa pregunta (confirmar o cancelar),
    // en vez de como un nuevo comando.
    if (confirmacionesPendientes.has(from)) {
      const pendiente = confirmacionesPendientes.get(from);
      confirmacionesPendientes.delete(from);

      if (/^(s[ií]|si|confirmar|correcto|ok)$/i.test(texto)) {
        await enviarArchivosPlantel(from, pendiente.codigoDea, pendiente.nombreIngeniero);
        return;
      }
      // Cualquier otra respuesta cancela la confirmación pendiente y se
      // procesa como un mensaje nuevo (por ejemplo, otro "plantel <código>").
    }

    if (texto.startsWith('credencial')) {
      // La cédula es solo numérica
      const cedula = texto.replace(/^credencial/, '').trim().match(/\d{6,9}/)?.[0];
      if (!cedula) {
        await enviarTexto(from, 'No entendí tu solicitud. Escribe: credencial <cédula>, plantel <código DEA> o foto <código DEA>.');
        return;
      }
      await manejarCredencial(from, cedula);
    } else if (texto.startsWith('plantel')) {
      // El código DEA puede incluir letras (ej. OD02321608), así que se toma la cadena completa
      const codigoDea = textoOriginal.replace(/^plantel/i, '').trim().toUpperCase();
      if (!codigoDea) {
        await enviarTexto(from, 'No entendí tu solicitud. Escribe: credencial <cédula>, plantel <código DEA> o foto <código DEA>.');
        return;
      }
      await manejarPlantel(from, codigoDea);
    } else if (texto.startsWith('foto')) {
      await manejarComandoFoto(from, textoOriginal.replace(/^foto/i, '').trim().toUpperCase());
    } else {
      await enviarTexto(from, 'No entendí tu solicitud. Escribe: credencial <cédula>, plantel <código DEA> o foto <código DEA>.');
    }
  } catch (err) {
    console.error('Error procesando mensaje:', err.message);
  }
});

// ---------- Flujo 1: credencial de profesor (un solo PDF) ----------
async function manejarCredencial(from, cedula) {
  const { autorizado, nombre } = validarProfesor(cedula, from);
  if (!autorizado) {
    await enviarTexto(from, 'No pudimos verificar que este número de WhatsApp corresponda a esa cédula. Contacta a la oficina.');
    return;
  }
  const primerNombre = obtenerPrimerNombre(nombre);

  const rutaArchivo = await descargarCredencialDesdeDrive(cedula);
  if (!rutaArchivo) {
    const saludo = primerNombre ? `${primerNombre}, no` : 'No';
    await enviarTexto(from, `${saludo} encontramos una credencial digitalizada para la cédula ${cedula}. Contacta a la oficina.`);
    return;
  }

  if (primerNombre) {
    await enviarTexto(from, `${primerNombre}, aquí tienes tu credencial:`);
  }
  await enviarDocumento(from, rutaArchivo, `credencial_${cedula}.pdf`);
  fs.unlinkSync(rutaArchivo);
}

// ---------- Flujo 2: plantel (varios archivos: informe, materiales, ficha) ----------
// Primero valida al ingeniero y busca el nombre del plantel en planteles.csv,
// y pide confirmación antes de enviar los archivos (para evitar errores de
// código DEA mal escrito).
async function manejarPlantel(from, codigoDea) {
  const { autorizado, nombre } = validarIngeniero(from);
  if (!autorizado) {
    await enviarTexto(from, 'No pudimos verificar que este número de WhatsApp esté autorizado para solicitar información de planteles. Contacta a la oficina.');
    return;
  }
  const primerNombre = obtenerPrimerNombre(nombre);

  const plantel = buscarPlantel(codigoDea);
  if (!plantel) {
    const saludo = primerNombre ? `${primerNombre}, no` : 'No';
    await enviarTexto(from, `${saludo} encontramos un plantel con código ${codigoDea}. Verifica el código y envía la consulta nuevamente.`);
    return;
  }

  confirmacionesPendientes.set(from, { codigoDea, nombreIngeniero: primerNombre });
  const pregunta = primerNombre ? `${primerNombre}, ¿es este el plantel correcto?` : '¿Es este el plantel correcto?';
  await enviarTexto(
    from,
    `${pregunta}\n\n*${plantel.nombre}*\nMunicipio: ${plantel.municipio}\nParroquia: ${plantel.parroquia}\nCódigo DEA: ${codigoDea}\n\nResponde *SI* para confirmar y recibir los documentos, o envía el código nuevamente si no es correcto.`
  );
}

// Envía los archivos del plantel una vez que el ingeniero ya confirmó.
async function enviarArchivosPlantel(from, codigoDea, nombreIngeniero) {
  const archivos = await descargarArchivosDePlantel(codigoDea);
  if (archivos.length === 0) {
    const saludo = nombreIngeniero ? `${nombreIngeniero}, no` : 'No';
    await enviarTexto(from, `${saludo} encontramos documentos para el plantel con código ${codigoDea}. Verifica el código o contacta a la oficina.`);
    return;
  }

  const saludo = nombreIngeniero ? `${nombreIngeniero}, encontramos` : 'Encontramos';
  await enviarTexto(from, `${saludo} ${archivos.length} documento(s) para el plantel ${codigoDea}. Te los envío a continuación:`);

  // Se envían uno por uno; WhatsApp los muestra como mensajes seguidos.
  for (const archivo of archivos) {
    await enviarDocumento(from, archivo.rutaLocal, archivo.nombre);
    fs.unlinkSync(archivo.rutaLocal);
  }
}

// ---------- Flujo 3: recepción de foto del plantel con coordenadas GPS ----------
// Paso 1: el ingeniero escribe "foto <código DEA>" (o varios códigos separados
// por coma, si el edificio alberga más de un plantel: "foto COD1, COD2") — se
// validan y se marcan como "pendientes" para el próximo documento que llegue
// de ese número.
// Paso 2: el ingeniero envía la foto COMO DOCUMENTO (así WhatsApp no le borra
// los metadatos EXIF, a diferencia de cuando se envía como foto normal). No
// hace falta escribir nada más en ese segundo mensaje. Esa misma foto y esas
// mismas coordenadas se guardan para TODOS los códigos que quedaron pendientes.
async function manejarComandoFoto(from, textoCodigos) {
  const { autorizado, nombre } = validarIngeniero(from);
  if (!autorizado) {
    await enviarTexto(from, 'No pudimos verificar que este número de WhatsApp esté autorizado para reportar fotos de planteles. Contacta a la oficina.');
    return;
  }
  const primerNombre = obtenerPrimerNombre(nombre);

  const codigos = extraerListaCodigos(textoCodigos);
  if (codigos.length === 0) {
    await enviarTexto(
      from,
      'Escribe el código DEA después de "foto", por ejemplo: foto OD02321608\n\nSi el edificio alberga más de un plantel, sepáralos por coma: foto OD02321608, S0937D1609'
    );
    return;
  }

  const planteles = [];
  const noEncontrados = [];
  for (const codigo of codigos) {
    const plantel = buscarPlantel(codigo);
    if (plantel) {
      planteles.push({ codigo, plantel });
    } else {
      noEncontrados.push(codigo);
    }
  }

  if (noEncontrados.length > 0) {
    const saludo = primerNombre ? `${primerNombre}, no` : 'No';
    await enviarTexto(from, `${saludo} encontramos plantel(es) con código: ${noEncontrados.join(', ')}. Verifica el/los código(s) e inténtalo de nuevo.`);
    return;
  }

  fotoPendientePorTelefono.set(from, planteles.map(p => p.codigo));
  const listaNombres = planteles.map(p => `*${p.plantel.nombre}* (${p.codigo})`).join('\n');
  const saludo = primerNombre ? `${primerNombre}, envía` : 'Envía';
  const plural = planteles.length > 1 ? 'los planteles' : 'el plantel';
  await enviarTexto(
    from,
    `${saludo} ahora la foto de ${plural}:\n\n${listaNombres}\n\ncomo *Documento* (📎 → Documento), con la ubicación GPS activada en la cámara.`
  );
}

// Convierte "OD1, OD2  od3" en ['OD1', 'OD2', 'OD3'] — separa por comas y/o
// espacios, quita vacíos y normaliza a mayúsculas.
function extraerListaCodigos(texto) {
  return (texto || '')
    .split(/[,\s]+/)
    .map(c => c.trim().toUpperCase())
    .filter(c => c.length > 0);
}

// A partir de un grupo de códigos DEA que comparten un mismo edificio, elige
// el código de infraestructura: el menor alfabéticamente del grupo. Así, sin
// necesidad de mantener una tabla aparte, todos los planteles de un mismo
// edificio quedan identificados con el mismo valor en esa columna.
function calcularCodigoInfraestructura(codigos) {
  return [...codigos].sort()[0];
}

async function manejarDocumentoRecibido(message) {
  const from = message.from;
  const documento = message.document;
  const caption = (documento?.caption || '').trim();
  console.log(`Documento recibido de ${from}. mime_type: ${documento?.mime_type}, caption: "${caption}"`);

  const { autorizado, nombre } = validarIngeniero(from);
  if (!autorizado) {
    await enviarTexto(from, 'No pudimos verificar que este número de WhatsApp esté autorizado para reportar fotos de planteles. Contacta a la oficina.');
    return;
  }
  const primerNombre = obtenerPrimerNombre(nombre);

  if (!documento || !documento.mime_type?.startsWith('image/')) {
    await enviarTexto(from, 'Para registrar la ubicación de un plantel, primero escribe: foto <código DEA>, y luego envía la foto como *documento*.');
    return;
  }

  // Prioridad: los códigos que quedaron pendientes de "foto <código(s)>". Si
  // por algún motivo no hay ninguno pendiente, se usa el caption como
  // respaldo (por si el cliente de WhatsApp sí permite escribir descripción).
  const codigosPendientes = fotoPendientePorTelefono.get(from) || extraerListaCodigos(caption);
  fotoPendientePorTelefono.delete(from);

  if (codigosPendientes.length === 0) {
    await enviarTexto(from, 'No sabemos a qué plantel corresponde esta foto. Primero escribe: foto <código DEA>, y luego envía la foto como documento.');
    return;
  }

  const planteles = [];
  const noEncontrados = [];
  for (const codigo of codigosPendientes) {
    const plantel = buscarPlantel(codigo);
    if (plantel) {
      planteles.push({ codigo, plantel });
    } else {
      noEncontrados.push(codigo);
    }
  }

  if (planteles.length === 0) {
    const saludo = primerNombre ? `${primerNombre}, no` : 'No';
    await enviarTexto(from, `${saludo} encontramos ningún plantel válido entre esos códigos. Verifica e inténtalo de nuevo.`);
    return;
  }

  let coordenadas;
  try {
    const bufferImagen = await descargarMediaDeWhatsApp(documento.id);
    coordenadas = await exifr.gps(bufferImagen);
  } catch (err) {
    console.error('Error descargando o leyendo la foto:', err.message);
  }

  if (!coordenadas || coordenadas.latitude == null) {
    const saludo = primerNombre ? `${primerNombre}, no` : 'No';
    const codigosTexto = planteles.map(p => p.codigo).join(', ');
    await enviarTexto(
      from,
      `${saludo} pudimos leer la ubicación GPS de esa foto. Verifica que la ubicación esté activada en la cámara del teléfono, que la envíes como *documento* (no como foto normal), e inténtalo de nuevo (escribe "foto ${codigosTexto}" de nuevo para reintentar).`
    );
    return;
  }

  // Se guarda la misma ubicación para todos los planteles que comparten esta foto/edificio.
  // El código de infraestructura del grupo es el menor alfabéticamente entre
  // los códigos enviados juntos; "esGrupoExplicito" indica si en este envío
  // se listó más de un código (agrupación activa) o solo uno.
  const esGrupoExplicito = planteles.length > 1;
  const codigoInfraestructura = calcularCodigoInfraestructura(planteles.map(p => p.codigo));
  for (const { codigo } of planteles) {
    await guardarCoordenadasEnSheet(codigo, codigoInfraestructura, esGrupoExplicito, coordenadas.latitude, coordenadas.longitude, primerNombre || from);
  }

  const enlaceMaps = `https://www.google.com/maps?q=${coordenadas.latitude.toFixed(6)},${coordenadas.longitude.toFixed(6)}`;
  const saludo = primerNombre ? `${primerNombre}, guardamos` : 'Guardamos';
  const listaNombres = planteles.map(p => `*${p.plantel.nombre}* (${p.codigo})`).join('\n');
  let mensaje = `${saludo} la ubicación de:\n\n${listaNombres}\n\n📍 ${enlaceMaps}`;
  if (noEncontrados.length > 0) {
    mensaje += `\n\n(No se encontró: ${noEncontrados.join(', ')})`;
  }
  await enviarTexto(from, mensaje);
}

// Descarga un archivo multimedia de WhatsApp a partir de su media ID.
// WhatsApp requiere dos pasos: primero pedir la URL temporal de descarga,
// luego descargar el contenido de esa URL (ambos pasos llevan el token).
async function descargarMediaDeWhatsApp(mediaId) {
  const metaRes = await axios.get(`https://graph.facebook.com/v20.0/${mediaId}`, {
    headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}` },
  });
  const urlDescarga = metaRes.data.url;

  const archivoRes = await axios.get(urlDescarga, {
    headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}` },
    responseType: 'arraybuffer',
  });
  return Buffer.from(archivoRes.data);
}

// Guarda (o actualiza) la fila de coordenadas de un plantel en la hoja de
// Google Sheets "Coordenadas Planteles". Columnas: codigo_dea,
// codigo_infraestructura, latitud, longitud, fecha_captura, ingeniero.
//
// codigoInfraestructuraPropuesto: el código "candidato" a usar como ID del
// edificio (normalmente el menor alfabéticamente del grupo enviado ahora).
// esGrupoExplicito: true si en ESTE envío se listaron varios códigos juntos
// (el ingeniero está confirmando activamente que comparten edificio). Si es
// false (se envió un solo código) y esa fila YA tenía un código de
// infraestructura asignado antes, se conserva ese valor en vez de
// sobrescribirlo — así un reenvío individual no "desagrupa" el edificio.
async function guardarCoordenadasEnSheet(codigoDea, codigoInfraestructuraPropuesto, esGrupoExplicito, latitud, longitud, ingeniero) {
  const respuesta = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEETS_COORDENADAS_ID,
    range: 'A:B',
  });
  const filas = respuesta.data.values || [];
  const indiceFila = filas.findIndex(f => (f[0] || '').trim().toUpperCase() === codigoDea);

  let codigoInfraestructuraFinal = codigoInfraestructuraPropuesto;
  if (indiceFila !== -1 && !esGrupoExplicito) {
    const infraExistente = (filas[indiceFila][1] || '').trim();
    if (infraExistente) {
      codigoInfraestructuraFinal = infraExistente;
    }
  }

  const fechaCaptura = new Date().toISOString();
  const nuevaFila = [codigoDea, codigoInfraestructuraFinal, latitud, longitud, fechaCaptura, ingeniero];

  if (indiceFila === -1) {
    // No existe todavía: se agrega una fila nueva al final
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEETS_COORDENADAS_ID,
      range: 'A:F',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [nuevaFila] },
    });
  } else {
    // Ya existe una fila para este plantel: se actualiza esa misma fila
    // (indiceFila es base 0 según el array; las filas de Sheets son base 1)
    const numeroFila = indiceFila + 1;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEETS_COORDENADAS_ID,
      range: `A${numeroFila}:F${numeroFila}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [nuevaFila] },
    });
  }
}

// ---------- Panel interno: búsqueda en nómina + RAC ----------
// Página protegida (usuario/contraseña) para que un encargado de la oficina
// busque una cédula y vea sus datos de nómina/RAC, como primer paso hacia la
// generación automática de credenciales.

// Autenticación básica HTTP: el navegador pide usuario/contraseña con un
// cuadro nativo. Solo protege las rutas /panel y /api/*.
function requiereAutenticacion(req, res, next) {
  const encabezado = req.headers.authorization;
  if (!encabezado || !encabezado.startsWith('Basic ')) {
    res.set('WWW-Authenticate', 'Basic realm="Panel interno"');
    return res.status(401).send('Autenticación requerida.');
  }
  const [usuario, clave] = Buffer.from(encabezado.slice(6), 'base64').toString().split(':');
  if (usuario === PANEL_USER && clave === PANEL_PASSWORD) {
    return next();
  }
  res.set('WWW-Authenticate', 'Basic realm="Panel interno"');
  return res.status(401).send('Usuario o contraseña incorrectos.');
}

// Cachea el contenido de nomina_rac.csv en memoria la primera vez que se pide,
// para no releer/parsear un archivo de ~8MB en cada búsqueda. Se recarga solo
// si el servidor se reinicia (o redespliega).
let cacheNominaRac = null;
function cargarNominaRac() {
  if (cacheNominaRac) return cacheNominaRac;
  if (!fs.existsSync(NOMINA_RAC_CSV)) {
    cacheNominaRac = [];
    return cacheNominaRac;
  }
  const contenido = fs.readFileSync(NOMINA_RAC_CSV, 'utf8');
  cacheNominaRac = parse(contenido, { columns: true, skip_empty_lines: true });
  console.log(`nomina_rac.csv cargado en memoria: ${cacheNominaRac.length} registros`);
  return cacheNominaRac;
}

// Busca todas las filas que correspondan a una cédula (una persona puede
// tener varias filas: distintas materias, secciones, o asignaciones).
function buscarEmpleadoPorCedula(cedula) {
  const registros = cargarNominaRac();
  const cedulaLimpia = cedula.trim();
  return registros.filter(r => (r.cedula || '').trim() === cedulaLimpia);
}

app.get('/api/empleado/:cedula', requiereAutenticacion, (req, res) => {
  const filas = buscarEmpleadoPorCedula(req.params.cedula);
  res.json({ encontrado: filas.length > 0, registros: filas });
});

app.get('/panel', requiereAutenticacion, (req, res) => {
  res.type('html').send(`<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Panel interno — Búsqueda de personal</title>
<style>
  body { font-family: system-ui, sans-serif; max-width: 720px; margin: 40px auto; padding: 0 20px; background: #f6f5f2; color: #1f1f1f; }
  h1 { font-size: 1.4rem; }
  .buscador { display: flex; gap: 10px; margin: 20px 0; }
  input[type=text] { flex: 1; padding: 10px 14px; font-size: 1rem; border: 1px solid #ccc; border-radius: 6px; }
  button { padding: 10px 18px; font-size: 1rem; background: #1f1f1f; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
  button:hover { background: #3a3a3a; }
  .registro { background: #fff; border: 1px solid #ddd; border-radius: 8px; padding: 16px 20px; margin-bottom: 14px; }
  .registro .nombre { font-weight: 700; font-size: 1.05rem; }
  .campo { margin: 4px 0; font-size: 0.92rem; }
  .campo b { display: inline-block; min-width: 150px; color: #555; }
  .situacion-activo { color: #0a7a2e; font-weight: 700; }
  .situacion-otro { color: #b3261e; font-weight: 700; }
  .mensaje { color: #555; margin-top: 20px; }
</style>
</head>
<body>
  <h1>Búsqueda de personal (nómina + RAC)</h1>
  <div class="buscador">
    <input type="text" id="cedula" placeholder="Escribe una cédula, ej. 9293431" />
    <button onclick="buscar()">Buscar</button>
  </div>
  <div id="resultados"></div>

<script>
async function buscar() {
  const cedula = document.getElementById('cedula').value.trim();
  const resultados = document.getElementById('resultados');
  if (!cedula) { resultados.innerHTML = '<p class="mensaje">Escribe una cédula.</p>'; return; }
  resultados.innerHTML = '<p class="mensaje">Buscando…</p>';
  try {
    const res = await fetch('/api/empleado/' + encodeURIComponent(cedula));
    const data = await res.json();
    if (!data.encontrado) {
      resultados.innerHTML = '<p class="mensaje">No se encontró ningún registro con esa cédula.</p>';
      return;
    }
    resultados.innerHTML = data.registros.map(r => {
      const activo = (r.situacion || '').toUpperCase() === 'ACTIVO';
      return \`<div class="registro">
        <div class="nombre">\${r.nombre || ''}</div>
        <div class="campo"><b>Cédula:</b> \${r.cedula || ''}</div>
        <div class="campo"><b>Situación:</b> <span class="\${activo ? 'situacion-activo' : 'situacion-otro'}">\${r.situacion || '—'}</span></div>
        <div class="campo"><b>Cargo:</b> \${r.cargo || '—'} (código \${r.codigo_rac || '—'})</div>
        <div class="campo"><b>Tipo de personal:</b> \${r.tipo_personal || '—'}</div>
        <div class="campo"><b>Plantel:</b> \${r.nombre_plantel || '—'} (código DEA \${r.codigo_dea || '—'})</div>
        <div class="campo"><b>Municipio / Parroquia:</b> \${r.municipio || '—'} / \${r.parroquia || '—'}</div>
        <div class="campo"><b>Horas administrativas:</b> \${r.horas_adm || '—'}</div>
        <div class="campo"><b>Fecha de ingreso:</b> \${r.fecha_ingreso || '—'}</div>
      </div>\`;
    }).join('');
  } catch (err) {
    resultados.innerHTML = '<p class="mensaje">Ocurrió un error al buscar. Intenta de nuevo.</p>';
  }
}
document.getElementById('cedula').addEventListener('keydown', e => { if (e.key === 'Enter') buscar(); });
</script>
</body>
</html>`);
});

// ---------- Utilidades ----------

// Busca el plantel completo (nombre, estado, municipio, parroquia) en
// planteles.csv a partir del código DEA. Devuelve null si no lo encuentra.
function buscarPlantel(codigoDea) {
  if (!fs.existsSync(PLANTELES_CSV)) return null;
  const contenido = fs.readFileSync(PLANTELES_CSV, 'utf8');
  const registros = parse(contenido, { columns: true, skip_empty_lines: true });
  const registro = registros.find(r => r.codigo_dea.trim().toUpperCase() === codigoDea.toUpperCase());
  if (!registro) return null;
  return {
    nombre: registro.nombre.trim(),
    estado: registro.estado.trim(),
    municipio: registro.municipio.trim(),
    parroquia: registro.parroquia.trim(),
  };
}


// Revisa en profesores.csv que el (cedula, telefono) coincidan, para evitar
// que alguien pida la credencial de otra persona adivinando la cédula.
// Devuelve { autorizado, nombre } — nombre es null si no se encontró registro.
function validarProfesor(cedula, telefonoEntrante) {
  if (!fs.existsSync(PROFESORES_CSV)) return { autorizado: true, nombre: null }; // si aún no tienes el CSV, deja pasar (ajusta luego)
  const contenido = fs.readFileSync(PROFESORES_CSV, 'utf8');
  const registros = parse(contenido, { columns: true, skip_empty_lines: true });
  const registro = registros.find(r => r.cedula.trim() === cedula && telefonoEntrante.endsWith(r.telefono.trim()));
  return { autorizado: !!registro, nombre: registro ? registro.nombre.trim() : null };
}

// Revisa en ingenieros.csv que el número de teléfono esté en la lista blanca,
// para que solo ingenieros autorizados puedan pedir información de planteles.
// Devuelve { autorizado, nombre } — nombre es null si no se encontró registro.
function validarIngeniero(telefonoEntrante) {
  if (!fs.existsSync(INGENIEROS_CSV)) return { autorizado: true, nombre: null }; // si aún no tienes el CSV, deja pasar (ajusta luego)
  const contenido = fs.readFileSync(INGENIEROS_CSV, 'utf8');
  const registros = parse(contenido, { columns: true, skip_empty_lines: true });
  const registro = registros.find(r => telefonoEntrante.endsWith(r.telefono.trim()));
  return { autorizado: !!registro, nombre: registro ? registro.nombre.trim() : null };
}

// A partir de un nombre completo (ej. "JENNIFER VILLAHERMOSA" o "Victor Molina"),
// devuelve solo el primer nombre, con formato "Jennifer" (primera letra mayúscula).
function obtenerPrimerNombre(nombreCompleto) {
  if (!nombreCompleto) return null;
  const primera = nombreCompleto.trim().split(/\s+/)[0];
  return primera.charAt(0).toUpperCase() + primera.slice(1).toLowerCase();
}

// Busca <cedula>.pdf dentro de la carpeta de Drive y lo descarga a un archivo
// temporal local. Devuelve la ruta local, o null si no existe.
async function descargarCredencialDesdeDrive(cedula) {
  const res = await drive.files.list({
    q: `'${DRIVE_FOLDER_ID}' in parents and name = '${cedula}.pdf' and trashed = false`,
    fields: 'files(id, name)',
  });
  const archivo = res.data.files?.[0];
  if (!archivo) return null;

  const rutaTemporal = path.join(os.tmpdir(), `${cedula}.pdf`);
  const destino = fs.createWriteStream(rutaTemporal);
  const descarga = await drive.files.get(
    { fileId: archivo.id, alt: 'media' },
    { responseType: 'stream' }
  );

  await new Promise((resolve, reject) => {
    descarga.data.pipe(destino).on('finish', resolve).on('error', reject);
  });

  return rutaTemporal;
}

// Busca la subcarpeta <codigoDea> dentro de PLANTELES_FOLDER_ID, y descarga
// TODOS los archivos que contenga (informe, materiales, ficha, etc.) a rutas
// temporales locales. Devuelve un arreglo [{ rutaLocal, nombre }].
async function descargarArchivosDePlantel(codigoDea) {
  // 1) Buscar la subcarpeta con el nombre del código DEA
  const resCarpeta = await drive.files.list({
    q: `'${PLANTELES_FOLDER_ID}' in parents and name = '${codigoDea}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
    fields: 'files(id, name)',
  });
  const carpeta = resCarpeta.data.files?.[0];
  if (!carpeta) return [];

  // 2) Listar todos los archivos dentro de esa subcarpeta
  const resArchivos = await drive.files.list({
    q: `'${carpeta.id}' in parents and trashed = false`,
    fields: 'files(id, name)',
  });
  const listaArchivos = resArchivos.data.files || [];

  // 3) Descargar cada uno a una ruta temporal
  const resultado = [];
  for (const archivo of listaArchivos) {
    const rutaTemporal = path.join(os.tmpdir(), `${codigoDea}_${archivo.name}`);
    const destino = fs.createWriteStream(rutaTemporal);
    const descarga = await drive.files.get(
      { fileId: archivo.id, alt: 'media' },
      { responseType: 'stream' }
    );
    await new Promise((resolve, reject) => {
      descarga.data.pipe(destino).on('finish', resolve).on('error', reject);
    });
    resultado.push({ rutaLocal: rutaTemporal, nombre: archivo.name });
  }

  return resultado;
}

async function enviarTexto(to, body) {
  await axios.post(
    `https://graph.facebook.com/v20.0/${PHONE_NUMBER_ID}/messages`,
    { messaging_product: 'whatsapp', to, type: 'text', text: { body } },
    { headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}` } }
  );
}

// Tipos MIME que WhatsApp acepta para documentos, según la extensión del archivo
const MIME_POR_EXTENSION = {
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

async function enviarDocumento(to, rutaArchivo, nombreArchivo) {
  const extension = path.extname(nombreArchivo).toLowerCase();
  const tipoMime = MIME_POR_EXTENSION[extension] || 'application/octet-stream';

  // 1) Subir el archivo a los servidores de Meta
  const FormData = require('form-data');
  const form = new FormData();
  form.append('file', fs.createReadStream(rutaArchivo));
  form.append('type', tipoMime);
  form.append('messaging_product', 'whatsapp');

  const uploadRes = await axios.post(
    `https://graph.facebook.com/v20.0/${PHONE_NUMBER_ID}/media`,
    form,
    { headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}`, ...form.getHeaders() } }
  );
  const mediaId = uploadRes.data.id;

  // 2) Enviar el documento usando el media_id
  await axios.post(
    `https://graph.facebook.com/v20.0/${PHONE_NUMBER_ID}/messages`,
    {
      messaging_product: 'whatsapp',
      to,
      type: 'document',
      document: { id: mediaId, filename: nombreArchivo },
    },
    { headers: { Authorization: `Bearer ${WHATSAPP_TOKEN}` } }
  );
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor escuchando en puerto ${PORT}`));
