require('dotenv').config();
const crypto = require('crypto');
const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const Groq = require('groq-sdk');

const app = express();
const PORT = process.env.PORT || 3000;
const CHAT_MODEL = process.env.GROQ_CHAT_MODEL || 'openai/gpt-oss-120b';
// Modelo para buscar empresas. groq/compound trae búsqueda web integrada.
const SEARCH_MODEL = process.env.GROQ_SEARCH_MODEL || 'groq/compound';
const groq = process.env.GROQ_API_KEY ? new Groq({ apiKey: process.env.GROQ_API_KEY }) : null;
const files = new Map(); // id -> buffer (en memoria)

app.use(cors());
app.use(express.json());

const AGENT_PROMPT = `Eres co-finder, un agente amable que ayuda a encontrar empresas. Hablas en español, con mensajes cortos y naturales.
Necesitas tres datos: 1) giro de la empresa, 2) ciudad, 3) tamaño. Solo hay dos bandas válidas: "250-1000" empleados o "más de 1000" empleados.
Conversa libremente: si el usuario da varios datos juntos, no los repitas; pide solo lo que falte, de a uno. Si el tamaño es vago o está fuera de esas bandas (menos de 250), explica que solo buscas empresas de 250 en adelante y pide que elija una de las dos.
Cuando tengas los tres datos, responde ÚNICAMENTE con este formato exacto y nada más:
<READY>{"giro":"...","ciudad":"...","tamano":"..."}</READY>
Si el usuario pregunta algo fuera de tu función, responde breve y regresa a los datos que faltan.`;

async function grok(messages, temperature = 0.4, model = CHAT_MODEL) {
  if (!groq) throw new Error('Falta GROQ_API_KEY en backend/.env');
  const res = await groq.chat.completions.create({ model, temperature, messages });
  return res.choices?.[0]?.message?.content || '';
}

const FIELDS = '{"nombre":"","giro":"","ciudad":"","empleados":"","sitio_web":"","telefono":"","correo":"","direccion":""}';

function extractList(text) {
  const clean = (text || '').replace(/```json|```/g, '');
  // 1) arreglo directo
  const s = clean.indexOf('[');
  const e = clean.lastIndexOf(']');
  if (s !== -1 && e > s) {
    try {
      const arr = JSON.parse(clean.slice(s, e + 1));
      if (Array.isArray(arr) && arr.length) return arr;
    } catch (_) {}
  }
  // 2) objeto que contiene un arreglo
  const os = clean.indexOf('{');
  const oe = clean.lastIndexOf('}');
  if (os !== -1 && oe > os) {
    try {
      const obj = JSON.parse(clean.slice(os, oe + 1));
      const arr = Object.values(obj).find(Array.isArray);
      if (arr && arr.length) return arr;
    } catch (_) {}
  }
  return null;
}

async function findCompanies({ giro, ciudad, tamano }) {
  const tools = SEARCH_MODEL.startsWith('openai/gpt-oss') ? [{ type: 'browser_search' }] : undefined;
  const raw = await grok([
    { role: 'system', content: 'Eres un investigador de empresas. Usa la búsqueda web y no inventes empresas.' },
    {
      role: 'user',
      content: `Busca en internet empresas reales del giro "${giro}" ubicadas en "${ciudad}" con entre ${tamano} empleados.
Encuentra entre 10 y 15. Para cada una da: nombre, giro, ciudad, empleados (aprox.), sitio web, teléfono, correo y dirección.
Si no encuentras un dato, déjalo vacío. Al final entrega la lista como arreglo JSON con elementos de esta forma: ${FIELDS}`
    }
  ], 0.2, SEARCH_MODEL, tools);

  console.log('--- Respuesta cruda de búsqueda ---\n', raw.slice(0, 1500));
  if (!raw.trim()) throw new Error('La búsqueda no devolvió contenido. Puede que browser_search no esté disponible en tu plan.');

  let list = extractList(raw);
  if (!list) {
    // Segunda pasada: convertir lo encontrado a JSON estricto
    const fixed = await groq.chat.completions.create({
      model: CHAT_MODEL,
      temperature: 0,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: `Convierte el texto en JSON con la forma {"empresas":[${FIELDS}]}. Usa solo empresas mencionadas en el texto. Responde solo JSON.` },
        { role: 'user', content: raw }
      ]
    });
    list = extractList(fixed.choices?.[0]?.message?.content);
  }
  if (!list) throw new Error('No pude extraer empresas de la respuesta. Revisa la terminal del backend y prueba con otro giro o ciudad.');
  return list;
}

async function buildExcel(companies) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Empresas');
  ws.columns = [
    { header: 'Nombre', key: 'nombre', width: 32 },
    { header: 'Giro', key: 'giro', width: 24 },
    { header: 'Ciudad', key: 'ciudad', width: 18 },
    { header: 'Empleados', key: 'empleados', width: 14 },
    { header: 'Sitio web', key: 'sitio_web', width: 32 },
    { header: 'Teléfono', key: 'telefono', width: 18 },
    { header: 'Correo', key: 'correo', width: 28 },
    { header: 'Dirección', key: 'direccion', width: 40 }
  ];
  companies.forEach((c) => ws.addRow(c));
  const head = ws.getRow(1);
  head.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  head.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF12263F' } };
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// Conversación: recibe el historial y responde como agente
app.post('/api/chat', async (req, res) => {
  const history = Array.isArray(req.body?.messages) ? req.body.messages : [];
  try {
    const reply = await grok([{ role: 'system', content: AGENT_PROMPT }, ...history]);
    const match = reply.match(/<READY>([\s\S]*?)<\/READY>/);
    if (!match) return res.json({ reply });

    const params = JSON.parse(match[1]);
    const companies = await findCompanies(params);
    const id = crypto.randomUUID();
    files.set(id, await buildExcel(companies));
    res.json({
      reply: `Listo. Encontré ${companies.length} empresas de ${params.giro} en ${params.ciudad} (${params.tamano} empleados).`,
      fileUrl: `/api/files/${id}`,
      total: companies.length
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/files/:id', (req, res) => {
  const buf = files.get(req.params.id);
  if (!buf) return res.status(404).json({ error: 'Archivo no encontrado' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="empresas.xlsx"');
  res.send(buf);
});

app.listen(PORT, () => console.log(`co-finder backend en http://localhost:${PORT}`));
