# Garrido Fontal SLU — Xerador de Documentos

App web para xerar orzamentos e facturas en formato Word (.docx) con logo, cores e formato profesional.

---

## Estrutura

```
garrido-fontal-api/
├── server.js              ← Servidor Express (API)
├── package.json
├── logo.png               ← Logo Garrido Fontal + Velux Pro
├── builders/
│   ├── helpers.js         ← Elementos compartidos (cabeceira, totais, etc.)
│   ├── presuposto.js      ← Constructor de orzamentos
│   └── factura.js         ← Constructor de facturas
└── public/
    └── index.html         ← App web (formulario)
```

---

## Executar en local

```bash
npm install
npm start
# Abrir http://localhost:3000
```

---

## Desplegar en Railway (gratuito)

1. Crear conta en https://railway.app (con GitHub)
2. **New Project → Deploy from GitHub repo**
3. Sube esta carpeta a un repositorio de GitHub (público ou privado)
4. Railway detecta o `package.json` e despliega automaticamente
5. En **Settings → Domains** xera un dominio gratuíto
6. Garda a URL (exemplo: `https://garrido-fontal-api.railway.app`)

**IMPORTANTE:** Tras o despliege, abre `public/index.html` e verifica que os endpoints `/api/presuposto` e `/api/factura` funcionan.

---

## Variables de entorno (Railway)

Non se necesitan variables de entorno. O porto asígnase automaticamente coa variable `PORT`.

---

## Endpoints API

| Método | URL | Body (JSON) | Resposta |
|--------|-----|-------------|----------|
| POST | `/api/presuposto` | Ver payload abaixo | `.docx` (descarga) |
| POST | `/api/factura`    | Ver payload abaixo | `.docx` (descarga) |
| GET  | `/health`         | —                  | `{"ok":true}` |

### Payload orzamento
```json
{
  "num": "P26-001",
  "fecha": "2026-04-13",
  "validez": "2026-07-13",
  "cnome": "Comunidade de Propietarios...",
  "cnif": "H-27XXXXXX",
  "cdir": "Rúa...",
  "ccp": "27002 Lugo",
  "objeto": "Cambio de ventás de tellado",
  "lineas": [
    { "ud": "2", "concepto": "Liña en galego\nLínia en castellano", "precio": "1461.00", "subtotal": "2922.00" }
  ],
  "ivaPct": "21",
  "base": "2922.00",
  "ivaVal": "613.62",
  "total": "3535.62",
  "pago1": "1800.00",
  "pago2": "1735.62"
}
```

### Payload factura
```json
{
  "num": "32/26",
  "fecha": "2026-04-13",
  "cnome": "Carmen Pita Castro",
  "cnif": "33808393A",
  "cdir": "Rúa Armónica Nº 42, 2ºB",
  "ccp": "27002 Lugo",
  "lineas": [...],
  "ivaPct": "21",
  "base": "1271.00",
  "ivaVal": "266.91",
  "total": "1537.91",
  "banco": "Sabadell",
  "iban": "ES12 0081 0497 7400 0148 7849",
  "notas": "Factura de provisión de materiais..."
}
```
