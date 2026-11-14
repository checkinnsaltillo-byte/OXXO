# OXXO · Directorio de empresas de residuos de manejo especial (México)

Scraper + dashboard. Busca en Google «frase + estado» para los 32 estados, visita el sitio de cada
empresa y arma un directorio (nombre, razón social, RFC, servicios, tipos de residuos, ubicación,
teléfonos, correos, sitio web, autorizaciones).

| Archivo | Qué es |
|---|---|
| `scraper_rme.py` | Búsqueda, filtrado, extracción y Excel |
| `nube.py` | Corrida en Google Cloud (Cloud Run Job `oxxo-scraper`) |
| `app.py` + `static/index.html` | Dashboard (servicio de Cloud Run `oxxo`) con el botón «Ejecutar extracción» |
| `serp_google.json` | Búsquedas de Google capturadas (se usan mientras no haya clave de API) |

## Google Cloud (proyecto `check-in-493804`, región `us-central1`)

- Servicio `oxxo`: dashboard público.
- Job `oxxo-scraper`: hace la extracción; lo lanza el botón del dashboard.
- Bucket `check-in-493804-oxxo`: `directorio_rme.json`, `directorio_rme.xlsx`, `ultima_corrida.json`, `estado.json`.
- Cuenta de servicio `oxxo-run`.

## Fuente de búsqueda

El Job elige sola: `GOOGLE_MAPS_KEY` → Places API; `GOOGLE_CSE_KEY` + `GOOGLE_CSE_CX` → Custom Search;
sin claves → reutiliza `serp_google.json` del bucket.
