#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Directorio nacional de empresas de residuos de manejo especial (RME) - México.

Etapas:
  1. BUSCAR    -> consultas "<palabra clave> <estado>" en Google (32 estados x N frases).
  2. FILTRAR   -> descarta dominios que no son empresas (gobierno, noticias, redes, directorios).
  3. ENRIQUECER-> visita el sitio de cada empresa (inicio, contacto, nosotros, servicios,
                  aviso de privacidad) y extrae razón social, RFC, teléfonos, correos,
                  dirección, servicios, tipos de residuos, autorizaciones, redes.
  4. GUARDAR   -> base acumulada (directorio_rme.json) + Excel (directorio_rme.xlsx).
                  Cada corrida mensual marca: Nuevo / Actualizado / Sin cambios / Sin respuesta.

Motores de búsqueda (--motor):
  places   Google Places API (Text Search).   Requiere GOOGLE_MAPS_KEY.
  cse      Google Custom Search JSON API.     Requiere GOOGLE_CSE_KEY y GOOGLE_CSE_CX.
  importar Lee resultados ya capturados (serp_google.json).

Nota: google.com no entrega resultados a programas sin navegador (exige JavaScript y
muestra CAPTCHA), por eso el bot mensual debe usar una de las dos APIs oficiales.

Uso:
  python3 scraper_rme.py --motor importar
  GOOGLE_MAPS_KEY=xxx python3 scraper_rme.py --motor places
  python3 scraper_rme.py --motor places --estados "Coahuila,Nuevo León"
"""
import argparse
import datetime as dt
import json
import os
import re
import sys
import time
import unicodedata
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

AQUI = os.path.dirname(os.path.abspath(__file__))
DATOS = os.environ.get("RME_DIR", AQUI)  # carpeta de trabajo (en la nube: /tmp/rme)
BASE_JSON = os.path.join(DATOS, "directorio_rme.json")
SALIDA_XLSX = os.path.join(DATOS, "directorio_rme.xlsx")
SERP_JSON = os.path.join(DATOS, "serp_google.json")
CORRIDA_JSON = os.path.join(DATOS, "ultima_corrida.json")


def AVISAR(etapa, hecho=0, total=0):
    """Aviso de avance; nube.py lo reemplaza para publicar el progreso."""
    print("  %s %s/%s" % (etapa, hecho, total) if total else "  " + etapa)

ESTADOS = [
    "Aguascalientes", "Baja California", "Baja California Sur", "Campeche", "Chiapas",
    "Chihuahua", "Ciudad de México", "Coahuila", "Colima", "Durango", "Estado de México",
    "Guanajuato", "Guerrero", "Hidalgo", "Jalisco", "Michoacán", "Morelos", "Nayarit",
    "Nuevo León", "Oaxaca", "Puebla", "Querétaro", "Quintana Roo", "San Luis Potosí",
    "Sinaloa", "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán",
    "Zacatecas",
]

# Frases de búsqueda. La consulta final es "<frase> <estado>".
PALABRAS_CLAVE = [
    "empresa recolección de residuos de manejo especial",
    "empresa reciclaje de residuos industriales",
    "empresa autorizada tratamiento y disposición de residuos",
    "recuperación y compra de materiales reciclables",
    "manejo integral de residuos empresa",
    "centro de acopio reciclaje cartón plástico chatarra",
]

UA = "Mozilla/5.0 (compatible; DirectorioRME-bot/1.0; directorio de empresas de manejo de residuos)"
CABECERAS = {"User-Agent": UA, "Accept-Language": "es-MX,es;q=0.9"}

# Dominios que nunca son la empresa buscada.
NO_EMPRESA = {
    "gobierno": [r"(^|\.)gob\.mx$", r"\.gov$", r"\.gob$", r"semarnat", r"profepa", r"inegi", r"congreso"],
    "educación": [r"\.edu\.mx$", r"\.edu$", r"unam\.mx$", r"ipn\.mx$", r"scielo", r"redalyc", r"researchgate", r"academia\.edu"],
    "red social": [r"facebook\.com$", r"instagram\.com$", r"linkedin\.com$", r"youtube\.com$", r"twitter\.com$",
                   r"^x\.com$", r"tiktok\.com$", r"pinterest\."],
    "directorio": [r"seccionamarilla", r"paginasamarillas", r"cylex", r"infobel", r"yelp\.", r"foursquare",
                   r"empresite", r"dnb\.com", r"kompass", r"guiamexico", r"mexicoo\.mx", r"directorio",
                   r"quienesquien", r"europages", r"solostocks", r"quiminet", r"cosmos\.com\.mx", r"tuugo",
                   r"opcionempleo", r"firmania", r"infoisinfo", r"enmexico\.com", r"mx\.near-place", r"waze\.com",
                   r"google\.", r"maps\.", r"tripadvisor", r"nexdu", r"todosnegocios", r"mapcarta", r"yellow\.place",
                   r"datamexico", r"econodata", r"rme\.mx$"],
    "empleo": [r"indeed\.", r"computrabajo", r"occ\.com\.mx", r"glassdoor", r"bumeran", r"jooble", r"talent\.com"],
    "noticias": [r"milenio", r"eluniversal", r"elfinanciero", r"excelsior", r"reforma\.com", r"elnorte\.com",
                 r"vanguardia", r"elsoldemexico", r"oem\.com\.mx", r"jornada", r"forbes", r"expansion\.mx",
                 r"eleconomista", r"infobae", r"proceso\.com", r"heraldo", r"zocalo", r"posta\.com", r"noticias",
                 r"diario", r"periodico", r"informador", r"telediario", r"msn\.com", r"media"],
    "enciclopedia/otro": [r"wikipedia", r"concepto\.de", r"scribd", r"slideshare", r"issuu", r"studocu", r"amazon\.",
                          r"mercadolibre", r"blogspot", r"wordpress\.com$", r"medium\.com$", r"change\.org",
                          r"dof\.gob", r"justia", r"vlex", r"un\.org$", r"greenpeace", r"wwf\."],
}

SERVICIOS = {
    "Recolección": [r"recolecci[oó]n", r"recolectamos", r"recolectar"],
    "Transporte": [r"transporte de residuos", r"traslado de residuos", r"transportaci[oó]n"],
    "Acopio": [r"\bacopio\b", r"centro de acopio"],
    "Almacenamiento temporal": [r"almacenamiento temporal", r"almac[eé]n temporal"],
    "Reciclaje": [r"recicla(je|mos|do|r)\b", r"recicladora"],
    "Recuperación / valorización": [r"recuperaci[oó]n de (materiales|residuos|metales)", r"valorizaci[oó]n",
                                    r"recuperadora", r"aprovechamiento de residuos"],
    "Compra de materiales": [r"compra(mos)? (de |y venta de )?(chatarra|cart[oó]n|pl[aá]stico|metal|material|pet|scrap|desperdicio)",
                             r"compra-?venta de (materiales|desperdicios)"],
    "Tratamiento": [r"tratamiento de residuos", r"tratamiento (f[ií]sico|qu[ií]mico|t[eé]rmico|biol[oó]gico)",
                    r"tratamiento y disposici[oó]n"],
    "Co-procesamiento": [r"co-?procesamiento", r"combustible alterno"],
    "Disposición final": [r"disposici[oó]n final", r"relleno sanitario", r"confinamiento"],
    "Destrucción fiscal / confidencial": [r"destrucci[oó]n (fiscal|de documentos|confidencial|de producto|de mercanc[ií]a)"],
    "Renta de contenedores": [r"renta de (contenedores|tolvas|compactadores|roll-?off)", r"contenedores? (roll-?off|para residuos)"],
    "Limpieza industrial": [r"limpieza industrial", r"limpieza de (tanques|fosas|drenajes)", r"vactor"],
    "Remediación": [r"remediaci[oó]n", r"atenci[oó]n a (derrames|emergencias ambientales)"],
    "Consultoría / gestoría ambiental": [r"consultor[ií]a ambiental", r"gestor[ií]a ambiental", r"plan(es)? de manejo",
                                         r"tr[aá]mites ambientales", r"asesor[ií]a ambiental"],
    "Manejo integral": [r"manejo integral", r"gesti[oó]n integral de residuos"],
}

RESIDUOS = {
    "Manejo especial (RME)": [r"manejo especial", r"\bRME\b"],
    "Sólidos urbanos (RSU)": [r"s[oó]lidos urbanos", r"\bRSU\b"],
    "Peligrosos (RP)": [r"residuos peligrosos", r"\bCRETI?B?\b"],
    "Biológico-infecciosos (RPBI)": [r"\bRPBI\b", r"biol[oó]gico[- ]infeccios"],
    "Industriales no peligrosos": [r"residuos industriales", r"no peligrosos"],
    "Cartón y papel": [r"cart[oó]n", r"\bpapel\b", r"archivo muerto"],
    "Plásticos": [r"pl[aá]sticos?", r"polietileno", r"polipropileno", r"\bHDPE\b", r"\bPVC\b", r"playo"],
    "PET": [r"\bPET\b"],
    "Metales / chatarra": [r"chatarra", r"metales", r"aluminio", r"\bcobre\b", r"fierro", r"acero", r"\bscrap\b"],
    "Vidrio": [r"\bvidrio\b"],
    "Madera / tarimas": [r"tarimas?", r"\bmadera\b", r"pallets?"],
    "Electrónicos (RAEE)": [r"electr[oó]nicos", r"\bRAEE\b", r"e-?waste", r"chatarra electr[oó]nica"],
    "Llantas": [r"llantas", r"neum[aá]ticos"],
    "Aceites usados": [r"aceites? (usados?|gastados?|lubricantes?|quemado|vegetal)"],
    "Orgánicos": [r"org[aá]nicos", r"composta"],
    "Construcción y demolición": [r"escombro", r"construcci[oó]n y demolici[oó]n", r"\bRCD\b", r"cascajo"],
    "Textiles": [r"textiles?", r"\btrapo\b", r"estopa"],
    "Lodos": [r"\blodos\b"],
    "Baterías / pilas": [r"bater[ií]as", r"acumuladores", r"\bpilas\b"],
    "Tóner / cartuchos": [r"t[oó]ner", r"cartuchos"],
    "Unicel": [r"unicel", r"poliestireno"],
    "Solventes / químicos": [r"solventes", r"qu[ií]micos caducos"],
}

AUTORIZACIONES = {
    "SEMARNAT": [r"SEMARNAT"],
    "SCT / SICT": [r"\bSCT\b", r"\bSICT\b"],
    "Autoridad ambiental estatal": [r"secretar[ií]a de(l)? medio ambiente", r"\bSEDEMA\b", r"\bSMA\b", r"\bSEMADET\b",
                                    r"\bSEMA\b", r"\bPROPAEG\b", r"autorizaci[oó]n estatal"],
    "ISO 14001": [r"ISO[ -]?14001"],
    "ISO 9001": [r"ISO[ -]?9001"],
    "Industria Limpia (PROFEPA)": [r"industria limpia"],
    "R2 / e-Stewards": [r"\bR2v?3?\b certif", r"e-?stewards"],
}

SUFIJO_RS = (r"S\.?\s?A\.?\s?P\.?\s?I\.?\s+de\s+C\.?\s?V\.?|S\.?\s?A\.?\s?B\.?\s+de\s+C\.?\s?V\.?|"
             r"S\.?\s?A\.?\s?S\.?\s+de\s+C\.?\s?V\.?|S\.?\s?A\.?\s+de\s+C\.?\s?V\.?|"
             r"S\.?\s+de\s+R\.?\s?L\.?(?:\s+de\s+C\.?\s?V\.?)?|S\.?\s?P\.?\s?R\.?\s+de\s+R\.?\s?L\.?(?:\s+de\s+C\.?\s?V\.?)?")
RE_RS = re.compile(
    r"((?:[A-ZÁÉÍÓÚÑ0-9&][\wÁÉÍÓÚÑáéíóúñü&.\-]*,?\s+(?:(?:de|del|la|las|los|y|e|en|para|el)\s+)*){1,8})(" + SUFIJO_RS + r")(?![\w])",
    re.UNICODE)
RE_RFC = re.compile(r"R\.?\s?F\.?\s?C\.?[^A-Z0-9]{0,25}([A-ZÑ&]{3,4})[- ]?(\d{2})(\d{2})(\d{2})[- ]?([A-Z0-9]{3})\b")
RE_MAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,10}")
RE_TEL = re.compile(r"(?:\+?52[\s.\-]?)?(?:1[\s.\-]?)?\(?\d{2,3}\)?[\s.\-]\d{3,4}[\s.\-]?\d{4}\b")
RE_TEL_CTX = re.compile(r"(?:tel[eé]fonos?|tels?|cel(?:ular)?|whats\s?app|ll[aá]manos|conmutador)[\s.:#]*((?:\+?52)?[\s.\-]?\d{10})", re.I)
RE_CP = re.compile(r"C\.?\s?P\.?\s*:?\s*(\d{5})\b")
RE_ENLACE_UTIL = re.compile(r"contact|nosotros|about|quienes|qui[eé]nes|empresa|servicio|aviso|privacidad|privacy|sucursal|ubicaci", re.I)


RE_COMPARTIDO = re.compile(r"^(sites\.google\.com|linktr\.ee|wa\.me|wa\.link|api\.whatsapp\.com|bit\.ly|goo\.gl|g\.page|"
                           r"business\.site|t\.me|beacons\.ai|linkr\.bio|taplink\.cc|walink\.co|msha\.ke|canva\.site|my\.canva\.site)$")


def sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


# Primeros dos dígitos del código postal -> estado (rangos de SEPOMEX).
CP_RANGOS = [(1, 16, "Ciudad de México"), (20, 20, "Aguascalientes"), (21, 22, "Baja California"),
             (23, 23, "Baja California Sur"), (24, 24, "Campeche"), (25, 27, "Coahuila"), (28, 28, "Colima"),
             (29, 30, "Chiapas"), (31, 33, "Chihuahua"), (34, 35, "Durango"), (36, 38, "Guanajuato"),
             (39, 41, "Guerrero"), (42, 43, "Hidalgo"), (44, 49, "Jalisco"), (50, 57, "Estado de México"),
             (58, 61, "Michoacán"), (62, 62, "Morelos"), (63, 63, "Nayarit"), (64, 67, "Nuevo León"),
             (68, 71, "Oaxaca"), (72, 75, "Puebla"), (76, 76, "Querétaro"), (77, 77, "Quintana Roo"),
             (78, 79, "San Luis Potosí"), (80, 82, "Sinaloa"), (83, 85, "Sonora"), (86, 86, "Tabasco"),
             (87, 89, "Tamaulipas"), (90, 90, "Tlaxcala"), (91, 96, "Veracruz"), (97, 97, "Yucatán"),
             (98, 99, "Zacatecas")]


def estado_por_cp(cp):
    n = int(cp[:2])
    return next((e for a, b, e in CP_RANGOS if a <= n <= b), "")


def dominio_de(url):
    h = urlparse(url if "//" in url else "//" + url).netloc.lower()
    return h[4:] if h.startswith("www.") else h


def motivo_descartar(dominio):
    for motivo, pats in NO_EMPRESA.items():
        if any(re.search(p, dominio) for p in pats):
            return motivo
    return ""


# ───────────────────────── 1. BÚSQUEDA ─────────────────────────
def buscar_places(estados, frases, clave, pausa=0.3):
    """Google Places API (New) Text Search. Hasta 60 resultados por consulta."""
    url = "https://places.googleapis.com/v1/places:searchText"
    campos = ("places.id,places.displayName,places.formattedAddress,places.nationalPhoneNumber,"
              "places.websiteUri,places.googleMapsUri,places.businessStatus,places.rating,"
              "places.userRatingCount,places.location,nextPageToken")
    cab = {"X-Goog-Api-Key": clave, "X-Goog-FieldMask": campos, "Content-Type": "application/json"}
    cands, bitacora = [], []
    for i_est, est in enumerate(estados):
        AVISAR("Buscando en Google Maps", i_est, len(estados))
        for fr in frases:
            consulta, token, n = "%s %s" % (fr, est), None, 0
            for _ in range(3):
                cuerpo = {"textQuery": consulta, "languageCode": "es", "regionCode": "MX", "pageSize": 20}
                if token:
                    cuerpo["pageToken"] = token
                r = requests.post(url, headers=cab, json=cuerpo, timeout=30)
                if r.status_code != 200:
                    bitacora.append((est, fr, consulta, "places", 0, "error %s: %s" % (r.status_code, r.text[:120])))
                    break
                d = r.json()
                for p in d.get("places", []):
                    n += 1
                    web = p.get("websiteUri", "")
                    cands.append({"estado": est, "consulta": consulta, "titulo": p.get("displayName", {}).get("text", ""),
                                  "url": web, "dominio": dominio_de(web) if web else "",
                                  "extra": {"place_id": p.get("id", ""), "direccion": p.get("formattedAddress", ""),
                                            "telefono": p.get("nationalPhoneNumber", ""), "maps": p.get("googleMapsUri", ""),
                                            "estatus_negocio": p.get("businessStatus", ""), "calificacion": p.get("rating", ""),
                                            "resenas": p.get("userRatingCount", ""),
                                            "lat": p.get("location", {}).get("latitude", ""),
                                            "lng": p.get("location", {}).get("longitude", "")}})
                token = d.get("nextPageToken")
                if not token:
                    break
                time.sleep(2)
            else:
                pass
            if n:
                bitacora.append((est, fr, consulta, "places", n, "ok"))
            time.sleep(pausa)
    return cands, bitacora


def buscar_cse(estados, frases, clave, cx, paginas=2, pausa=0.3):
    """Google Custom Search JSON API. 10 resultados por página."""
    cands, bitacora = [], []
    for est in estados:
        for fr in frases:
            consulta, n, estado_txt = "%s %s" % (fr, est), 0, "ok"
            for pag in range(paginas):
                r = requests.get("https://www.googleapis.com/customsearch/v1", timeout=30, params={
                    "key": clave, "cx": cx, "q": consulta, "gl": "mx", "hl": "es", "num": 10, "start": 1 + pag * 10})
                if r.status_code != 200:
                    estado_txt = "error %s: %s" % (r.status_code, r.text[:120])
                    break
                items = r.json().get("items", [])
                for it in items:
                    n += 1
                    cands.append({"estado": est, "consulta": consulta, "titulo": it.get("title", ""),
                                  "url": it.get("link", ""), "dominio": dominio_de(it.get("link", "")),
                                  "extra": {"fragmento": it.get("snippet", "")}})
                if len(items) < 10:
                    break
                time.sleep(pausa)
            bitacora.append((est, fr, consulta, "cse", n, estado_txt))
    return cands, bitacora


GEMINI_MODELO = os.environ.get("GEMINI_MODELO", "gemini-3-flash-preview")
GEMINI_GRUPO = 3  # frases por pregunta (cada pregunta dispara varias búsquedas de Google)


def buscar_gemini(estados, frases, proyecto, hilos=4):
    """Gemini (Vertex AI) con la herramienta de búsqueda de Google.

    Gemini busca en Google y devuelve las PÁGINAS que encontró (dominio de cada fuente). Aquí solo se
    usa para descubrir sitios; los datos de cada empresa los extrae después enriquecer() visitando el sitio.
    """
    if os.environ.get("GEMINI_TOKEN"):  # prueba local
        sesion = requests.Session()
        sesion.headers["Authorization"] = "Bearer " + os.environ["GEMINI_TOKEN"]
    else:
        import google.auth
        from google.auth.transport.requests import AuthorizedSession
        cred, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
        sesion = AuthorizedSession(cred)
    url = ("https://aiplatform.googleapis.com/v1/projects/%s/locations/global/publishers/google/models/%s:generateContent"
           % (proyecto, GEMINI_MODELO))
    tareas = [(est, frases[i:i + GEMINI_GRUPO]) for est in estados for i in range(0, len(frases), GEMINI_GRUPO)]

    def preguntar(est, grupo):
        texto = ("Busca en Google empresas privadas de %s, México, con estas búsquedas: %s. "
                 "Haz una lista lo más completa posible (al menos 15 si existen) de las empresas que ofrecen esos "
                 "servicios en %s, cada una con su nombre y la dirección de su sitio web oficial. "
                 "No incluyas directorios, noticias, redes sociales ni dependencias de gobierno."
                 % (est, "; ".join("«%s %s»" % (f, est) for f in grupo), est))
        cuerpo = {"contents": [{"role": "user", "parts": [{"text": texto}]}], "tools": [{"googleSearch": {}}],
                  "generationConfig": {"temperature": 0}}
        for intento in range(3):
            try:
                r = sesion.post(url, json=cuerpo, timeout=150)
            except Exception as e:
                r, err = None, str(e)[:120]
            if r is not None and r.status_code == 200:
                break
            if r is not None and r.status_code not in (429, 500, 503):
                return [], 0, "error %s: %s" % (r.status_code, r.text[:120].replace("\n", " "))
            time.sleep(8 * (intento + 1))
        else:
            return [], 0, "error: sin respuesta tras 3 intentos"
        cand = (r.json().get("candidates") or [{}])[0]
        meta = cand.get("groundingMetadata", {})
        doms = {}
        for ch in meta.get("groundingChunks", []):
            w = ch.get("web", {})
            d = dominio_de(w.get("domain") or w.get("title") or "")
            if re.match(r"^[a-z0-9.\-]+\.[a-z]{2,}$", d):
                doms.setdefault(d, w.get("title", d))
        respuesta = "".join(p.get("text", "") for p in cand.get("content", {}).get("parts", []))
        for u in re.findall(r"https?://[^\s)\]>\"']+", respuesta):
            d = dominio_de(u)
            if d and "vertexaisearch" not in d and re.match(r"^[a-z0-9.\-]+\.[a-z]{2,}$", d):
                doms.setdefault(d, d)
        return list(doms.items()), len(meta.get("webSearchQueries", [])), "ok"

    cands, bitacora, hechas = [], [], 0
    with ThreadPoolExecutor(max_workers=hilos) as ex:
        futuros = {ex.submit(preguntar, est, grupo): (est, grupo) for est, grupo in tareas}
        for fu in as_completed(futuros):
            est, grupo = futuros[fu]
            hechas += 1
            AVISAR("Buscando en Google web (Gemini)", hechas, len(tareas))
            try:
                doms, n_busq, estado_txt = fu.result()
            except Exception as e:
                doms, n_busq, estado_txt = [], 0, "error: %s" % str(e)[:120]
            fr = " + ".join(grupo)
            for d, tit in doms:
                cands.append({"estado": est, "consulta": "%s %s" % (fr, est), "titulo": tit, "url": "https://" + d,
                              "dominio": d, "extra": {}})
            bitacora.append((est, fr, "%s %s" % (fr, est), "gemini (%d búsquedas)" % n_busq, len(doms), estado_txt))
    bitacora.sort(key=lambda b: (estados.index(b[0]), b[1]))
    return cands, bitacora


def buscar_importar(ruta, estados):
    """Resultados de Google ya capturados: {"estados":[], "consultas":[], "resultados":[[iEstado,iFrase,dominio,titulo],...]}"""
    with open(ruta, encoding="utf-8") as f:
        d = json.load(f)
    cands, cuenta = [], Counter()
    for s, k, dom, tit in d["resultados"]:
        est, fr = d["estados"][s], d["consultas"][k]
        if estados and est not in estados:
            continue
        dom = dominio_de(dom.strip().rstrip("/"))
        if not dom or "." not in dom:
            continue
        cuenta[(est, fr)] += 1
        cands.append({"estado": est, "consulta": "%s %s" % (fr, est), "titulo": tit, "url": "https://" + dom,
                      "dominio": dom, "extra": {}})
    bitacora = [(e, f, "%s %s" % (f, e), "google (captura)", n, "ok") for (e, f), n in sorted(cuenta.items())]
    return cands, bitacora


# ───────────────────────── 3. ENRIQUECIMIENTO ─────────────────────────
def bajar(url, t=12):
    try:
        r = requests.get(url, headers=CABECERAS, timeout=t, allow_redirects=True, stream=True)
        if r.status_code >= 400 or "html" not in r.headers.get("Content-Type", "html"):
            return None, r.url
        crudo = r.raw.read(1500000, decode_content=True)
        declarado = re.search(r"charset=([\w\-]+)", r.headers.get("Content-Type", ""), re.I)
        try:
            return crudo.decode("utf-8"), r.url
        except UnicodeDecodeError:
            try:
                return crudo.decode(declarado.group(1) if declarado else "latin-1", errors="ignore"), r.url
            except LookupError:
                return crudo.decode("latin-1", errors="ignore"), r.url
    except Exception:
        return None, url


def marcar(texto, diccionario):
    hallados = []
    for etiqueta, pats in diccionario.items():
        if any(re.search(p, texto, re.I) for p in pats):
            hallados.append(etiqueta)
    return hallados


def limpiar_tel(s):
    d = re.sub(r"\D", "", s)
    if len(d) == 13 and d.startswith("521"):
        d = d[3:]
    elif len(d) == 12 and d.startswith("52"):
        d = d[2:]
    elif len(d) == 11 and d.startswith("1"):
        d = d[1:]
    if len(d) != 10 or d[0] in "01" or len(set(d)) < 4:
        return ""
    return "%s %s %s" % (d[:3], d[3:6], d[6:])


def nombre_de(sopa, dominio):
    m = sopa.find("meta", property="og:site_name")
    if m and m.get("content", "").strip():
        return m["content"].strip()[:80]
    t = sopa.title.get_text(" ", strip=True) if sopa.title else ""
    partes = [p.strip() for p in re.split(r"\s[|\-–—·:»]\s", t) if p.strip()]
    genericas = re.compile(r"^(inicio|home|bienvenid|contacto|nosotros|servicios|p[aá]gina)", re.I)
    partes = [p for p in partes if not genericas.match(p)]
    if partes:
        return min(partes, key=len)[:80] if len(partes) > 1 and len(min(partes, key=len)) > 3 else partes[0][:80]
    return dominio.split(".")[0].capitalize()


def enriquecer(dominio, url_inicio):
    """Visita inicio + hasta 4 páginas internas útiles y extrae los campos."""
    reg = {"respondio": False, "paginas": 0}
    html, final = bajar(url_inicio)
    for alterna in ("https://www." + dominio, "http://" + dominio):
        if html is None:
            html, final = bajar(alterna)
    if html is None:
        return reg
    reg["respondio"] = True
    reg["url_final"] = final
    sopa = BeautifulSoup(html, "html.parser")
    reg["nombre"] = nombre_de(sopa, dominio)
    md = sopa.find("meta", attrs={"name": "description"}) or sopa.find("meta", property="og:description")
    reg["descripcion"] = (md.get("content", "").strip()[:300] if md else "")

    internos, vistos = [], set()
    for a in sopa.find_all("a", href=True):
        h = urljoin(final, a["href"]).split("#")[0]
        if dominio_de(h) != dominio_de(final) or h in vistos or re.search(r"\.(pdf|jpg|png|zip|docx?)$", h, re.I):
            continue
        if RE_ENLACE_UTIL.search(a["href"]) or RE_ENLACE_UTIL.search(sin_acentos(a.get_text(" ", strip=True))):
            vistos.add(h)
            prioridad = 0 if re.search(r"contact", h, re.I) else 1 if re.search(r"privac|aviso", h, re.I) else 2
            internos.append((prioridad, h))
    sopas = [sopa]
    for _, h in sorted(internos)[:4]:
        ht, _u = bajar(h, 10)
        if ht:
            sopas.append(BeautifulSoup(ht, "html.parser"))
    reg["paginas"] = len(sopas)

    tels, mails, redes, direccion = [], [], {}, ""
    for s in sopas:
        for a in s.find_all("a", href=True):
            h = a["href"].strip()
            hl = h.lower()
            if hl.startswith("tel:"):
                tels.append(limpiar_tel(h[4:]))
            elif hl.startswith("mailto:"):
                mails.append(h[7:].split("?")[0].strip().lower())
            elif "wa.me/" in hl or "api.whatsapp.com" in hl:
                m = re.search(r"(\d{10,13})", h)
                if m:
                    redes.setdefault("WhatsApp", limpiar_tel(m.group(1)))
            else:
                for red, pat in (("Facebook", "facebook.com/"), ("LinkedIn", "linkedin.com/"), ("Instagram", "instagram.com/")):
                    if pat in hl and "share" not in hl and red not in redes:
                        redes[red] = h.split("?")[0]
        for sc in s.find_all("script", type="application/ld+json"):
            try:
                datos = json.loads(sc.string or "")
            except Exception:
                continue
            for nodo in (datos if isinstance(datos, list) else datos.get("@graph", [datos]) if isinstance(datos, dict) else []):
                if not isinstance(nodo, dict):
                    continue
                ad = nodo.get("address")
                if isinstance(ad, dict) and not direccion:
                    direccion = ", ".join(str(ad.get(k, "")) for k in ("streetAddress", "addressLocality", "addressRegion", "postalCode") if ad.get(k))
                if nodo.get("telephone"):
                    tels.append(limpiar_tel(str(nodo["telephone"])))
                if nodo.get("legalName"):
                    reg["razon_social"] = str(nodo["legalName"])[:120]
        for x in s(["script", "style", "noscript", "svg"]):
            x.decompose()
    texto = "\n".join(re.sub(r"[ \t\r\f\v]+", " ", s.get_text("\n", strip=True)) for s in sopas)
    plano = re.sub(r"\s+", " ", texto)

    tels += [limpiar_tel(m.group(0)) for m in RE_TEL.finditer(plano)]
    tels += [limpiar_tel(m.group(1)) for m in RE_TEL_CTX.finditer(plano)]
    mails += [m.lower() for m in RE_MAIL.findall(plano)]
    mails = [m.strip(".") for m in mails if not re.search(r"\.(png|jpe?g|gif|webp|svg)$|sentry|wixpress|example\.|tudominio|@2x|correo@|email@|nombre@", m)]
    reg["telefonos"] = list(dict.fromkeys(t for t in tels if t))[:5]
    reg["correos"] = list(dict.fromkeys(mails))[:5]
    reg["redes"] = redes

    if "razon_social" not in reg:
        halladas = Counter()
        for m in RE_RS.finditer(plano):
            nom = re.sub(r"^(?:.*\b(?:Privacidad|Responsable|Empresa|Somos|Copyright|Derechos Reservados|Reservados|Bienvenidos a|Contacto)\b[\s,:]*)", "", m.group(1).strip(" ,"))
            nom = re.split(r"(?<=[.:;])\s+", nom.strip(" ,.-"))[-1]
            nom = re.sub(r"^(?:(?:[\d.,©®]+|Ver|Lead|Site|CP|Todos|Los|Derechos)\s+)+", "", nom).strip(" ,.-")
            if 2 < len(nom) < 90:
                halladas[re.sub(r"\s+", " ", nom + " " + m.group(2))] += 1
        if halladas:
            reg["razon_social"] = halladas.most_common(1)[0][0]
    m = RE_RFC.search(plano)
    if m and 1 <= int(m.group(3)) <= 12 and 1 <= int(m.group(4)) <= 31:
        reg["rfc"] = "".join(m.groups())

    cp = RE_CP.search(plano)
    if cp:
        reg["cp"] = cp.group(1)
        if not direccion:
            ini = max(0, cp.start() - 130)
            trozo = plano[ini:cp.end() + 45]
            antes, despues = trozo[:cp.start() - ini], trozo[cp.start() - ini:]
            vias = list(re.finditer(r"\b(?:Calle|Av\.|Avenida|Blvd\.?|Boulevard|Bulevar|Carretera|Carr\.|Calzada|Calz\.|Perif[eé]rico|Prolongaci[oó]n|Privada|Km\.?|Libramiento|Circuito|Camino|Paseo|Autopista)\s", antes))
            corte = re.search(r"(?:Direcci[oó]n|Domicilio|Ubicaci[oó]n|ubicad[oa]s? en|Oficinas?|Planta|Matriz|Sucursal)\s*:?\s*", antes)
            inicio = vias[0].start() if vias else (corte.end() if corte else len(antes))
            cola = re.match(r"C\.?\s?P\.?\s*:?\s*\d{5}(?:[\s,]+[A-ZÁÉÍÓÚÑ][\wáéíóúñ]+(?:[\s,]+[A-ZÁÉÍÓÚÑ][\wáéíóúñ.]+){0,4})?", despues)
            direccion = (antes[inicio:] + (cola.group(0) if cola else despues[:14])).strip(" ,.;:-")
    reg["direccion"] = direccion[:220]
    base_dir = sin_acentos(direccion).lower()
    reg["estado_sede"] = next((e for e in sorted(ESTADOS, key=len, reverse=True) if sin_acentos(e).lower() in base_dir), "")
    if not reg["estado_sede"]:
        for abr, e in (("cdmx", "Ciudad de México"), ("n.l.", "Nuevo León"), ("edo. de mex", "Estado de México"), ("edomex", "Estado de México"), ("b.c.", "Baja California"), ("s.l.p.", "San Luis Potosí"), ("qro", "Querétaro"), ("jal.", "Jalisco"), ("coah", "Coahuila")):
            if abr in base_dir:
                reg["estado_sede"] = e
                break
    if not reg["estado_sede"] and reg.get("cp"):
        reg["estado_sede"] = estado_por_cp(reg["cp"])
    plano_sa = sin_acentos(plano).lower()
    reg["estados_mencionados"] = [e for e in ESTADOS if re.search(r"\b" + re.escape(sin_acentos(e).lower()) + r"\b", plano_sa)]

    reg["servicios"] = marcar(plano, SERVICIOS)
    reg["residuos"] = marcar(plano, RESIDUOS)
    reg["autorizaciones"] = marcar(plano, AUTORIZACIONES)
    aut = re.search(r"(?:autorizaci[oó]n|registro|permiso)[^.]{0,60}?\b(?:n[uú]m(?:ero)?\.?|no\.|#)\s*:?\s*([A-Z0-9][A-Z0-9\-/.]{5,30})", plano, re.I)
    reg["num_autorizacion"] = aut.group(1).rstrip(".") if aut and re.search(r"\d", aut.group(1)) else ""
    palabras_sector = len(re.findall(r"residuos|recicl|recolecci|chatarra|acopio|disposici[oó]n final|desperdicio|scrap", plano, re.I))
    reg["relevancia"] = min(100, len(reg["servicios"]) * 8 + len(reg["residuos"]) * 3 + min(palabras_sector, 30)
                            + (10 if reg["telefonos"] or reg["correos"] else 0))
    reg["menciones_sector"] = palabras_sector
    return reg


# ───────────────────────── 4. BASE + EXCEL ─────────────────────────
CAMPOS_COMPARAR = ["nombre", "razon_social", "rfc", "telefonos", "correos", "direccion", "servicios", "residuos"]

COLUMNAS = [  # (encabezado, ancho, descripción)
    ("ID", 8, "Consecutivo del directorio."),
    ("Nombre comercial", 34, "Nombre del sitio (og:site_name / título) o nombre en Google Maps."),
    ("Razón social", 40, "Denominación legal detectada (S.A. de C.V., S. de R.L., etc.), normalmente en el aviso de privacidad."),
    ("RFC", 16, "RFC publicado en el sitio (solo si aparece junto a la sigla RFC)."),
    ("Estado sede", 20, "Estado detectado en la dirección publicada."),
    ("Estados donde aparece (búsqueda)", 34, "Estados cuyas consultas en Google devolvieron a esta empresa."),
    ("Cobertura mencionada", 34, "Estados que el sitio menciona en su texto."),
    ("Dirección", 50, "Dirección publicada (datos estructurados o texto alrededor del C.P.)."),
    ("C.P.", 9, "Código postal."),
    ("Teléfonos", 30, "Hasta 5 teléfonos (10 dígitos)."),
    ("WhatsApp", 15, "Número del enlace de WhatsApp."),
    ("Correos", 38, "Hasta 5 correos publicados."),
    ("Sitio web", 34, "Página de la empresa."),
    ("Servicios", 48, "Servicios detectados por palabras clave."),
    ("Tipos de residuos", 48, "Tipos de residuos / materiales detectados."),
    ("Autorizaciones / certificaciones", 30, "Menciones de SEMARNAT, SCT, autoridad estatal, ISO, etc."),
    ("Núm. de autorización", 22, "Primer número de autorización o registro encontrado en el texto."),
    ("Descripción", 55, "Meta descripción del sitio."),
    ("Facebook", 28, ""), ("LinkedIn", 28, ""), ("Instagram", 28, ""),
    ("Google Maps", 28, "Ficha de Google Maps (solo motor places)."),
    ("Calificación Google", 12, "Solo motor places."),
    ("Relevancia (0-100)", 12, "Puntaje según servicios, residuos y menciones del sector."),
    ("Consultas que lo encontraron", 14, "Número de consultas distintas."),
    ("Frases de búsqueda", 45, "Frases con las que apareció."),
    ("Estatus del mes", 15, "Nuevo / Actualizado / Sin cambios / Sin respuesta."),
    ("Primera detección", 14, ""), ("Última verificación", 14, ""),
    ("Fuente", 16, "Motor de búsqueda de origen."),
]


def a_fila(i, k, r):
    j = lambda v: "; ".join(v) if isinstance(v, list) else (v or "")
    rd = r.get("redes", {})
    return [i, r.get("nombre", ""), r.get("razon_social", ""), r.get("rfc", ""), r.get("estado_sede", ""),
            j(r.get("estados_busqueda", [])), j(r.get("estados_mencionados", [])), r.get("direccion", ""), r.get("cp", ""),
            j(r.get("telefonos", [])), rd.get("WhatsApp", ""), j(r.get("correos", [])), r.get("sitio", ""),
            j(r.get("servicios", [])), j(r.get("residuos", [])), j(r.get("autorizaciones", [])), r.get("num_autorizacion", ""),
            r.get("descripcion", ""), rd.get("Facebook", ""), rd.get("LinkedIn", ""), rd.get("Instagram", ""),
            r.get("maps", ""), r.get("calificacion", ""), r.get("relevancia", ""), r.get("n_consultas", ""),
            j(r.get("frases", [])), r.get("estatus_mes", ""), r.get("primera_deteccion", ""), r.get("ultima_verificacion", ""),
            r.get("fuente", "")]


def hoja(wb, titulo, encabezados, filas, anchos=None, primera=False):
    ws = wb.active if primera else wb.create_sheet()
    ws.title = titulo
    ws.append(encabezados)
    for f in filas:
        ws.append(f)
    relleno = PatternFill("solid", fgColor="1F4E46")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = relleno
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "C2" if titulo == "Directorio" else "A2"
    if filas:
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(encabezados)), len(filas) + 1)
    for i, _ in enumerate(encabezados, 1):
        ws.column_dimensions[get_column_letter(i)].width = (anchos[i - 1] if anchos else 22)
    return ws


def exportar_xlsx(base, descartados, bitacora, ruta, hoy):
    regs = [(k, r) for k, r in base.items() if r.get("es_empresa")]
    regs.sort(key=lambda kr: (kr[1].get("estados_busqueda", ["~"])[0], -(kr[1].get("relevancia") or 0)))
    wb = Workbook()
    ws = hoja(wb, "Directorio", [c[0] for c in COLUMNAS], [a_fila(i, k, r) for i, (k, r) in enumerate(regs, 1)],
              [c[1] for c in COLUMNAS], primera=True)
    col_web = [c[0] for c in COLUMNAS].index("Sitio web") + 1
    for fila in ws.iter_rows(min_row=2):
        for c in fila:
            c.alignment = Alignment(vertical="top", wrap_text=False)
        c = fila[col_web - 1]
        if c.value:
            c.hyperlink = c.value
            c.font = Font(color="0563C1", underline="single")

    por_estado = defaultdict(lambda: [0, 0, 0, 0, 0])
    for _, r in regs:
        for e in r.get("estados_busqueda", []):
            p = por_estado[e]
            p[0] += 1
            p[1] += 1 if r.get("telefonos") else 0
            p[2] += 1 if r.get("correos") else 0
            p[3] += 1 if r.get("razon_social") else 0
            p[4] += 1 if r.get("rfc") else 0
    filas = [[e] + por_estado[e] + [sum(1 for _, r in regs if r.get("estado_sede") == e)] for e in ESTADOS]
    hoja(wb, "Resumen por estado", ["Estado", "Empresas encontradas", "Con teléfono", "Con correo", "Con razón social",
                                    "Con RFC", "Con sede detectada en el estado"], filas, [24, 16, 14, 14, 16, 12, 22])
    hoja(wb, "Consultas", ["Estado", "Frase", "Consulta enviada", "Motor", "Resultados", "Estado de la consulta"],
         [list(b) for b in bitacora], [22, 52, 70, 18, 12, 30])
    hoja(wb, "Descartados", ["Dominio", "Título en Google", "Estado(s)", "Motivo"],
         [[d["dominio"], d["titulo"], "; ".join(d["estados"]), d["motivo"]] for d in descartados], [38, 60, 40, 30])
    hoja(wb, "Campos", ["Campo", "Cómo se obtiene"], [[c[0], c[2]] for c in COLUMNAS] + [
        ["", ""], ["Corrida", hoy], ["Empresas en el directorio", len(regs)]], [34, 110])
    wb.save(ruta)
    return len(regs)


BUSQUEDA_JSON = os.path.join(DATOS, "busqueda_cache.json")
BUSQUEDA_DIAS = float(os.environ.get("BUSQUEDA_DIAS", "7"))  # las búsquedas de paga se reutilizan este tiempo


def motores_disponibles(serp=SERP_JSON):
    """Todas las fuentes que se pueden usar; los resultados se combinan sin duplicados."""
    m = []
    if os.environ.get("GOOGLE_MAPS_KEY"):
        m.append("places")
    if os.environ.get("GEMINI_PROYECTO"):
        m.append("gemini")
    if os.environ.get("GOOGLE_CSE_KEY") and os.environ.get("GOOGLE_CSE_CX"):
        m.append("cse")
    if os.path.exists(serp):
        m.append("importar")
    return m


def buscar(motores, estados, serp):
    """Corre cada motor. Las búsquedas por API (de paga) se guardan por motor y se reutilizan BUSQUEDA_DIAS días."""
    cands, bitacora, reutilizadas = [], [], []
    cache = {}
    if os.path.exists(BUSQUEDA_JSON):
        with open(BUSQUEDA_JSON, encoding="utf-8") as f:
            cache = json.load(f)
        if "cands" in cache:  # formato anterior: un solo bloque (places)
            cache = {"places": {"fecha": cache["fecha"], "estados": cache["estados"], "cands": cache["cands"],
                                "bitacora": cache["bitacora"]}}
    cambio = False
    for m in motores:
        if m == "importar":
            c, b = buscar_importar(serp, estados)
        else:
            g = cache.get(m)
            if g:
                edad = (dt.datetime.utcnow() - dt.datetime.strptime(g["fecha"], "%Y-%m-%dT%H:%M:%SZ")).total_seconds() / 86400
                if g.get("estados") != estados or edad > BUSQUEDA_DIAS:
                    g = None
            if g:
                c, b = g["cands"], [tuple(x) for x in g["bitacora"]]
                reutilizadas.append("%s (%s)" % (m, g["fecha"][:10]))
            else:
                if m == "places":
                    c, b = buscar_places(estados, PALABRAS_CLAVE, os.environ["GOOGLE_MAPS_KEY"])
                elif m == "gemini":
                    c, b = buscar_gemini(estados, PALABRAS_CLAVE, os.environ["GEMINI_PROYECTO"])
                else:
                    c, b = buscar_cse(estados, PALABRAS_CLAVE, os.environ["GOOGLE_CSE_KEY"], os.environ["GOOGLE_CSE_CX"])
                if c:
                    cache[m] = {"fecha": dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"), "estados": estados,
                                "cands": c, "bitacora": b}
                    cambio = True
        for x in c:
            x["motor"] = m
        cands += c
        bitacora += b
    if cambio:
        with open(BUSQUEDA_JSON, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False)
    return cands, bitacora, ", ".join(reutilizadas)


def ejecutar(motor="auto", estados=None, hilos=12, min_relevancia=25, salida=SALIDA_XLSX, serp=SERP_JSON):
    hoy = dt.date.today().isoformat()
    estados = estados or ESTADOS
    motores = motores_disponibles(serp) if motor == "auto" else [motor]
    if not motores:
        sys.exit("No hay fuente de búsqueda: falta GOOGLE_MAPS_KEY, GOOGLE_CSE_KEY+GOOGLE_CSE_CX o serp_google.json")

    # 1. Buscar
    AVISAR("Buscando en Google")
    cands, bitacora, de_cache = buscar(motores, estados, serp)
    print("Resultados de búsqueda: %d (consultas: %d)" % (len(cands), len(bitacora)))

    # 2. Agrupar por empresa (dominio o ficha de Maps) y filtrar
    grupos, descartados = {}, {}
    for c in cands:
        if c["extra"].get("place_id") and c["dominio"] and (motivo_descartar(c["dominio"]) or RE_COMPARTIDO.search(c["dominio"])):
            # Negocio de Maps cuyo "sitio" es una red social o un servicio compartido: se conserva sin visitar el sitio
            c = dict(c, dominio="", extra=dict(c["extra"], web=c["url"]))
        clave_emp = c["dominio"] or ("place:" + c["extra"].get("place_id", c["titulo"]))
        motivo = motivo_descartar(c["dominio"]) if c["dominio"] else ""
        destino = descartados if motivo else grupos
        g = destino.setdefault(clave_emp, {"dominio": c["dominio"], "titulo": c["titulo"], "url": c["url"], "estados": [],
                                           "frases": [], "consultas": set(), "extra": c["extra"], "motivo": motivo,
                                           "motores": []})
        if c.get("motor", motores[0]) not in g["motores"]:
            g["motores"].append(c.get("motor", motores[0]))
        if c["extra"].get("place_id") and not g["extra"].get("place_id"):
            g["extra"], g["titulo"] = c["extra"], c["titulo"]  # la ficha de Maps manda sobre el resultado web
        if c["estado"] not in g["estados"]:
            g["estados"].append(c["estado"])
        fr = c["consulta"][: -len(c["estado"])].strip()
        if fr not in g["frases"]:
            g["frases"].append(fr)
        g["consultas"].add(c["consulta"])
    print("Candidatos únicos: %d | descartados por tipo de dominio: %d" % (len(grupos), len(descartados)))

    # 3. Enriquecer
    base = {}
    if os.path.exists(BASE_JSON):
        with open(BASE_JSON, encoding="utf-8") as f:
            base = json.load(f)
    resultados = {}
    con_sitio = {k: g for k, g in grupos.items() if g["dominio"]}
    AVISAR("Visitando sitios de empresas", 0, len(con_sitio))
    with ThreadPoolExecutor(max_workers=hilos) as ex:
        futuros = {ex.submit(enriquecer, g["dominio"], "https://" + g["dominio"]): k for k, g in con_sitio.items()}
        for n, fu in enumerate(as_completed(futuros), 1):
            try:
                resultados[futuros[fu]] = fu.result()
            except Exception as e:
                resultados[futuros[fu]] = {"respondio": False, "error": str(e)[:100]}
            if n % 5 == 0 or n == len(con_sitio):
                AVISAR("Visitando sitios de empresas", n, len(con_sitio))

    AVISAR("Guardando resultados")
    cuenta = Counter()
    lista_desc = [{"dominio": d["dominio"], "titulo": d["titulo"], "estados": d["estados"], "motivo": d["motivo"]}
                  for d in descartados.values()]
    for k, g in grupos.items():
        r = resultados.get(k, {"respondio": False})
        x = g["extra"]
        previo = base.get(k, {})
        if g["dominio"] and not r.get("respondio"):
            if previo.get("es_empresa"):
                previo.update({"estatus_mes": "Sin respuesta", "ultima_verificacion": hoy})
                base[k] = previo
                cuenta["Sin respuesta"] += 1
            else:
                lista_desc.append({"dominio": g["dominio"], "titulo": g["titulo"], "estados": g["estados"], "motivo": "el sitio no respondió"})
            continue
        nuevo = dict(r)
        nuevo.update({
            "sitio": ("https://" + g["dominio"]) if g["dominio"] else x.get("web", ""),
            "nombre": x.get("place_id") and g["titulo"] or r.get("nombre") or g["titulo"],
            "estados_busqueda": sorted(set(previo.get("estados_busqueda", [])) | set(g["estados"]), key=ESTADOS.index),
            "frases": g["frases"], "n_consultas": len(g["consultas"]), "fuente": ", ".join(sorted(g["motores"])),
            "maps": x.get("maps", ""), "calificacion": x.get("calificacion", ""),
        })
        if x.get("telefono"):
            nuevo["telefonos"] = list(dict.fromkeys([limpiar_tel(x["telefono"]) or x["telefono"]] + nuevo.get("telefonos", [])))[:5]
        if x.get("direccion"):
            nuevo["direccion"] = x["direccion"]
            d_sa = sin_acentos(x["direccion"]).lower()
            cp_maps = re.findall(r"\b(\d{5})\b", x["direccion"])
            if cp_maps:
                nuevo["cp"] = cp_maps[-1]
            nuevo["estado_sede"] = (estado_por_cp(cp_maps[-1]) if cp_maps else "") or next(
                (e for e in sorted(ESTADOS, key=len, reverse=True) if sin_acentos(e).lower() in d_sa), nuevo.get("estado_sede", ""))
        if not nuevo.get("servicios"):  # sin sitio web: al menos lo que diga el nombre del negocio
            nuevo["servicios"] = marcar(nuevo["nombre"], SERVICIOS)
            nuevo["residuos"] = nuevo.get("residuos") or marcar(nuevo["nombre"], RESIDUOS)
        de_maps = bool(x.get("place_id"))
        nuevo["es_empresa"] = de_maps or ((nuevo.get("relevancia") or 0) >= min_relevancia and bool(nuevo.get("servicios")))
        if not nuevo["es_empresa"]:
            lista_desc.append({"dominio": g["dominio"], "titulo": g["titulo"], "estados": g["estados"],
                               "motivo": "baja relevancia (%s)" % nuevo.get("relevancia", 0)})
            continue
        if not previo:
            nuevo["estatus_mes"], nuevo["primera_deteccion"] = "Nuevo", hoy
        else:
            cambio = any(previo.get(c) != nuevo.get(c) for c in CAMPOS_COMPARAR)
            nuevo["estatus_mes"] = "Actualizado" if cambio else "Sin cambios"
            nuevo["primera_deteccion"] = previo.get("primera_deteccion", hoy)
        nuevo["ultima_verificacion"] = hoy
        cuenta[nuevo["estatus_mes"]] += 1
        base[k] = nuevo

    with open(BASE_JSON, "w", encoding="utf-8") as f:
        json.dump(base, f, ensure_ascii=False, indent=1)
    lista_desc.sort(key=lambda d: (d["motivo"], d["dominio"]))
    n = exportar_xlsx(base, lista_desc, bitacora, salida, hoy)
    resumen = {"fecha": dt.datetime.utcnow().isoformat(timespec="seconds") + "Z", "motor": ", ".join(motores),
               "busqueda_reutilizada_de": de_cache or "",
               "consultas": len(bitacora), "resultados_busqueda": len(cands), "candidatos": len(grupos),
               "empresas": n, "descartados": len(lista_desc), "por_estatus": dict(cuenta),
               "estados_con_resultados": sorted({b[0] for b in bitacora if b[4]}, key=ESTADOS.index),
               "bitacora": [list(b) for b in bitacora], "lista_descartados": lista_desc}
    with open(CORRIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(resumen, f, ensure_ascii=False)
    print("Empresas en el directorio: %d | descartados: %d" % (n, len(lista_desc)))
    print("Excel: %s" % salida)
    return resumen


def main():
    ap = argparse.ArgumentParser(description="Directorio de empresas de residuos de manejo especial en México")
    ap.add_argument("--motor", choices=["auto", "places", "cse", "importar"], default="auto")
    ap.add_argument("--estados", default="", help="Lista separada por comas (por omisión, los 32)")
    ap.add_argument("--serp", default=SERP_JSON, help="Archivo de resultados para --motor importar")
    ap.add_argument("--hilos", type=int, default=12)
    ap.add_argument("--min-relevancia", type=int, default=25)
    ap.add_argument("--salida", default=SALIDA_XLSX)
    a = ap.parse_args()
    estados = [e.strip() for e in a.estados.split(",") if e.strip()] or None
    ejecutar(a.motor, estados, a.hilos, a.min_relevancia, a.salida, a.serp)


if __name__ == "__main__":
    main()
