#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Perfil y calificación de confianza de cada empresa del directorio.

  1. identidad()  Cruce con el DENUE (INEGI) por teléfono, sitio, correo o ubicación+nombre
                  -> razón social, actividad, tamaño, antigüedad. (Después corre REPSE -> RFC.)
  2. perfilar()   Listas del SAT (art. 69, 69-B y 69-B Bis), padrones SEMARNAT de residuos
                  peligrosos, señales del sitio (antigüedad del dominio, correo propio, reseñas)
                  y calificación 0-100 por bloques.

Todas las fuentes son archivos públicos descargables (sin sesión ni CAPTCHA):
  DENUE     https://www.inegi.org.mx/app/descarga/?ti=6  (descarga masiva por sector)
  SAT       http://omawww.sat.gob.mx/cifras_sat/  (datos abiertos, listados 69 / 69-B)
  SEMARNAT  https://www.gob.mx/semarnat/documentos/empresas-autorizadas-para-el-manejo-de-residuos-peligrosos
  Dominios  whois.mx (NIC México) y rdap.org
"""
import csv
import datetime as dt
import io
import math
import os
import re
import socket
import tempfile
import time
import zipfile
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher

import requests

import scraper_rme as s

UA = {"User-Agent": "Mozilla/5.0 (compatible; DirectorioRME-bot/1.0; cruce de datos abiertos)"}
DENUE_DIAS = float(os.environ.get("DENUE_DIAS", "178"))
DENUE_SECTORES = ["56", "43", "31-33", "48-49", "54", "81_1", "81_2", "46111", "46112-46311", "46321-46531",
                  "46591-46911", "23", "22", "21", "11", "51", "52", "53", "55", "61", "62", "71", "72_1", "72_2", "93"]
CORREO_GENERICO = {"gmail.com", "hotmail.com", "outlook.com", "yahoo.com", "yahoo.com.mx", "live.com", "live.com.mx",
                   "prodigy.net.mx", "icloud.com", "msn.com", "hotmail.es", "outlook.es", "aol.com", "me.com",
                   "protonmail.com", "telmexmail.com", "infinitummail.com"}
HOY = lambda: dt.date.today().isoformat()  # noqa: E731


def tel10(x):
    d = re.sub(r"\D", "", x or "")
    return d[-10:] if len(d) >= 10 else ""


def avisar(etapa, hecho=0, total=0):
    s.AVISAR(etapa, hecho, total)


def bajar(url, intentos=3, t=180):
    for i in range(intentos):
        try:
            r = requests.get(url, headers=UA, timeout=t)
            if r.status_code == 200:
                return r.content
            if r.status_code == 404:
                return None
        except Exception:
            pass
        time.sleep(3 * (i + 1))
    return None


def texto_csv(crudo):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return crudo.decode(enc)
        except UnicodeDecodeError:
            continue
    return crudo.decode("latin-1", errors="ignore")


def parecido(a, b):
    """0-1 entre dos nombres ya normalizados con s.nucleo()."""
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    corto, largo = (a, b) if len(a) <= len(b) else (b, a)
    if len(corto) >= 6 and re.search(r"\b%s\b" % re.escape(corto), largo):
        return 0.88
    return SequenceMatcher(None, a, b).ratio()


# ───────────────────────── 1. IDENTIDAD: DENUE ─────────────────────────
CELDA = 0.002  # ~220 m
# Actividades (SCIAN) que entran al directorio como registros propios: manejo de residuos (562) y
# comercio al por mayor de materiales de desecho (4343). Se excluyen las del sector público.
SCIAN_GIRO = ("562", "4343")
SCIAN_RESIDUO = {"434311": ["Metales / chatarra"], "434312": ["Cartón y papel"], "434313": ["Vidrio"], "434314": ["Plásticos"]}
ENTIDAD = {"coahuila de zaragoza": "Coahuila", "michoacan de ocampo": "Michoacán", "veracruz de ignacio de la llave": "Veracruz",
           "mexico": "Estado de México", "distrito federal": "Ciudad de México"}
_SEM_GEO = {}  # clave del grupo SEMARNAT -> fila DENUE (para ubicar en el mapa las empresas del padrón)


def estado_std(x):
    n = s.sin_acentos(x or "").strip().lower()
    if n in ENTIDAD:
        return ENTIDAD[n]
    return next((e for e in s.ESTADOS if s.sin_acentos(e).lower() == n), (x or "").strip())


def titulo(x):
    x = re.sub(r"\s+", " ", x or "").strip()
    return x.title() if x.isupper() else x


def denue_a_registro(f, previo):
    """Homologa un establecimiento del DENUE a los campos del directorio."""
    via = " ".join(x for x in (titulo(f["tipo_vial"]), titulo(f["nom_vial"]), f["numero_ext"] or f["letra_ext"]) if x)
    col = " ".join(x for x in (titulo(f["tipo_asent"]), titulo(f["nomb_asent"])) if x)
    cp = f["cod_postal"].zfill(5) if f["cod_postal"] else ""
    estado = estado_std(f["entidad"])
    direccion = ", ".join(x for x in (via, col, ("C.P. " + cp) if cp else "", titulo(f["municipio"]), estado) if x)
    texto = "%s %s" % (f["nombre_act"], f["nom_estab"])
    servicios = s.marcar(texto, s.SERVICIOS)
    if f["codigo_act"].startswith("4343"):
        servicios = sorted(set(servicios) | {"Compra de materiales", "Acopio"})
    elif not servicios:
        servicios = ["Manejo integral"]
    residuos = sorted(set(SCIAN_RESIDUO.get(f["codigo_act"], [])) | set(s.marcar(texto, s.RESIDUOS)))
    tel = s.limpiar_tel(f["telefono"])
    web = f["www"].strip().lower()
    if web and not web.startswith("http"):
        web = "http://" + web
    hoy = HOY()
    r = dict(previo or {})
    r.update({"es_empresa": True, "origen": "DENUE", "nombre": titulo(f["nom_estab"]), "direccion": direccion, "cp": cp,
              "estado_sede": estado, "estados_busqueda": [estado] if estado in s.ESTADOS else [],
              "telefonos": [tel] if tel else [], "correos": [f["correoelec"].strip().lower()] if "@" in f["correoelec"] else [],
              "sitio": web if "." in web else "", "servicios": servicios, "residuos": residuos,
              "denue_estatus": "Encontrado", "denue_coincidencia": "Registro del DENUE", "denue_id": f["id"],
              "denue_nombre": f["nom_estab"], "denue_razon_social": f["raz_social"], "denue_scian": f["codigo_act"],
              "denue_actividad": f["nombre_act"], "denue_personal": f["per_ocu"], "denue_alta": f["fecha_alta"],
              "denue_ubicacion": ", ".join(x for x in (f["municipio"], f["entidad"]) if x), "denue_verificado": hoy,
              "estatus_mes": "Sin cambios" if previo else "Nuevo", "primera_deteccion": (previo or {}).get("primera_deteccion", hoy),
              "ultima_verificacion": hoy})
    if f["raz_social"] and (not r.get("razon_social") or r.get("razon_social_fuente") == "DENUE"):
        r["razon_social"], r["razon_social_fuente"] = f["raz_social"], "DENUE"
    try:
        r["lat"], r["lng"], r["ubicacion_fuente"] = float(f["latitud"]), float(f["longitud"]), "DENUE"
    except ValueError:
        pass
    return r


def identidad(base):
    """Cruza con el DENUE las empresas no cruzadas en los últimos DENUE_DIAS días."""
    limite = (dt.date.today() - dt.timedelta(days=DENUE_DIAS)).isoformat()
    propios = [k for k in base if k.startswith("denue:")]
    importar = not propios or any((base[k].get("denue_verificado") or "") < limite for k in propios)
    pend = [k for k, r in base.items() if r.get("es_empresa") and not k.startswith("denue:") and
            ((r.get("denue_verificado") or "") < limite or (importar and r.get("denue_estatus") != "Encontrado")
             or (not r.get("lat") and r.get("razon_social") and r.get("denue_regla") != 2))]
    if not pend and not importar:
        return {"denue_pendientes": 0}
    # Empresas del padrón SEMARNAT: se buscan en el DENUE por teléfono o razón social para darles ubicación
    sem_nom, sem_tel = defaultdict(list), defaultdict(list)
    for g, filas in semarnat_grupos().items():
        sem_nom[g[0]].append(g)
        for x in filas:
            for tt in re.split(r"[;/,y]| ext", x["telefono"]):
                if tel10(tt):
                    sem_tel[tel10(tt)].append(g)
    _SEM_GEO.clear()
    giro = {}  # id DENUE -> fila completa (establecimientos del giro)
    por_rs = defaultdict(list)  # empresas sin ubicación: se buscan en el DENUE por razón social dentro del mismo estado
    for k in pend:
        r = base[k]
        n = s.nucleo(r.get("razon_social"))
        if not r.get("lat") and len(n) >= 10 and len(n.split()) >= 2:
            por_rs[n].append((k, sinac(r.get("estado_sede"))))
    por_tel, por_dom, por_mail, por_celda = defaultdict(list), defaultdict(list), defaultdict(list), defaultdict(list)
    nombres = {}
    for k in pend:
        r = base[k]
        nombres[k] = [n for n in (s.nucleo(r.get("nombre")), s.nucleo(r.get("razon_social"))) if n]
        for t in r.get("telefonos", []):
            if tel10(t):
                por_tel[tel10(t)].append(k)
        if ":" not in k:
            por_dom[k].append(k)
        for c in r.get("correos", []):
            por_mail[c.lower()].append(k)
            d = c.split("@")[-1].lower()
            if d not in CORREO_GENERICO:
                por_dom[d].append(k)
        if r.get("lat") and r.get("lng"):
            ci, cj = int(r["lat"] / CELDA), int(r["lng"] / CELDA)
            for a in (-1, 0, 1):
                for b in (-1, 0, 1):
                    por_celda[(ci + a, cj + b)].append(k)

    mejor = {}  # k -> (puntaje, motivo, fila)
    filas_total, sectores_ok = 0, 0
    for n_sec, sec in enumerate(DENUE_SECTORES):
        avisar("Cruzando con el DENUE (INEGI)", n_sec, len(DENUE_SECTORES))
        crudo = bajar("https://www.inegi.org.mx/contenidos/masiva/denue/denue_00_%s_csv.zip" % sec, t=600)
        if not crudo:
            continue
        with tempfile.TemporaryFile() as tmp:
            tmp.write(crudo)
            del crudo
            try:
                z = zipfile.ZipFile(tmp)
            except zipfile.BadZipFile:
                continue
            for info in z.infolist():
                if not info.filename.lower().endswith(".csv") or "diccionario" in info.filename.lower():
                    continue
                lector = csv.reader(io.TextIOWrapper(z.open(info), encoding="latin-1", newline=""))
                h = next(lector, None)
                if not h or "nom_estab" not in h:
                    continue
                ix = {c: h.index(c) for c in ("id", "nom_estab", "raz_social", "codigo_act", "nombre_act", "per_ocu",
                                              "cod_postal", "entidad", "municipio", "telefono", "correoelec", "www",
                                              "latitud", "longitud", "fecha_alta", "tipo_vial", "nom_vial", "numero_ext",
                                              "letra_ext", "tipo_asent", "nomb_asent")}
                i_cod, i_act, i_rs, i_tel = ix["codigo_act"], ix["nombre_act"], ix["raz_social"], ix["telefono"]
                sectores_ok += 1
                for f in lector:
                    filas_total += 1
                    if len(f) < len(h):
                        continue
                    hits = []
                    del_giro = f[i_cod].startswith(SCIAN_GIRO) and "blico" not in f[i_act]
                    if del_giro:
                        giro[f[ix["id"]]] = {c: f[i].strip() for c, i in ix.items()}
                    if sem_tel or sem_nom:
                        gs = sem_tel.get(tel10(f[i_tel])) if f[i_tel] else None
                        if not gs and f[i_rs]:
                            gs = [g for g in sem_nom.get(s.nucleo(f[i_rs]), ()) if not g[1] or sinac(estado_std(f[ix["entidad"]])) == g[1]]
                        for g in gs or ():
                            _SEM_GEO.setdefault(g, {c: f[i].strip() for c, i in ix.items()})
                    if por_rs and f[i_rs]:
                        cand = por_rs.get(s.nucleo(f[i_rs]))
                        if cand:
                            edo = sinac(estado_std(f[ix["entidad"]]))
                            hits += [(80, "Razón social", k) for k, e in cand if not e or e == edo]
                    t = f[ix["telefono"]]
                    if t and tel10(t) in por_tel:
                        hits += [(92, "Teléfono", k) for k in por_tel[tel10(t)]]
                    w = f[ix["www"]]
                    if w:
                        d = s.dominio_de(w.strip().lower())
                        if d in por_dom:
                            hits += [(92, "Sitio web", k) for k in por_dom[d]]
                    c = f[ix["correoelec"]]
                    if c:
                        c = c.strip().lower()
                        if c in por_mail:
                            hits += [(92, "Correo", k) for k in por_mail[c]]
                        else:
                            d = c.split("@")[-1]
                            if d in por_dom and d not in CORREO_GENERICO:
                                hits += [(84, "Dominio del correo", k) for k in por_dom[d]]
                    try:
                        la, lo = float(f[ix["latitud"]]), float(f[ix["longitud"]])
                    except ValueError:
                        la = lo = None
                    if la is not None:
                        cerca = por_celda.get((int(la / CELDA), int(lo / CELDA)))
                        if cerca:
                            nd = [n for n in (s.nucleo(f[ix["nom_estab"]]), s.nucleo(f[ix["raz_social"]])) if n]
                            for k in cerca:
                                r = base[k]
                                dist = math.hypot((la - r["lat"]) * 111320, (lo - r["lng"]) * 111320 * math.cos(math.radians(la)))
                                if dist > 180:
                                    continue
                                sim = max([parecido(a, b) for a in nd for b in nombres[k]] or [0])
                                if sim >= 0.72:
                                    hits.append((60 + sim * 25 - dist / 30, "Ubicación y nombre", k))
                                elif del_giro and (dist <= 25 or (dist <= 60 and sim >= 0.4)):
                                    hits.append((55 - dist / 30, "Ubicación y giro", k))  # un negocio del giro en el mismo punto
                    for punt, motivo, k in hits:
                        if k not in mejor or punt > mejor[k][0]:
                            mejor[k] = (punt, motivo, {c: f[i].strip() for c, i in ix.items()})
    avisar("Cruzando con el DENUE (INEGI)", len(DENUE_SECTORES), len(DENUE_SECTORES))
    if not sectores_ok:
        return {"denue_pendientes": len(pend), "denue_error": "no se pudo descargar el DENUE"}
    # un establecimiento del DENUE corresponde a una sola empresa: gana la coincidencia más fuerte
    dueno = {}
    for k, (punt, motivo, f) in sorted(mejor.items(), key=lambda kv: -kv[1][0]):
        if motivo in ("Ubicación y nombre", "Ubicación y giro") and f["id"] in dueno:
            del mejor[k]
        else:
            dueno.setdefault(f["id"], k)

    hallados = nuevas_rs = 0
    for k in pend:
        r = base[k]
        for c in [c for c in r if c.startswith("denue_")]:
            r.pop(c)
        r["denue_verificado"] = HOY()
        r["denue_regla"] = 2
        if k not in mejor:
            r["denue_estatus"] = "No encontrado"
            continue
        _, motivo, f = mejor[k]
        hallados += 1
        r.update({"denue_estatus": "Encontrado", "denue_coincidencia": motivo, "denue_id": f["id"],
                  "denue_nombre": f["nom_estab"], "denue_razon_social": f["raz_social"], "denue_scian": f["codigo_act"],
                  "denue_actividad": f["nombre_act"], "denue_personal": f["per_ocu"], "denue_alta": f["fecha_alta"],
                  "denue_ubicacion": ", ".join(x for x in (f["municipio"], f["entidad"]) if x)})
        rs = f["raz_social"]
        if rs and not r.get("razon_social") and len(s.nucleo(rs)) >= 5:
            r["razon_social"], r["razon_social_fuente"] = rs, "DENUE"
            nuevas_rs += 1
            if r.get("repse_estatus") in (None, "No encontrado", "Sin datos para buscar", "Coincidencia dudosa"):
                r.pop("repse_verificado", None)  # con razón social nueva se vuelve a consultar REPSE
        if not (r.get("lat") and r.get("lng")):
            try:
                r["lat"], r["lng"], r["ubicacion_fuente"] = float(f["latitud"]), float(f["longitud"]), "DENUE"
            except ValueError:
                pass
        if not r.get("cp") and f["cod_postal"]:
            r["cp"] = f["cod_postal"].zfill(5)
        if not r.get("estado_sede") and r.get("cp"):
            r["estado_sede"] = s.estado_por_cp(r["cp"])
    # Establecimientos del giro que no corresponden a ninguna empresa ya registrada -> registros nuevos
    usados = {r.get("denue_id") for k, r in base.items() if not k.startswith("denue:") and r.get("denue_id")}
    nuevos = 0
    if importar:
        for i, f in giro.items():
            if i in usados:
                base.pop("denue:" + i, None)  # ya está representado por otra fuente
                continue
            nuevos += "denue:" + i not in base
            base["denue:" + i] = denue_a_registro(f, base.get("denue:" + i))
    return {"denue_pendientes": len(pend), "denue_encontradas": hallados, "denue_razones_sociales": nuevas_rs,
            "denue_filas": filas_total, "denue_giro": len(giro), "denue_registros_nuevos": nuevos,
            "denue_registros": sum(1 for k in base if k.startswith("denue:")), "semarnat_ubicadas_en_denue": len(_SEM_GEO)}


# ───────────────────────── 2a. SAT ─────────────────────────
SAT_BASE = "http://omawww.sat.gob.mx/cifras_sat/Documents/"
# archivo -> (lista, descripción por omisión, gravedad)   gravedad: 3 descarte · 2 grave · 1 leve · 0 informativo
SAT_ARCHIVOS = {
    "Listado_Completo_69-B.csv": ("69-B", "", None),
    "ListadoGlobalDefinitivo.csv": ("69-B Bis", "Transmisión indebida de pérdidas fiscales (definitivo)", 3),
    "Firmes.csv": ("69", "Crédito fiscal firme", 2),
    "Exigibles.csv": ("69", "Crédito fiscal exigible", 2),
    "No%20localizados.csv": ("69", "Contribuyente no localizado", 2),
    "Sentencias.csv": ("69", "Sentencia condenatoria por delito fiscal", 2),
    "Cancelados.csv": ("69", "Crédito fiscal cancelado", 1),
    "Condonadosart146BCFF.csv": ("69", "Crédito condonado (concurso mercantil)", 1),
    "Condonadosart21CFF.csv": ("69", "Recargos condonados", 1),
    "CondonadosporDecreto.csv": ("69", "Crédito condonado por decreto", 1),
}
GRAVEDAD_69B = {"definitivo": 3, "presunto": 2, "desvirtuado": 0, "sentencia favorable": 0}


def sat_cruzar(base):
    emp = [(k, r) for k, r in base.items() if r.get("es_empresa")]
    por_rfc, por_nom = defaultdict(list), defaultdict(list)
    for k, r in emp:
        if r.get("rfc"):
            por_rfc[r["rfc"].upper().replace("-", "").replace(" ", "")].append(k)
        for campo in ("razon_social", "repse_razon_social"):
            if campo == "repse_razon_social" and not str(r.get("repse_coincidencia", "")).startswith(("Alta", "Media")):
                continue
            n = s.nucleo(r.get(campo))
            if len(n) >= 15 and len(n.split()) >= 3:  # nombres cortos se prestan a homónimos
                por_nom[n].append(k)
    hallazgos, cargados, fechas = defaultdict(list), [], {}
    for n_arch, (arch, (lista, desc, grav)) in enumerate(SAT_ARCHIVOS.items()):
        avisar("Cruzando con listas del SAT", n_arch, len(SAT_ARCHIVOS))
        crudo = bajar(SAT_BASE + arch)
        if not crudo:
            continue
        filas = list(csv.reader(io.StringIO(texto_csv(crudo))))
        del crudo
        ih = next((i for i, f in enumerate(filas[:8]) if any(c.strip().upper() == "RFC" for c in f)), None)
        if ih is None:
            continue
        m = re.search(r"actualizada al ([^;,]+)", " ".join(filas[0])) if ih else None
        fechas[arch] = m.group(1).strip() if m else ""
        cab = [c.strip().upper() for c in filas[ih]]
        i_rfc = cab.index("RFC")
        i_nom = next((i for i, c in enumerate(cab) if "RAZ" in c or "NOMBRE" in c or "CONTRIBUYENTE" in c), None)
        i_sit = next((i for i, c in enumerate(cab) if "SITUACI" in c or c == "SUPUESTO"), None)
        i_fec = next((i for i, c in enumerate(cab) if "FECHA" in c or "PUBLICACI" in c), None)
        i_tipo = next((i for i, c in enumerate(cab) if c == "TIPO PERSONA"), None)
        i_ent = next((i for i, c in enumerate(cab) if c.startswith("ENTIDAD")), None)
        cargados.append(arch)
        for f in filas[ih + 1:]:
            if len(f) <= i_rfc:
                continue
            rfc = f[i_rfc].strip().upper()
            claves = [(k, "RFC") for k in por_rfc.get(rfc, [])]
            if not claves and i_nom is not None and len(f) > i_nom and por_nom:
                if i_tipo is not None and len(f) > i_tipo and f[i_tipo].strip().upper().startswith("F"):
                    continue  # personas físicas: el nombre no basta para identificar
                ent = sinac(f[i_ent]) if i_ent is not None and len(f) > i_ent else ""
                ent = {"mexico": "estado de mexico"}.get(ent, ent)
                claves = [(k, "razón social") for k in por_nom.get(s.nucleo(f[i_nom]), [])
                          if not base[k].get("rfc") and (not ent or not base[k].get("estado_sede") or sinac(base[k]["estado_sede"]) == ent)]
            if not claves:
                continue
            sit = f[i_sit].strip() if i_sit is not None and len(f) > i_sit else ""
            g = GRAVEDAD_69B.get(sit.lower(), 2) if lista == "69-B" else grav
            for k, cruce in claves:
                hallazgos[k].append({"lista": lista, "situacion": sit.capitalize() if lista == "69-B" else desc,
                                     "gravedad": g, "cruce": cruce, "rfc_lista": rfc,
                                     "nombre_lista": f[i_nom].strip() if i_nom is not None and len(f) > i_nom else "",
                                     "fecha": f[i_fec].strip() if i_fec is not None and len(f) > i_fec else ""})
    avisar("Cruzando con listas del SAT", len(SAT_ARCHIVOS), len(SAT_ARCHIVOS))
    if not cargados:
        return {"sat_error": "no se pudo descargar ninguna lista del SAT"}
    con_hallazgo = 0
    for k, r in emp:
        hs = sorted(hallazgos.get(k, []), key=lambda h: -h["gravedad"])
        tiene_rfc = bool(r.get("rfc"))
        tiene_nombre = any(k in v for v in por_nom.values())
        r["sat_hallazgos"] = hs[:6]
        r["sat_cruce"] = "RFC" if tiene_rfc else "razón social" if tiene_nombre else ""
        r["sat_verificado"] = HOY()
        if hs:
            con_hallazgo += 1
            peor = hs[0]
            r["sat_estatus"] = "%s: %s" % (peor["lista"], peor["situacion"])
        elif tiene_rfc or tiene_nombre:
            r["sat_estatus"] = "Sin hallazgos"
        else:
            r["sat_estatus"] = "No verificable"
    return {"sat_listas": cargados, "sat_con_hallazgo": con_hallazgo, "sat_corte": fechas.get("Listado_Completo_69-B.csv", "")}


# ───────────────────────── 2b. SEMARNAT ─────────────────────────
SEMARNAT_BASE = "https://dsiappsdev.semarnat.gob.mx/datos/residuos/"
SEMARNAT_RUBROS = {
    "Rubro1_Reciclaje_de_residuos_peligrosos_industriales": "Reciclaje",
    "Rubro2_Aprovechamiento_de_residuos_peligrosos_industriales": "Aprovechamiento",
    "Rubro3_Coprocesamiento_de_residuos_peligrosos_industriales": "Co-procesamiento",
    "Rubro4_Reutilizacion_de_residuos_peligrosos_industriales": "Reutilización",
    "Rubro5_Tratamiento_de_residuos_peligrosos_industriales": "Tratamiento",
    "Rubro6_Incineracion_de_residuos_peligrosos_industriales": "Incineración",
    "Rubro7_Confinamiento_de_residuos_peligrosos_industriales": "Confinamiento",
    "Rubro8_Recoleccion_y_Transporte_de_residuos_peligrosos": "Recolección y transporte",
    "Rubro9_Almacenamiento_y_Acopio_de_residuos_peligrosos": "Almacenamiento y acopio",
    "Rubro10_Tratamiento_de_rpbi_in_situ": "Tratamiento RPBI in situ",
    "Rubro11_Tratamiento_de_rpbi_ex_situ": "Tratamiento RPBI ex situ",
    "Rubro12_Incineracion_de_residuos_peligrosos_biologico_infecciosos": "Incineración RPBI",
    "Rubro13_Manejo_de_BPCs": "Manejo de BPC",
    "Rubro14_Exportacion-Importacion_de_residuos_peligrosos_industriales": "Exportación / importación",
    "Rubro15_Remediacion_de_suelos_contaminados": "Remediación de suelos",
}


def fecha_iso(x):
    m = re.search(r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})", x or "")
    if not m:
        return ""
    d, mes = int(m.group(1)), int(m.group(2))
    if mes > 12:  # algunos renglones vienen como mes/día/año
        d, mes = mes, d
    return "%s-%02d-%02d" % (m.group(3), mes, d)


RUBRO_SERVICIOS = {"Reciclaje": ["Reciclaje"], "Aprovechamiento": ["Recuperación / valorización"], "Co-procesamiento": ["Co-procesamiento"],
                   "Reutilización": ["Recuperación / valorización"], "Tratamiento": ["Tratamiento"], "Incineración": ["Tratamiento"],
                   "Confinamiento": ["Disposición final"], "Recolección y transporte": ["Recolección", "Transporte"],
                   "Almacenamiento y acopio": ["Acopio", "Almacenamiento temporal"], "Tratamiento RPBI in situ": ["Tratamiento"],
                   "Tratamiento RPBI ex situ": ["Tratamiento"], "Incineración RPBI": ["Tratamiento"], "Manejo de BPC": ["Tratamiento"],
                   "Remediación de suelos": ["Remediación"]}
_SEM = {"filas": None, "rubros": []}


def semarnat_filas():
    """Descarga (una vez por corrida) los padrones de residuos peligrosos y los devuelve como lista de autorizaciones."""
    if _SEM["filas"] is not None:
        return _SEM["filas"]
    todas, cargados = [], []
    for n_r, (arch, rubro) in enumerate(SEMARNAT_RUBROS.items()):
        avisar("Descargando padrones SEMARNAT", n_r, len(SEMARNAT_RUBROS))
        crudo = bajar(SEMARNAT_BASE + arch + ".csv", t=120)
        if not crudo:
            continue
        filas = list(csv.reader(io.StringIO(texto_csv(crudo))))
        ih = next((i for i, f in enumerate(filas[:6]) if any(c.strip().lower() == "empresa" for c in f)), None)
        if ih is None:
            continue
        cab = [sinac(c) for c in filas[ih]]
        col = lambda *pal: next((i for i, c in enumerate(cab) if all(p in c for p in pal)), None)  # noqa: E731
        i_emp, i_res = col("empresa"), col("residuo") if col("residuo") is not None else col("autorizado para")
        i_tel = col("telefono") if col("telefono") is not None else col("contacto")
        i_dir = col("direccion") if col("direccion") is not None else col("domicilio")
        i_aut, i_fin = col("autorizacion"), col("termino")
        i_edo = next((i for i, c in enumerate(cab) if c == "estado"), None)
        i_mun = next((i for i, c in enumerate(cab) if c == "municipio"), None)
        cargados.append(rubro)
        for f in filas[ih + 1:]:
            if i_emp is None or len(f) <= i_emp or not f[i_emp].strip():
                continue
            g = lambda i: re.sub(r"\s+", " ", f[i]).strip() if i is not None and len(f) > i else ""  # noqa: E731
            todas.append({"rubro": rubro, "empresa": g(i_emp), "direccion": g(i_dir), "municipio": g(i_mun),
                          "estado": estado_std(g(i_edo)), "telefono": g(i_tel), "residuos": g(i_res),
                          "autorizacion": g(i_aut), "vigencia": fecha_iso(g(i_fin))})
    _SEM["filas"], _SEM["rubros"] = todas, cargados
    return todas


def semarnat_grupos():
    """Autorizaciones agrupadas por empresa: (razón social normalizada, estado) -> filas."""
    grupos = defaultdict(list)
    for x in semarnat_filas():
        n = s.nucleo(x["empresa"])
        if n:
            grupos[(n, sinac(x["estado"]))].append(x)
    return grupos


def semarnat_importar(base):
    """Crea un registro por cada empresa del padrón SEMARNAT que no esté ya en el directorio."""
    por_nom, por_tel = set(), set()
    for k, r in base.items():
        if not r.get("es_empresa") or k.startswith("semarnat:"):
            continue
        for campo in ("razon_social", "repse_razon_social", "denue_razon_social", "nombre"):
            n = s.nucleo(r.get(campo))
            if len(n) >= 8 and n not in s.GENERICAS:
                por_nom.add(n)
        por_tel.update(tel10(t) for t in r.get("telefonos", []) if tel10(t))
    hoy, nuevos, vigentes = HOY(), 0, set()
    for (n, edo), filas in semarnat_grupos().items():
        tels = [tel10(t) for x in filas for t in re.split(r"[;/,y]| ext", x["telefono"]) if tel10(t)]
        if n in por_nom or any(t in por_tel for t in tels):
            continue
        k = "semarnat:%s|%s" % (n, edo)
        vigentes.add(k)
        previo = base.get(k)
        x0 = filas[0]
        texto = " ".join(x["residuos"] for x in filas)
        residuos = {"Peligrosos (RP)"} | set(s.marcar(texto, s.RESIDUOS))
        if any("RPBI" in x["rubro"] for x in filas):
            residuos.add("Biológico-infecciosos (RPBI)")
        r = dict(previo or {})
        r.update({"es_empresa": True, "origen": "SEMARNAT", "nombre": x0["empresa"], "razon_social": x0["empresa"],
                  "razon_social_fuente": "SEMARNAT", "direccion": x0["direccion"], "estado_sede": x0["estado"],
                  "estados_busqueda": sorted({x["estado"] for x in filas if x["estado"] in s.ESTADOS}, key=s.ESTADOS.index),
                  "telefonos": list(dict.fromkeys(s.limpiar_tel(t) for t in tels if s.limpiar_tel(t)))[:5],
                  "servicios": sorted({sv for x in filas for sv in RUBRO_SERVICIOS.get(x["rubro"], [])}),
                  "residuos": sorted(residuos), "autorizaciones": ["SEMARNAT"],
                  "num_autorizacion": next((x["autorizacion"] for x in filas if x["autorizacion"]), ""),
                  "estatus_mes": "Sin cambios" if previo else "Nuevo", "primera_deteccion": (previo or {}).get("primera_deteccion", hoy),
                  "ultima_verificacion": hoy})
        r.setdefault("correos", [])
        cp = re.search(r"C\.?\s?P\.?\s*(\d{5})", x0["direccion"])
        if cp:
            r["cp"] = cp.group(1)
        f = _SEM_GEO.get((n, edo))
        if f:  # la misma empresa localizada en el DENUE: ubicación, tamaño y actividad
            try:
                r["lat"], r["lng"], r["ubicacion_fuente"] = float(f["latitud"]), float(f["longitud"]), "DENUE"
            except ValueError:
                pass
            r.update({"denue_estatus": "Encontrado", "denue_coincidencia": "Teléfono o razón social", "denue_id": f["id"],
                      "denue_nombre": f["nom_estab"], "denue_razon_social": f["raz_social"], "denue_scian": f["codigo_act"],
                      "denue_actividad": f["nombre_act"], "denue_personal": f["per_ocu"], "denue_alta": f["fecha_alta"],
                      "denue_verificado": hoy})
            if not r.get("cp") and f["cod_postal"]:
                r["cp"] = f["cod_postal"].zfill(5)
            if "@" in f["correoelec"] and not r["correos"]:
                r["correos"] = [f["correoelec"].lower()]
            if f["www"] and not r.get("sitio"):
                r["sitio"] = f["www"].lower() if f["www"].lower().startswith("http") else "http://" + f["www"].lower()
        elif not previo:
            r.update({"denue_estatus": "No encontrado", "denue_verificado": hoy})
        nuevos += previo is None
        base[k] = r
    for k in [k for k in base if k.startswith("semarnat:") and k not in vigentes]:
        base.pop(k)  # ya está representada por un registro de otra fuente
    return {"semarnat_registros_nuevos": nuevos, "semarnat_registros": len(vigentes),
            "semarnat_con_ubicacion": sum(1 for k in vigentes if base[k].get("lat"))}


def semarnat_cruzar(base):
    emp = [(k, r) for k, r in base.items() if r.get("es_empresa")]
    por_nom, por_tel = defaultdict(set), defaultdict(set)
    for k, r in emp:
        for campo in ("razon_social", "repse_razon_social", "denue_razon_social", "nombre"):
            if campo == "repse_razon_social" and not str(r.get("repse_coincidencia", "")).startswith(("Alta", "Media")):
                continue
            n = s.nucleo(r.get(campo))
            if len(n) >= 8 and n not in s.GENERICAS:
                por_nom[n].add(k)
        for t in r.get("telefonos", []):
            if tel10(t):
                por_tel[tel10(t)].add(k)
    filas = semarnat_filas()
    if not _SEM["rubros"]:
        return {"semarnat_error": "no se pudo descargar ningún padrón SEMARNAT"}
    aut = defaultdict(list)
    for x in filas:
        ks = set(por_nom.get(s.nucleo(x["empresa"]), ()))
        cruce = "razón social"
        if not ks:
            for t in re.split(r"[;/,y]| ext", x["telefono"]):
                if tel10(t) in por_tel:
                    ks |= por_tel[tel10(t)]
                    cruce = "teléfono"
        for k in ks:
            aut[k].append({"rubro": x["rubro"], "empresa": x["empresa"], "autorizacion": x["autorizacion"],
                           "vigencia": x["vigencia"], "estado": x["estado"], "municipio": x["municipio"],
                           "residuos": x["residuos"][:220], "cruce": cruce})
    hoy, con = HOY(), 0
    for k, r in emp:
        a = sorted(aut.get(k, []), key=lambda x: x["vigencia"], reverse=True)
        r["semarnat"] = a[:10]
        r["semarnat_rubros"] = sorted({x["rubro"] for x in a})
        r["semarnat_verificado"] = hoy
        if not a:
            r["semarnat_estatus"] = "Sin autorización federal localizada"
        else:
            con += 1
            r["semarnat_estatus"] = "Autorización vigente" if any(x["vigencia"] >= hoy for x in a) else \
                "Autorización vencida" if all(x["vigencia"] for x in a) else "Autorización (vigencia no publicada)"
    return {"semarnat_rubros": _SEM["rubros"], "semarnat_con_autorizacion": con}


# ───────────────────────── 2b-bis. TRANSPARENCIA (PNT) ─────────────────────────
# Directorio «Concesiones, licencias, permisos y autorizaciones» de la dependencia ambiental de cada estado
# (Plataforma Nacional de Transparencia). Los CSV se bajan a mano desde el navegador (el sitio bloquea
# programas) y se guardan en fuentes/transparencia/<año>/NN_XXX.csv, donde NN es la clave INEGI del estado.
ENT_INEGI = ["Aguascalientes", "Baja California", "Baja California Sur", "Campeche", "Coahuila", "Colima", "Chiapas",
             "Chihuahua", "Ciudad de México", "Durango", "Guanajuato", "Guerrero", "Hidalgo", "Jalisco", "Estado de México",
             "Michoacán", "Morelos", "Nayarit", "Nuevo León", "Oaxaca", "Puebla", "Querétaro", "Quintana Roo",
             "San Luis Potosí", "Sinaloa", "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"]
TR_RES = re.compile(r"residuo|manejo especial|recicl|acopio|chatarra|desecho|relleno sanitario|\brme\b|\brsu\b|llantas|escombro|aceites? usados?")
TR_SERV = re.compile(r"recolec|transport|acopio|almacenamiento|recicl|tratamiento|disposicion final|co-?proces|prestador|transferencia|"
                     r"valoriz|aprovechamiento|destino final|manejo integral|compra|comercializ|relleno sanitario|confinamiento|trituraci|compost")
TR_FUERTE = re.compile(r"recolec|transport|acopio|recicl|tratamiento|disposicion final|co-?proces|transferencia|relleno sanitario|confinamiento|prestador")
TR_GEN = re.compile(r"generador|plan de manejo|cedula de operacion")
TR_GEN_SIEMPRE = re.compile(r"llevar un manejo integral|que generaran|residuos que genera")
RE_TITULAR_NULO = re.compile(r"^(no aplica|ver nota|persona voluntaria|persona fisica|persona moral|no se |sin |n/?a\b|ningun)")
TR_VACIO = {"", "no aplica", "n/a", "na", "no dato", "sin dato", "ver nota", "persona fisica", "persona moral", "publica", "privada",
            "s/n", "no", "no se genero", "ninguno", "ninguna", "no disponible", "nd", "sin informacion", "no se cuenta", "-", "x"}


def fecha_any(x):
    x = (x or "").strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", x)
    return "%s-%s-%s" % m.groups() if m else fecha_iso(x)


def transparencia_filas():
    """Lee los CSV de Transparencia y devuelve (autorizaciones de servicio de residuos, resumen por estado)."""
    import glob
    csv.field_size_limit(10 ** 9)
    filas, resumen = [], {}
    for ruta in sorted(glob.glob(os.path.join(s.DATOS, "fuentes", "transparencia", "**", "*.csv"), recursive=True)):
        m = re.match(r"(\d{2})_", os.path.basename(ruta))
        if not m or not 1 <= int(m.group(1)) <= 32:
            continue
        estado = ENT_INEGI[int(m.group(1)) - 1]
        with open(ruta, "rb") as fh:
            lineas = list(csv.reader(io.StringIO(texto_csv(fh.read()))))
        if len(lineas) < 2:
            continue
        cab = [sinac(c) for c in lineas[0]]
        col = lambda pref: next((i for i, c in enumerate(cab) if c.startswith(pref)), None)  # noqa: E731
        ix = {"tipo": col("tipo de acto"), "numero": col("numero de control"), "objeto": col("objeto"), "area": col("unidad(es)"),
              "nom": col("nombre(s) de la persona fisica"), "ap1": col("primer apellido de la persona fisica"),
              "ap2": col("segundo apellido de la persona fisica"), "rs": col("razon social"), "vini": col("fecha de inicio de vigencia"),
              "vfin": col("fecha de termino de vigencia"), "doc": col("hipervinculo al contrato"),
              "pini": col("fecha de inicio del periodo"), "pfin": col("fecha de termino del periodo")}
        cuenta = {"filas": len(lineas) - 1, "residuos": 0, "servicio": 0, "generadores": 0, "sin_titular": 0}
        for f in lineas[1:]:
            g = lambda c: re.sub(r"\s+", " ", f[ix[c]]).strip() if ix[c] is not None and len(f) > ix[c] else ""  # noqa: E731
            txt = sinac(g("objeto") + " " + g("numero"))
            if not TR_RES.search(txt):
                continue
            cuenta["residuos"] += 1
            if TR_GEN_SIEMPRE.search(txt) or (TR_GEN.search(txt) and not TR_FUERTE.search(txt)):
                cuenta["generadores"] += 1  # quien genera residuos, no quien presta el servicio
                continue
            if not TR_SERV.search(txt) or sinac(g("tipo")).startswith("convenio"):
                continue
            rs = g("rs")
            nulo = lambda x: sinac(x).strip(" .") in TR_VACIO or bool(RE_TITULAR_NULO.match(sinac(x)))  # noqa: E731
            moral = not nulo(rs)
            titular = rs if moral else ""
            if not titular and not nulo(g("nom")):
                titular = " ".join(x for x in (g("nom"), g("ap1"), g("ap2")) if not nulo(x))
            if len(s.nucleo(titular)) < 5:
                cuenta["sin_titular"] += 1
                continue
            cuenta["servicio"] += 1
            filas.append({"estado": estado, "titular": titular, "moral": moral, "tipo": g("tipo"), "numero": g("numero"),
                          "objeto": g("objeto")[:260], "area": g("area")[:90], "vigencia_inicio": fecha_any(g("vini")),
                          "vigencia": fecha_any(g("vfin")), "periodo": "%s a %s" % (fecha_any(g("pini")), fecha_any(g("pfin"))),
                          "documento": g("doc") if g("doc").lower().startswith("http") else ""})
        resumen[estado] = cuenta
    return filas, resumen


def transparencia_importar(base):
    """Agrega las autorizaciones estatales a las empresas ya registradas y da de alta las que no estaban."""
    filas, resumen = transparencia_filas()
    for r in base.values():  # se recalcula completo en cada corrida
        r.pop("transparencia", None)
        r.pop("transparencia_estatus", None)
    if not filas:
        for k in [k for k in base if k.startswith("transp:")]:
            base.pop(k)
        return {"transparencia_estados": 0}
    por_nom = defaultdict(set)
    for k, r in base.items():
        if not r.get("es_empresa") or k.startswith("transp:"):
            continue
        for campo in ("razon_social", "repse_razon_social", "denue_razon_social", "nombre"):
            n = s.nucleo(r.get(campo))
            if len(n) >= 8 and n not in s.GENERICAS:
                por_nom[n].add(k)
    grupos = defaultdict(list)
    for x in filas:
        grupos[(s.nucleo(x["titular"]), x["estado"])].append(x)
    hoy, nuevos, vigentes, ligadas = HOY(), 0, set(), set()

    def compacta(xs):
        vistos, sal = set(), []
        for x in sorted(xs, key=lambda x: x["vigencia"] or x["vigencia_inicio"], reverse=True):
            clave = (x["numero"], x["objeto"][:60])
            if clave in vistos:
                continue
            vistos.add(clave)
            sal.append({c: x[c] for c in ("estado", "tipo", "numero", "objeto", "area", "vigencia_inicio", "vigencia", "periodo", "documento")})
        return sal[:12]

    def estatus(a):
        if any(x["vigencia"] and x["vigencia"] >= hoy for x in a):
            return "Autorización estatal vigente"
        if all(x["vigencia"] for x in a):
            return "Autorización estatal vencida"
        return "Autorización estatal (sin fecha de término)"

    for (n, estado), xs in grupos.items():
        ks = por_nom.get(n)
        if ks:
            for k in ks:
                base[k]["transparencia"] = (base[k].get("transparencia") or []) + compacta(xs)
                ligadas.add(k)
            continue
        k = "transp:%s|%s" % (n, sinac(estado))
        vigentes.add(k)
        previo = base.get(k)
        x0 = xs[0]
        texto = " ".join(x["objeto"] + " " + x["numero"] for x in xs)
        residuos = set(s.marcar(texto, s.RESIDUOS))
        r = dict(previo or {})
        r.update({"es_empresa": True, "origen": "Transparencia", "nombre": titulo(x0["titular"]), "estado_sede": estado,
                  "estados_busqueda": [estado], "servicios": s.marcar(texto, s.SERVICIOS) or ["Manejo integral"],
                  "residuos": sorted(residuos), "autorizaciones": ["Autoridad ambiental estatal"],
                  "num_autorizacion": next((x["numero"] for x in xs if x["numero"]), ""),
                  "estatus_mes": "Sin cambios" if previo else "Nuevo", "primera_deteccion": (previo or {}).get("primera_deteccion", hoy),
                  "ultima_verificacion": hoy, "transparencia": compacta(xs)})
        if x0["moral"]:
            r["razon_social"], r["razon_social_fuente"] = x0["titular"], "Transparencia"
        for c in ("telefonos", "correos"):
            r.setdefault(c, [])
        nuevos += previo is None
        base[k] = r
    for k in [k for k in base if k.startswith("transp:") and k not in vigentes]:
        base.pop(k)
    for k in ligadas | vigentes:
        base[k]["transparencia"] = base[k]["transparencia"][:12]
        base[k]["transparencia_estatus"] = estatus(base[k]["transparencia"])
    return {"transparencia_estados": len(resumen), "transparencia_autorizaciones": len(filas),
            "transparencia_generadores_omitidos": sum(c["generadores"] for c in resumen.values()),
            "transparencia_registros_nuevos": nuevos, "transparencia_registros": len(vigentes),
            "transparencia_ligadas_a_existentes": len(ligadas), "transparencia_por_estado": resumen}


FUENTE_MOTOR = {"places": "Google Places", "gemini": "Buscador web", "importar": "Buscador web", "cse": "Buscador web"}


def marcar_fuentes(k, r):
    """`origen`: de dónde salió el registro. `fuentes`: todas las fuentes donde aparece."""
    motores = [FUENTE_MOTOR.get(m.strip(), m.strip()) for m in str(r.get("fuente") or "").split(",") if m.strip()]
    fs = list(dict.fromkeys(motores))
    if r.get("denue_estatus") == "Encontrado":
        fs.append("DENUE")
    if r.get("semarnat"):
        fs.append("SEMARNAT")
    if r.get("transparencia"):
        fs.append("Transparencia")
    r["fuentes"] = fs
    r["origen"] = "DENUE" if k.startswith("denue:") else "SEMARNAT" if k.startswith("semarnat:") else \
        "Transparencia" if k.startswith("transp:") else \
        "Google Places" if "Google Places" in motores else "Buscador web"


def sinac(x):
    return s.sin_acentos(x or "").strip().lower()


# ───────────────────────── 2c. Señales del sitio ─────────────────────────
def dominio_creado(dominio):
    """Fecha de registro del dominio (AAAA-MM-DD) o ''. .mx -> whois.mx; el resto -> rdap.org."""
    registrable = ".".join(dominio.split(".")[-3:]) if re.search(r"\.(com|org|net|gob|edu)\.mx$", dominio) else \
        ".".join(dominio.split(".")[-2:])
    try:
        if registrable.endswith(".mx"):
            c = socket.create_connection(("whois.mx", 43), timeout=15)
            c.sendall((registrable + "\r\n").encode())
            b = b""
            while True:
                x = c.recv(4096)
                if not x:
                    break
                b += x
            c.close()
            m = re.search(r"Created On:\s*(\d{4}-\d{2}-\d{2})", b.decode("latin-1"))
            return m.group(1) if m else ""
        r = requests.get("https://rdap.org/domain/" + registrable, headers=UA, timeout=20)
        if r.status_code == 404:
            return ""
        if r.status_code != 200:
            return None  # límite de consultas u otro error: se reintenta
        for e in r.json().get("events", []):
            if e.get("eventAction") == "registration":
                return e.get("eventDate", "")[:10]
    except Exception:
        return None  # sin respuesta: se reintenta en otra corrida
    return ""


def dominios(base, hilos=3):
    pend = [k for k, r in base.items() if r.get("es_empresa") and ":" not in k and not r.get("dominio_creado") and r.get("dominio_intentos", 0) < 3
            and not re.search(r"(ueniweb|wixsite|blogspot|github\.io|negocio\.site|webnode|weebly|godaddysites)", k)]
    if not pend:
        return 0
    avisar("Antigüedad de dominios", 0, len(pend))
    hechas = 0
    with ThreadPoolExecutor(max_workers=hilos) as ex:
        fut = {ex.submit(dominio_creado, k): k for k in pend}
        for fu in as_completed(fut):
            hechas += 1
            v = fu.result()
            base[fut[fu]]["dominio_intentos"] = base[fut[fu]].get("dominio_intentos", 0) + 1
            if v:
                base[fut[fu]]["dominio_creado"] = v
            if hechas % 20 == 0 or hechas == len(pend):
                avisar("Antigüedad de dominios", hechas, len(pend))
    return hechas


# ───────────────────────── 2d. Calificación ─────────────────────────
FUENTES = {"fiscal": "SAT · listados art. 69, 69-B y 69-B Bis", "laboral": "STPS · padrón REPSE",
           "ambiental": "SEMARNAT (residuos peligrosos) y Transparencia (autorizaciones estatales)", "existencia": "INEGI · DENUE",
           "reputacion": "Google Maps y sitio web"}


def calificar(k, r):
    hoy = dt.date.today()
    bloques, alertas, descarte = [], [], False

    # Fiscal (25)
    hs = r.get("sat_hallazgos") or []
    peor = max([h["gravedad"] for h in hs] or [0])
    if not r.get("sat_estatus") or r.get("sat_estatus") == "No verificable":
        p, nota, v = 12, "No se pudo verificar: falta RFC o razón social", False
    elif peor == 3:
        p, nota, v, descarte = 0, r["sat_estatus"], True, True
    elif peor == 2:
        p, nota, v = 0 if hs[0]["lista"] == "69-B" else 5, r["sat_estatus"], True
    elif peor == 1:
        p, nota, v = 15, r["sat_estatus"], True
    else:
        limpio_antes = [h for h in hs if h["gravedad"] == 0]
        p, v = (25, True) if r.get("sat_cruce") == "RFC" else (20, True)
        nota = "Sin hallazgos en las listas (cruce por %s)" % r.get("sat_cruce")
        if limpio_antes:
            nota += "; antecedente 69-B %s" % limpio_antes[0]["situacion"].lower()
    for h in hs:
        if h["gravedad"] >= 1:
            alertas.append("SAT %s: %s%s" % (h["lista"], h["situacion"], " (por nombre, verificar)" if h["cruce"] != "RFC" else ""))
    bloques.append({"b": "Fiscal", "p": p, "m": 25, "nota": nota, "fuente": FUENTES["fiscal"], "fecha": r.get("sat_verificado", ""), "v": v})

    # Laboral (15)
    e = r.get("repse_estatus", "")
    if e == "Registrado vigente":
        p, nota, v = (15 if str(r.get("repse_coincidencia", "")).startswith("Alta") else 12), \
            "Registro vigente hasta %s (coincidencia %s)" % (r.get("repse_vigencia", ""), str(r.get("repse_coincidencia", "")).lower()), True
    elif e in ("Registro vencido", "Sin registro activo"):
        p, nota, v = 4, e, True
        alertas.append("REPSE: " + e.lower())
    elif e == "Coincidencia dudosa":
        p, nota, v = 7, "Hay un registro con el mismo nombre en otro estado", False
    else:
        p, nota, v = 6, "Sin registro localizado (no todas las empresas están obligadas)", False
    bloques.append({"b": "Laboral", "p": p, "m": 15, "nota": nota, "fuente": FUENTES["laboral"], "fecha": r.get("repse_verificado", ""), "v": v})

    # Ambiental (20)
    e = r.get("semarnat_estatus", "")
    te = r.get("transparencia_estatus", "")
    if e == "Autorización vigente":
        p, nota, v = 20, "Autorización federal vigente: " + ", ".join(r.get("semarnat_rubros", [])[:4]) + \
            ("; además autorización estatal" if te else ""), True
    elif te:
        edos = ", ".join(sorted({x["estado"] for x in r.get("transparencia", [])}))
        p, nota, v = (17 if te == "Autorización estatal vigente" else 12), "%s (%s), publicada en Transparencia" % (te, edos), True
        if te == "Autorización estatal vencida":
            alertas.append("Autorización estatal vencida")
    elif e.startswith("Autorización"):
        p, nota, v = 11, e + ": " + ", ".join(r.get("semarnat_rubros", [])[:4]), True
        if e == "Autorización vencida":
            alertas.append("SEMARNAT: autorización vencida")
    elif r.get("autorizaciones") or r.get("num_autorizacion"):
        p, nota, v = 8, "El sitio menciona autorizaciones (%s); no localizada en el padrón federal" % ", ".join(r.get("autorizaciones", [])[:3]), False
    else:
        p, nota, v = 4, "Sin autorización federal localizada (el manejo especial se autoriza por estado)", False
    bloques.append({"b": "Ambiental", "p": p, "m": 20, "nota": nota, "fuente": FUENTES["ambiental"], "fecha": r.get("semarnat_verificado", ""), "v": v})

    # Existencia (20)
    p, partes = 0, []
    if r.get("denue_estatus") == "Encontrado":
        p += 8
        partes.append("establecimiento en DENUE (%s%s)" % (r.get("denue_personal", ""), ", alta " + r["denue_alta"] if r.get("denue_alta") else ""))
    if r.get("razon_social"):
        p += 4
        partes.append("razón social")
    if r.get("rfc"):
        p += 4
        partes.append("RFC")
    if r.get("direccion"):
        p += 2
        partes.append("dirección")
    if r.get("telefonos"):
        p += 2
        partes.append("teléfono")
    bloques.append({"b": "Existencia", "p": p, "m": 20, "nota": "Con " + ", ".join(partes) if partes else "Sin datos de identidad",
                    "fuente": FUENTES["existencia"], "fecha": r.get("denue_verificado", ""), "v": r.get("denue_estatus") == "Encontrado"})

    # Reputación (20)
    p, partes = 0, []
    try:
        cal, res = float(r.get("calificacion") or 0), int(r.get("resenas") or 0)
    except (TypeError, ValueError):
        cal, res = 0, 0
    if cal:
        p += 8 if cal >= 4.5 and res >= 10 else 6 if cal >= 4 else 3 if cal >= 3 else 0
        partes.append("Google %.1f★%s" % (cal, " (%d reseñas)" % res if res else ""))
        if cal < 3 and res >= 5:
            alertas.append("Calificación baja en Google (%.1f)" % cal)
    else:
        p += 3
    if ":" not in k:
        p += 3
        partes.append("sitio propio")
        if str(r.get("url_final", "")).startswith("https://"):
            p += 1
        if any(c.split("@")[-1].lower() not in CORREO_GENERICO for c in r.get("correos", [])):
            p += 3
            partes.append("correo con dominio propio")
        if r.get("dominio_creado"):
            anios = (hoy - dt.date.fromisoformat(r["dominio_creado"])).days / 365.25
            p += 3 if anios >= 5 else 2 if anios >= 2 else 0
            partes.append("dominio de %d años" % anios)
    if len([x for x in (r.get("redes") or {}) if x != "WhatsApp"]) >= 1:
        p += 2
        partes.append("redes sociales")
    bloques.append({"b": "Reputación", "p": min(p, 20), "m": 20, "nota": ", ".join(partes) or "Sin señales públicas",
                    "fuente": FUENTES["reputacion"], "fecha": r.get("ultima_verificacion", ""), "v": bool(cal)})

    total = sum(b["p"] for b in bloques)
    verificadas = sum(1 for b in bloques[:4] if b["v"])
    # existir en el DENUE no dice nada sobre cumplimiento: sin dato fiscal, laboral ni ambiental no se califica
    nivel = "No confiable" if descarte else "Sin información suficiente" if not any(b["v"] for b in bloques[:3]) and not alertas else \
        "Alta" if total >= 75 else "Media" if total >= 55 else "Baja"
    r.update({"confianza": 0 if descarte else total, "confianza_nivel": nivel, "confianza_bloques": bloques,
              "confianza_alertas": alertas, "confianza_verificadas": verificadas,
              "confianza_fecha": hoy.isoformat()})


def padrones(base):
    """Da de alta los registros de SEMARNAT y Transparencia (antes de REPSE, para que también se validen)."""
    res = {}
    for paso in (semarnat_importar, transparencia_importar):
        try:
            res.update(paso(base))
        except Exception as e:
            res[paso.__name__ + "_error"] = str(e)[:160]
    return res


def perfilar(base):
    res = {}
    for paso in (sat_cruzar, semarnat_cruzar):
        try:
            res.update(paso(base))
        except Exception as e:  # una fuente caída no detiene la corrida
            res[paso.__name__ + "_error"] = str(e)[:160]
    try:
        res["dominios_consultados"] = dominios(base)
    except Exception as e:
        res["dominios_error"] = str(e)[:160]
    avisar("Calculando calificación de confianza")
    for k, r in base.items():
        if r.get("es_empresa"):
            marcar_fuentes(k, r)
            calificar(k, r)
    return res
