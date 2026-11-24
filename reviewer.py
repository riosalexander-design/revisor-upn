import os
import io
import fitz  # PyMuPDF
import docx
import re
import json
import time
from google import genai
from google.genai import types

import unicodedata
import pandas as pd
import requests

_STOPWORDS_BIB = {
    "de", "la", "el", "los", "las", "en", "del", "al", "un", "una", "unos", "unas",
    "por", "para", "con", "sin", "sobre", "entre", "hacia", "durante", "segun", "como",
    "mas", "sus", "este", "esta", "estos", "estas", "estudio", "investigacion", "nivel",
    "the", "of", "and", "in", "for", "with", "on", "to", "from", "among", "between",
    "during", "towards", "toward", "study", "research", "associated", "factors", "level",
    "disponible", "fuente", "fuentes", "verificada", "verificadas", "agregar", "bibliografia",
    "vancouver", "verifica", "lee", "corrobora", "personalmente", "enlace", "documento"
}

def _norm_ascii(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"[^a-z0-9\s]", " ", s)

def _content_words(s):
    return [w for w in _norm_ascii(s).split() if len(w) >= 4 and w not in _STOPWORDS_BIB and not w.isdigit()]

def _format_crossref_item(m):
    raw_doi = (m.get("DOI") or "").strip()
    title = " ".join(m.get("title") or []).strip().rstrip(".")
    authors_raw = m.get("author") or []
    surnames = []
    van_auths = []
    for a in authors_raw[:6]:
        fam = (a.get("family") or "").strip()
        giv = (a.get("given") or "").strip()
        if fam:
            for part_f in re.split(r"[\s\-]+", fam):
                if len(part_f) >= 3:
                    surnames.append(_norm_ascii(part_f).strip())
            inits = "".join(p[0].upper() for p in giv.split() if p and p[0].isalpha())
            van_auths.append(f"{fam} {inits}".strip())
        elif a.get("name"):
            inst_n = a.get("name").strip()
            van_auths.append(inst_n)
            surnames.extend(w for w in _norm_ascii(inst_n).split() if len(w) >= 4)
    if len(authors_raw) > 6:
        van_auths.append("et al")
    for a in authors_raw[6:]:
        fam = (a.get("family") or "").strip()
        for part_f in re.split(r"[\s\-]+", fam):
            if len(part_f) >= 3:
                surnames.append(_norm_ascii(part_f).strip())

    journal = " ".join(m.get("container-title") or m.get("short-container-title") or []).strip() or (m.get("publisher") or "").strip()
    year = ""
    for k in ("published-print", "published-online", "issued", "created"):
        dp = ((m.get(k) or {}).get("date-parts") or [[]])[0]
        if dp and dp[0]:
            year = str(dp[0])
            break
    vol = str(m.get("volume") or "").strip()
    iss = str(m.get("issue") or "").strip()
    page = str(m.get("page") or m.get("article-number") or "").strip()
    vol_p = f"{vol}({iss})" if (vol and iss) else vol
    loc_p = f"{year};{vol_p}:{page}" if (vol_p and page) else (f"{year};{vol_p}" if vol_p else year)
    auth_str = ", ".join(van_auths) or "Autor institucional"
    j_str = f" {journal}." if journal else ""
    l_str = f" {loc_p}." if loc_p else ""
    van = f"{auth_str}. {title}.{j_str}{l_str} Disponible en: https://doi.org/{raw_doi}"
    return {
        "title": title,
        "surnames": [s for s in surnames if s],
        "year": year,
        "journal": journal,
        "doi": raw_doi,
        "vancouver": re.sub(r"\s+", " ", van).strip()
    }

def _format_openalex_item(w):
    doi_url = (w.get("doi") or "").strip()
    raw_doi = re.sub(r"^https?://doi\.org/", "", doi_url, flags=re.IGNORECASE).strip()
    title = (w.get("title") or "").strip().rstrip(".")
    auth_list = w.get("authorships") or []
    surnames = []
    van_auths = []
    for idx, a in enumerate(auth_list):
        dn = ((a.get("author") or {}).get("display_name") or "").strip()
        if not dn:
            continue
        pts = [p for p in re.split(r"\s+", dn) if p]
        for p in pts:
            for sub_p in re.split(r"[\-.]+", p):
                if len(sub_p) >= 3:
                    surnames.append(_norm_ascii(sub_p).strip())
        if idx < 6:
            if len(pts) >= 2:
                # Si tiene 3+ palabras (típico hispano: Nombre Apellido1 Apellido2), conservar apellidos
                if len(pts) >= 3:
                    fam_part = f"{pts[-2]} {pts[-1]}"
                    init_part = "".join(x[0].upper() for x in pts[:-2] if x[0].isalpha())
                else:
                    fam_part = pts[-1]
                    init_part = "".join(x[0].upper() for x in pts[:-1] if x[0].isalpha())
                van_auths.append(f"{fam_part} {init_part}".strip())
            else:
                van_auths.append(dn)
    if len(auth_list) > 6:
        van_auths.append("et al")

    journal = (((w.get("primary_location") or {}).get("source") or {}).get("display_name") or "").strip()
    year = str(w.get("publication_year") or "").strip()
    bib = w.get("biblio") or {}
    vol = str(bib.get("volume") or "").strip()
    iss = str(bib.get("issue") or "").strip()
    fp = str(bib.get("first_page") or "").strip()
    lp = str(bib.get("last_page") or "").strip()
    page = f"{fp}-{lp}" if (fp and lp and fp != lp) else fp
    vol_p = f"{vol}({iss})" if (vol and iss) else vol
    loc_p = f"{year};{vol_p}:{page}" if (vol_p and page) else (f"{year};{vol_p}" if vol_p else year)
    auth_str = ", ".join(van_auths) or "Autor institucional"
    j_str = f" {journal}." if journal else ""
    l_str = f" {loc_p}." if loc_p else ""
    van = f"{auth_str}. {title}.{j_str}{l_str} Disponible en: https://doi.org/{raw_doi}"

    inv = w.get("abstract_inverted_index") or {}
    abs_txt = ""
    if inv:
        pos_map = {}
        for word, positions in inv.items():
            for p in positions:
                pos_map[p] = word
        abs_words = [pos_map[k] for k in sorted(pos_map.keys())[:45]]
        abs_txt = " ".join(abs_words) + "..."

    return {
        "title": title,
        "surnames": [s for s in surnames if s],
        "year": year,
        "journal": journal,
        "doi": raw_doi,
        "vancouver": re.sub(r"\s+", " ", van).strip(),
        "abstract_snippet": abs_txt
    }

_DOI_META_CACHE = {}

def _fetch_doi_metadata(raw_doi):
    key = raw_doi.lower().strip()
    if key in _DOI_META_CACHE:
        return _DOI_META_CACHE[key]
    meta = None
    try:
        r = requests.get(f"https://api.crossref.org/works/{raw_doi}", timeout=4.0)
        if r.status_code == 200:
            m = r.json().get("message", {})
            if m.get("title"):
                meta = _format_crossref_item(m)
    except Exception:
        pass
    if not meta:
        try:
            r = requests.get(f"https://api.openalex.org/works/https://doi.org/{raw_doi}", timeout=4.0)
            if r.status_code == 200:
                w = r.json()
                if w.get("title"):
                    meta = _format_openalex_item(w)
        except Exception:
            pass
    _DOI_META_CACHE[key] = meta
    return meta

def _check_title_and_author_match(context_text, meta, strict_title=True):
    """
    Corrobora simultáneamente coincidencia de AUTORES y TÍTULO entre lo escrito y el registro oficial del DOI.
    """
    if not meta or not meta.get("title"):
        return False
    ctx_norm = f" {_norm_ascii(context_text)} "
    surnames = [sn for sn in (meta.get("surnames") or []) if len(sn) >= 3]
    matched_authors = sum(1 for sn in set(surnames) if f" {sn} " in ctx_norm)
    if matched_authors == 0:
        return False

    # Si no se exige título en el mismo idioma (ej. párrafo en español que cita 'Janz y Becker (1984)' + DOI suelto),
    # basta con que coincida el autor + el año de publicación, o 2 apellidos de autores, o 1 palabra del título
    if not strict_title:
        yr = (meta.get("year") or "").strip()
        if yr and yr in context_text:
            return True
        if matched_authors >= 2:
            return True

    t_words = _content_words(meta["title"])
    if not t_words:
        return matched_authors >= 1
    ctx_words = set(_content_words(context_text))
    matched = sum(1 for w in set(t_words) if w in ctx_words)
    ratio = matched / max(1, len(set(t_words)))

    if strict_title:
        return matched >= 2 and ratio >= 0.42
    return matched >= 1 or ratio >= 0.20

def _search_verified_work_strict(query_txt):
    """
    Busca en Crossref y OpenAlex exigiendo MATCH ESTRICTO SIMULTÁNEO de TÍTULO + AUTORES antes de aceptarlo.
    """
    clean_q = re.sub(r"Fuente[s]?\s+verificada[s]?.*?:", "", query_txt, flags=re.IGNORECASE)
    clean_q = re.sub(r"\(Verifica.*?\)", "", clean_q, flags=re.IGNORECASE)
    clean_q = re.sub(r"(?:Disponible en:\s*)?https?://\S+", "", clean_q, flags=re.IGNORECASE).strip()
    if len(clean_q) < 20:
        return None

    # 1. Buscar en Crossref (top 5)
    try:
        r = requests.get(
            "https://api.crossref.org/works",
            params={"query.bibliographic": clean_q[:240], "rows": 5},
            timeout=4.0
        )
        if r.status_code == 200:
            for item in r.json().get("message", {}).get("items", []):
                if item.get("DOI") and item.get("title"):
                    cand = _format_crossref_item(item)
                    if _check_title_and_author_match(clean_q, cand, strict_title=True):
                        return cand
    except Exception:
        pass

    # 2. Buscar en OpenAlex (top 5, ideal para revistas peruanas/latinoamericanas OJS/SciELO)
    try:
        r = requests.get(
            "https://api.openalex.org/works",
            params={"search": clean_q[:200], "filter": "has_doi:true", "per_page": 5},
            timeout=4.0
        )
        if r.status_code == 200:
            for w in r.json().get("results", []):
                if w.get("doi") and w.get("title"):
                    cand = _format_openalex_item(w)
                    if _check_title_and_author_match(clean_q, cand, strict_title=True):
                        return cand
    except Exception:
        pass

    return None

def _fetch_verified_literature_bank(document_text, client):
    """
    Pre-busca artículos reales e indexados en OpenAlex (Internacionales, Nacionales-Perú y Teorías)
    con Título, Autores, Revista y DOI ya corroborados, para que la IA proponga estudios 100% existentes.
    """
    snippet = (document_text or "")[:3500]
    es_q = ""
    en_q = ""
    th_q = ""
    try:
        q_resp = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=(
                "Extrae del siguiente inicio de tesis 3 consultas cortas de búsqueda académica (máximo 6 palabras clave cada una, sin comillas ni operadores booleanos) "
                "y devuélvelas SOLO en JSON con llaves es_query (variables principales en español), en_query (variables principales en inglés) y theory_query (teoría o modelo clásico aplicable en inglés):\n\n"
                + snippet
            ),
            config=types.GenerateContentConfig(temperature=0.0, max_output_tokens=1024)
        )
        jm = re.search(r"\{.*?\}", q_resp.text or "", re.DOTALL)
        if jm:
            qd = json.loads(jm.group(0))
            es_q = (qd.get("es_query") or "").strip()
            en_q = (qd.get("en_query") or "").strip()
            th_q = (qd.get("theory_query") or "").strip()
    except Exception:
        pass

    if not es_q:
        words = _content_words(snippet[:800])
        es_q = " ".join(words[:6])
        en_q = es_q

    sections = []

    # A. Antecedentes Internacionales verificados (2020-2026, excluyendo Perú y repositorios genéricos Zenodo)
    if en_q:
        try:
            r_int = requests.get(
                "https://api.openalex.org/works",
                params={
                    "search": en_q,
                    "filter": "has_doi:true,publication_year:2020-2026,type:article",
                    "per_page": 10
                },
                timeout=5.0
            )
            if r_int.status_code == 200:
                int_items = []
                for w in r_int.json().get("results", []):
                    countries = []
                    for au in w.get("authorships") or []:
                        countries.extend(au.get("countries") or [])
                    if "PE" in countries:
                        continue
                    c = _format_openalex_item(w)
                    if c["title"] and c["surnames"] and c["doi"] and "zenodo" not in c["doi"].lower():
                        country_lbl = countries[0] if countries else "Internacional"
                        ab = f" | Hallazgos/Resumen: {c['abstract_snippet']}" if c.get("abstract_snippet") else ""
                        int_items.append((c["year"], f"- [{country_lbl}, {c['year']}] {c['vancouver']}{ab}"))
                int_items.sort(key=lambda x: x[0])
                if int_items:
                    sections.append("ANTECEDENTES INTERNACIONALES REALES VERIFICADOS (TÍTULO + AUTORES + DOI COMPROBADOS):\n" + "\n".join(x[1] for x in int_items[:6]))
        except Exception:
            pass

    # B. Antecedentes Nacionales (Perú) verificados (2019-2026)
    if es_q:
        try:
            r_nac = requests.get(
                "https://api.openalex.org/works",
                params={
                    "search": es_q,
                    "filter": "has_doi:true,publication_year:2019-2026,authorships.countries:countries/pe",
                    "per_page": 10
                },
                timeout=5.0
            )
            if r_nac.status_code == 200:
                nac_items = []
                for w in r_nac.json().get("results", []):
                    c = _format_openalex_item(w)
                    if c["title"] and c["surnames"] and c["doi"] and "zenodo" not in c["doi"].lower():
                        ab = f" | Hallazgos/Resumen: {c['abstract_snippet']}" if c.get("abstract_snippet") else ""
                        nac_items.append((c["year"], f"- [Perú, {c['year']}] {c['vancouver']}{ab}"))
                nac_items.sort(key=lambda x: x[0])
                if nac_items:
                    sections.append("ANTECEDENTES NACIONALES (PERÚ) REALES VERIFICADOS (TÍTULO + AUTORES + DOI COMPROBADOS):\n" + "\n".join(x[1] for x in nac_items[:6]))
        except Exception:
            pass

    # C. Artículos seminales de Teoría Científica verificados
    if th_q:
        try:
            r_th = requests.get(
                "https://api.openalex.org/works",
                params={
                    "search": th_q,
                    "filter": "has_doi:true",
                    "sort": "cited_by_count:desc",
                    "per_page": 4
                },
                timeout=4.5
            )
            if r_th.status_code == 200:
                th_items = []
                for w in r_th.json().get("results", []):
                    c = _format_openalex_item(w)
                    if c["title"] and c["surnames"] and c["doi"]:
                        th_items.append(f"- [Referencia Teórica Verificada] {c['vancouver']}")
                if th_items:
                    sections.append("REFERENCIAS TEÓRICAS SEMINALES VERIFICADAS (TÍTULO + AUTORES + DOI COMPROBADOS):\n" + "\n".join(th_items[:3]))
        except Exception:
            pass

    return "\n\n".join(sections)

def _verify_and_enrich_references(report_text):
    """
    Auditor de integridad bibliográfica post-generación:
    1. Para cada DOI citado en el reporte, consulta sus metadatos oficiales en Crossref/OpenAlex y verifica
       el MATCH SIMULTÁNEO entre el TÍTULO del artículo y los AUTORES.
    2. Si la IA entregó solo el DOI (sin autores/título completos), lo expande automáticamente a la cita
       Vancouver COMPLETA (Autores, Título exacto, Revista, Año;Vol(Núm):Págs + DOI) para que el usuario
       pueda corroborarlo visualmente por sí mismo.
    3. Si un DOI es falso (404) o NO coincide en Título + Autores con lo escrito (alucinación cruzada),
       busca en Crossref/OpenAlex exigiendo match estricto Título + Autor; si existe, coloca la referencia
       oficial completa; si no existe (artículo alucinado), elimina el DOI falso y advierte validación manual.
    """
    doi_regex = re.compile(r'\b(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)')
    verify_note = "(Verifica, lee y corrobora personalmente la coincidencia de título, autores y texto completo en el enlace antes de incorporarlo a tu versión final)."
    lines = report_text.splitlines()
    out_lines = []

    for idx, line in enumerate(lines):
        matches = list(doi_regex.finditer(line))
        is_fuente_line = bool(re.search(r"Fuente[s]?\s+verificada[s]?", line, re.IGNORECASE))

        if matches:
            raw_doi = matches[0].group(1).rstrip(".,;)]")
            # Extraer el texto bibliográfico que la IA puso antes del DOI
            stripped_bib = re.sub(r"Fuente[s]?\s+verificada[s]?.*?:", "", line, flags=re.IGNORECASE)
            stripped_bib = re.sub(r"\(Verifica.*?\)", "", stripped_bib, flags=re.IGNORECASE)
            stripped_bib = re.sub(r"(?:Disponible en:\s*)?https?://doi\.org/\S+", "", stripped_bib, flags=re.IGNORECASE)
            stripped_bib = stripped_bib.replace(raw_doi, "").replace("doi:", "").replace("DOI:", "").strip(" .,;-")

            meta = _fetch_doi_metadata(raw_doi)
            prev_context = " ".join(lines[max(0, idx - 3):idx + 1])

            # Caso A: La IA puso solo el DOI (o casi nada de referencia, < 45 caracteres)
            if len(stripped_bib) < 45:
                if meta and _check_title_and_author_match(prev_context, meta, strict_title=False):
                    prefix = "Fuente verificada para agregar a tu bibliografía (Vancouver): " if is_fuente_line else ""
                    out_lines.append(f"{prefix}{meta['vancouver']} {verify_note}")
                else:
                    # El DOI suelto ni siquiera coincide con el párrafo previo -> buscar con el párrafo o limpiar
                    alt = _search_verified_work_strict(prev_context)
                    if alt:
                        prefix = "Fuente verificada para agregar a tu bibliografía (Vancouver): " if is_fuente_line else ""
                        out_lines.append(f"{prefix}{alt['vancouver']} {verify_note}")
                    else:
                        cleaned = re.sub(r"(?:Disponible en:\s*)?https?://doi\.org/" + re.escape(raw_doi) + r"[.,;]?", "", line)
                        cleaned = cleaned.replace(raw_doi, "").strip()
                        if cleaned:
                            out_lines.append(cleaned)
                continue

            # Caso B: La IA escribió Autores + Título + DOI -> exigir MATCH ESTRICTO de Título + Autores
            if meta and _check_title_and_author_match(stripped_bib, meta, strict_title=True):
                # Coincidencia real verificada de Título y Autores: asegurar referencia Vancouver completa oficial
                prefix = "Fuente verificada para agregar a tu bibliografía (Vancouver): " if is_fuente_line else ""
                num_m = re.match(r"^(\d+[\.\)]\s*)", line.strip())
                num_prefix = num_m.group(1) if (num_m and not is_fuente_line) else ""
                note_str = f" {verify_note}" if is_fuente_line else ""
                out_lines.append(f"{prefix}{num_prefix}{meta['vancouver']}{note_str}")
                continue
            else:
                # El DOI no existe (404) o pertenece a OTRO artículo distinto al título/autores escritos
                replacement = _search_verified_work_strict(stripped_bib)
                if replacement:
                    prefix = "Fuente verificada para agregar a tu bibliografía (Vancouver): " if is_fuente_line else ""
                    num_m = re.match(r"^(\d+[\.\)]\s*)", line.strip())
                    num_prefix = num_m.group(1) if (num_m and not is_fuente_line) else ""
                    note_str = f" {verify_note}" if is_fuente_line else ""
                    out_lines.append(f"{prefix}{num_prefix}{replacement['vancouver']}{note_str}")
                else:
                    # No existe match Título+Autor en bases científicas -> retirar DOI alucinado y avisar con transparencia
                    no_doi_line = re.sub(r"(?:Disponible en:\s*)?https?://doi\.org/\S+", "", line, flags=re.IGNORECASE)
                    no_doi_line = no_doi_line.replace(raw_doi, "").replace("doi:", "").replace("DOI:", "").strip(" .,;")
                    no_doi_line = re.sub(
                        r"Fuente[s]?\s+verificada[s]?\s+para\s+agregar\s+a\s+tu\s+bibliograf[íi]a\s*\(Vancouver\)\s*:",
                        "Referencia sugerida (sin DOI confirmado automáticamente; corrobora título y autores en SciELO/PubMed antes de citarla):",
                        no_doi_line,
                        flags=re.IGNORECASE
                    )
                    out_lines.append(no_doi_line)
                continue

        # Caso C: Línea de 'Fuente verificada' donde la IA no puso ningún enlace ni DOI
        if is_fuente_line and "http://" not in line.lower() and "https://" not in line.lower():
            found = _search_verified_work_strict(line)
            if found:
                out_lines.append(f"Fuente verificada para agregar a tu bibliografía (Vancouver): {found['vancouver']} {verify_note}")
            else:
                relabeled = re.sub(
                    r"Fuente[s]?\s+verificada[s]?\s+para\s+agregar\s+a\s+tu\s+bibliograf[íi]a\s*\(Vancouver\)\s*:",
                    "Referencia clásica/institucional para agregar a tu bibliografía (corrobora edición o enlace oficial):",
                    line,
                    flags=re.IGNORECASE
                )
                out_lines.append(relabeled)
            continue

        out_lines.append(line)

    return "\n".join(out_lines)

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def _extract_xml_table(tbl_el, table_counter):
    table_counter[0] += 1
    t_num = table_counter[0]
    t_lines = [f"\n[TABLA DETECTADA #{t_num}]"]
    row_idx = 0
    for tr in tbl_el:
        if tr.tag.endswith("}tr"):
            row_cells = []
            for tc in tr:
                if tc.tag.endswith("}tc"):
                    cell_parts = []
                    for tc_child in tc:
                        c_tag = tc_child.tag.split("}")[-1]
                        if c_tag == "p":
                            p_txt = "".join(t.text for t in tc_child.iter(W_NS + "t") if t.text)
                            if p_txt.strip():
                                cell_parts.append(p_txt.strip())
                        elif c_tag == "tbl":
                            nested_lines = _extract_xml_table(tc_child, table_counter)
                            cell_parts.append(" ".join(nested_lines))
                        elif c_tag in ("sdt", "sdtContent", "customXml", "smartTag"):
                            sub_lines, _ = _extract_xml_node(tc_child, table_counter)
                            cell_parts.append(" ".join(sub_lines))
                    row_cells.append(" ".join(cell_parts).strip())
            if any(row_cells):
                row_idx += 1
                t_lines.append(f"Fila {row_idx}: " + " | ".join(row_cells))
    t_lines.append(f"[FIN DE TABLA #{t_num}]\n")
    return t_lines

def _extract_xml_node(node, table_counter):
    lines = []
    tables_only = []
    for child in node:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            p_txt = "".join(t.text for t in child.iter(W_NS + "t") if t.text)
            if p_txt.strip():
                lines.append(p_txt.strip())
            for txbx_tbl in child.iter(W_NS + "tbl"):
                t_block = _extract_xml_table(txbx_tbl, table_counter)
                lines.extend(t_block)
                tables_only.append("\n".join(t_block))
        elif tag == "tbl":
            t_block = _extract_xml_table(child, table_counter)
            lines.extend(t_block)
            tables_only.append("\n".join(t_block))
        elif tag in ("sdt", "sdtContent", "customXml", "smartTag", "txbxContent"):
            sub_lines, sub_tables = _extract_xml_node(child, table_counter)
            lines.extend(sub_lines)
            tables_only.extend(sub_tables)
    return lines, tables_only

def _extract_docx_content(doc):
    table_counter = [0]
    lines, tables_only = _extract_xml_node(doc.element.body, table_counter)
    full_txt = "\n".join(lines)
    if tables_only:
        full_txt += (
            "\n\n=================================================================\n"
            f"INVENTARIO EXPLÍCITO DE LAS {len(tables_only)} TABLAS EXTRAÍDAS DEL DOCUMENTO WORD:\n"
            "(DEBES AUDITAR OBLIGATORIAMENTE CADA UNA DE ESTAS TABLAS Y CITAR SU CONTENIDO EN 'DONDE DICE:'):\n"
            + "\n".join(tables_only)
        )
    return full_txt

def _extract_pdf_content(doc):
    pages_out = []
    for page in doc:
        pages_out.append(page.get_text("text", sort=True))
    return "\n".join(pages_out)

_NORMATIVAS_CACHE = {}

def extract_text_and_media_from_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()
    text = ""
    media_parts = []
    try:
        uploaded_file.seek(0)
        raw_bytes = uploaded_file.read()
        if name.endswith('.pdf'):
            doc = fitz.open(stream=raw_bytes, filetype="pdf")
            text = _extract_pdf_content(doc)
            # Adjuntar también el PDF nativo para que el motor multimodal de Gemini vea visualmente todas las tablas e imágenes
            if len(raw_bytes) <= 18 * 1024 * 1024:
                media_parts.append(types.Part.from_bytes(data=raw_bytes, mime_type="application/pdf"))
        elif name.endswith('.docx'):
            doc = docx.Document(io.BytesIO(raw_bytes))
            text = _extract_docx_content(doc)
            # Extraer imágenes incrustadas (ej. tablas de SPSS o matrices pegadas como imagen en el Word)
            img_count = 0
            for rel in doc.part.rels.values():
                if "image" in rel.reltype:
                    part = rel.target_part
                    ctype = getattr(part, "content_type", "")
                    blob = getattr(part, "blob", b"")
                    if ctype in ("image/png", "image/jpeg", "image/jpg", "image/webp") and 4000 <= len(blob) <= 4 * 1024 * 1024:
                        media_parts.append(types.Part.from_bytes(data=blob, mime_type=ctype))
                        img_count += 1
                        if img_count >= 15:
                            break
        elif name.endswith('.txt'):
            text = raw_bytes.decode('utf-8', errors='ignore')
        elif name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(io.BytesIO(raw_bytes))
            text = df.to_string()
        elif name.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(raw_bytes))
            text = df.to_string()
    except Exception as e:
        print(f"Error reading {name}: {e}")
    return text, media_parts

def extract_text_from_uploaded_file(uploaded_file):
    txt, _ = extract_text_and_media_from_uploaded_file(uploaded_file)
    return txt

def extract_text_from_local_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    text = ""
    try:
        if ext == '.pdf':
            doc = fitz.open(file_path)
            text = _extract_pdf_content(doc)
        elif ext == '.docx':
            doc = docx.Document(file_path)
            table_counter = [0]
            lines, _ = _extract_xml_node(doc.element.body, table_counter)
            text = "\n".join(lines)
        elif ext == '.txt':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                text = f.read()
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return text

def evaluate_project(uploaded_files, normative_folder, api_key, academic_level="Pregrado", work_stage="Proyecto de investigación"):
    client = genai.Client(api_key=api_key)
    
    # Extraer texto y contenido visual (tablas en PDF o imágenes de tablas en Word) de los archivos subidos
    document_text = ""
    all_media_parts = []
    for uf in uploaded_files:
        txt, m_parts = extract_text_and_media_from_uploaded_file(uf)
        document_text += f"\n\n--- INICIO DE DOCUMENTO: {uf.name} ---\n\n"
        document_text += txt
        document_text += f"\n\n--- FIN DE DOCUMENTO: {uf.name} ---\n\n"
        all_media_parts.extend(m_parts)
        
    if not document_text.strip() and not all_media_parts:
        return "No se encontraron documentos válidos o están vacíos."

    # Extraer texto de la carpeta de normativas local del servidor (con caché en memoria)
    if normative_folder in _NORMATIVAS_CACHE:
        institutional_normative = _NORMATIVAS_CACHE[normative_folder]
    elif os.path.isdir(normative_folder):
        institutional_normative = ""
        for root, dirs, files in os.walk(normative_folder):
            for file in files:
                if file.lower().endswith(('.pdf', '.docx', '.txt')):
                    path = os.path.join(root, file)
                    institutional_normative += f"\n\n--- NORMATIVA: {file} ---\n\n"
                    institutional_normative += extract_text_from_local_file(path)
        _NORMATIVAS_CACHE[normative_folder] = institutional_normative
    else:
        return f"No se encontró la carpeta de normativas en el servidor: {normative_folder}"
    
    # Pre-buscar banco de artículos científicos reales (Internacionales, Nacionales-Perú y Teorías) con match Título + Autores + DOI verificado
    verified_lit_bank = _fetch_verified_literature_bank(document_text, client)

    sys_instruction = (
        "Eres un panel de alta especialidad académica (Asesor Científico Doctoral, Metodólogo y Bioestadístico, Auditor Clínico-Epidemiológico, Corrector de Estilo Vancouver y Revisor Normativo) actuando como Editor Académico Principal de Perspecta Salud.\n"
        "Tu misión es auditar minuciosamente el documento completo (incluyendo cada párrafo, cada tabla de resultados y cada tabla o instrumento de los anexos) y elaborar una HOJA DE RUTA DE CORRECCIONES IMPERATIVAS secuencial (sección por sección, desde la carátula hasta el último anexo), empleando el formato quirúrgico de contraste directo 'DONDE DICE:' / 'COLOCA:'.\n\n"
        "ESTRUCTURA OBLIGATORIA DEL REPORTE:\n\n"
        "ENCABEZADO E INSTRUCCIÓN GENERAL (LO LOGRADO Y LO QUE NO DEBES MODIFICAR):\n"
        "- Inicia con el título: HOJA DE RUTA DE CORRECCIONES IMPERATIVAS — PERSPECTA SALUD\n"
        "- Debajo, redacta el bloque 'Instrucción general y aspectos logrados:':\n"
        "  * Reconoce con sobriedad académica (sin adjetivos exagerados ni halagos como 'excelente' o 'maravilloso') los aciertos metodológicos y estructurales del trabajo.\n"
        "  * VALIDACIÓN PRIORITARIA DEL ANÁLISIS ESTADÍSTICO: Verifica primero si la persona ya realizó o propuso un análisis estadístico en su documento. SI EL ANÁLISIS ESTADÍSTICO YA REALIZADO O PROPUESTO ESTÁ BIEN (es acertado para la naturaleza de sus variables y responde a sus objetivos, ej. Rho de Spearman, Chi-cuadrado, etc.), felicítala con sobriedad en esta Instrucción General, explícale en palabras sencillas por qué esa prueba es la correcta para la escala de sus variables y sus objetivos, e indícale explícitamente: 'NO modifiques la prueba estadística [Nombre de la prueba] (es metodológicamente correcta)'. Si su elección estadística ya es correcta, NO agregues pasos de SPSS para cambiar de prueba (solo corrige en las tablas detalles de reporte como cambiar p = 0,000 por p < 0,001 o usar coma decimal).\n"
        "  * Indica también qué otros elementos bien alineados NO debe modificar por inseguridad (por ejemplo, mantener los objetivos específicos en su formulación global si en SPSS los resultados fueron procesados por puntajes globales y no desglosados por dimensiones).\n\n"
        "SECCIONES SECUENCIALES DE LA HOJA DE RUTA (Recorre el documento en su orden físico real de principio a fin, numerando correlativamente CORRECCIÓN 01, CORRECCIÓN 02, CORRECCIÓN 03...):\n\n"
        "1. PÁGINAS PRELIMINARES Y RESUMEN\n"
        "- Audita Carátula y Título (delimitación precisa de variables, población, establecimiento/lugar y año).\n"
        "- Audita Dedicatoria, Agradecimiento e Informe de Similitud (detecta páginas en blanco y corrige errores ortográficos reales como 'fé' con tilde -> 'fe' sin tilde, 'qué ilumine' -> 'que ilumine', redacción cortada).\n"
        "- Audita Resumen y Palabras Claves: precisión metodológica ('tipo básico' en lugar de 'nivel básico' cuando el nivel/alcance es correlacional/descriptivo), uso obligatorio de coma decimal en todos los porcentajes (ej. 68,3 % en vez de 68.3 %), reemplazo obligatorio de 'p = 0.000' o 'p = 0,000' por 'p < 0,001' (explicando entre paréntesis que en bioestadística nunca se reporta p = 0,000), y al menos 5 descriptores científicos (DeCS/MeSH).\n\n"
        "2. CAPÍTULO I: INTRODUCCIÓN\n"
        "- 1.1 Realidad Problemática:\n"
        "  * Estructura en embudo estricto: Verifica que transite desde el nivel MUNDIAL (OMS), al REGIONAL (OPS/Latinoamérica), luego al NACIONAL (MINSA/INEI) y aterrice obligatoriamente en el nivel LOCAL (caracterización demográfica y epidemiológica del distrito y del establecimiento o institución específica donde se ejecuta el estudio). Si falta el cierre local o algún nivel del embudo, redacta en 'COLOCA:' el párrafo completo listo para insertar CON su llamada de cita Vancouver e incluye inmediatamente al final de ese mismo 'COLOCA:' la línea 'Fuente verificada para agregar a tu bibliografía (Vancouver):' con TODOS los datos bibliográficos completos (Autores, Título completo, Revista/Institución, Año y enlace oficial o DOI real verificado).\n"
        "  * Rigor epidemiológico: Detecta errores conceptuales o de indicadores (por ejemplo, confundir casos absolutos de muertes con tasas por cien mil nacidos vivos, o expresiones mal redactadas como 'maternidad materna' en vez de 'mortalidad materna').\n"
        "- 1.2 Antecedentes:\n"
        "  * Umbral mínimo según nivel académico declarado: Para Pregrado, exige como MÍNIMO 5 antecedentes internacionales y 5 nacionales (al menos 10 en total); para Posgrado (Maestría/Doctorado), exige como MÍNIMO 10 internacionales y 10 nacionales (al menos 20 en total).\n"
        "  * Orden cronológico ascendente estricto: Deben ordenarse desde el más antiguo hasta el más reciente (ej. 2020 -> 2023 -> 2024 -> 2025). En 'COLOCA:', muestra la secuencia cronológica exacta ordenada de los autores que ya tiene y selecciona los estudios reales faltantes prioritariamente del 'BANCO DE LITERATURA CIENTÍFICA PRE-VERIFICADA' incluido en este prompt (o búscalos en tiempo real con google_search corroborando el match exacto entre Título del artículo y Autores antes de proponerlos). Para CADA estudio sugerido incluye su párrafo redactado listo para insertar y su 'Fuente verificada para agregar a tu bibliografía (Vancouver):' con Autores completos, Título exacto del artículo, Revista científica, Año;Volumen(Número):Páginas y DOI real verificado (https://doi.org/...).\n"
        "  * Corrige erratas textuales o redundancias dentro de los antecedentes (ej. 'conocimiento alto sobre conocimiento').\n"
        "- 1.3 Bases Teóricas y Marco Conceptual:\n"
        "  * Precisión clínica y científica: Detecta errores terminológicos graves (ej. escribir 'hiperemia gravídica' en lugar de 'hiperémesis gravídica', mezclar signos clínicos distintos en una misma cita, o escribir nombres bacteriológicos como 'E-coli' en lugar de 'Escherichia coli' en cursiva).\n"
        "  * Articulación con Teoría Científica (OBLIGATORIO INCLUIR REFERENCIA COMPLETA VERIFICADA DE LA TEORÍA): Verifica que las variables y sus dimensiones estén respaldadas por una teoría o modelo científico clásico (ej. Modelo de Creencias en Salud de Rosenstock / Janz y Becker, Teoría de Jean Watson, Modelo de Dorothea Orem, etc.). Si falta o es débil, redacta en 'COLOCA:' el párrafo completo de articulación teórica listo para copiar e insertar al final de las Bases Teóricas con su llamada numérica de cita, e INCLUYE OBLIGATORIAMENTE al pie de ese mismo bloque 'COLOCA:' la referencia Vancouver COMPLETA (Autores completos, Título completo exacto de la obra/artículo, Revista o Editorial, Año;Volumen(Número):Páginas y DOI real verificado https://doi.org/...).\n"
        "- 1.4 / 1.5 Justificación, Problemas, Objetivos e Hipótesis:\n"
        "  * Verifica numeración de subtítulos, signos de interrogación '?' en todos los problemas, mayúsculas en nombres propios de instituciones, y coherencia exacta entre problemas, objetivos, hipótesis y lo procesado en resultados.\n\n"
        "3. CAPÍTULO II: METODOLOGÍA\n"
        "- Tipo, alcance y diseño; Población, muestra y muestreo:\n"
        "  * Precisa 'tipo de investigación básica (o aplicada)' y 'alcance/nivel descriptivo/correlacional/explicativo', así como los meses exactos del periodo de estudio.\n"
        "  * Si el muestreo es no probabilístico por conveniencia o censal y su justificación es débil, entrega en 'COLOCA:' el párrafo técnico blindado (criterio de captación continua / saturación de flujo y cautela en la extrapolación).\n"
        "  * Verifica que no haya saltos en la numeración correlativa de citas Vancouver (ej. pasar de la cita 32 a la 34 omitiendo la 33).\n"
        "- Plan de Análisis y Tratamiento Estadístico de los Datos:\n"
        "  * SI EL ANÁLISIS ESTADÍSTICO DEL ESTUDIANTE YA ES ACERTADO: Recuerda que ya lo validaste en la Instrucción General; no agregues una corrección para cambiar de prueba ni des pasos en SPSS.\n"
        "  * SOLO SI EL ANÁLISIS ESTADÍSTICO ES INCORRECTO, INCOMPLETO O FUE OMITIDO: Crea un bloque 'CORRECCIÓN XX — Plan de Análisis Estadístico y Procesamiento en SPSS [CRÍTICO]'. En 'DONDE DICE:' cita lo que propuso u omitió. En 'COLOCA:' incluye obligatoriamente:\n"
        "    1) Explicación sencilla, con tacto pedagógico y sin jerga compleja, de por qué su propuesta anterior no se ajusta a la escala de sus variables y por qué hemos escogido UN ÚNICO tratamiento estadístico definitivo (PROHIBIDO usar la palabra 'ruta' o 'rutas' o dar alternativas condicionales), con un ejemplo práctico aplicado a sus variables.\n"
        "    2) El párrafo modelo del Plan de Análisis listo para insertar en su metodología.\n"
        "    3) El paso a paso numerado en SPSS: (1) Menú exacto en SPSS, (2) En qué casilla ubicar cada variable, (3) Qué botones y casillas marcar, y (4) Cómo interpretar el valor 'Sig.' o p-valor (< 0,05) y reportar p < 0,001 cuando SPSS muestre 0,000.\n"
        "- Aspectos Éticos (Marco bioético internacional y legal peruano vigente, SIN pedir trámite de Comité de Ética):\n"
        "  * NUNCA solicites tramitar aprobación ante un Comité de Ética ni adjuntar cartas de comité. En su lugar, en 'COLOCA:' entrega directamente el párrafo normativo blindado listo para reemplazar/insertar que integre: Declaración de Helsinki, Pautas Éticas CIOMS 2016, Informe Belmont (autonomía, beneficencia, no maleficencia y justicia), Consideraciones Éticas para la Investigación en Salud con Seres Humanos del MINSA (Resolución Ministerial N.° 233-2020-MINSA) y la Ley N.° 29733 (Ley de Protección de Datos Personales) junto con su reglamento vigente aprobado por Decreto Supremo N.° 016-2024-JUS (codificación alfanumérica de instrumentos).\n\n"
        "4. CAPÍTULO III: RESULTADOS (Aplica cuando el documento incluya capítulo o tablas de resultados)\n"
        "- Audita fila por fila cada tabla extraída del documento ([INICIO DE TABLA] / [TABLA X]) y contrástala con el cuestionario del Anexo:\n"
        "  * Detecta cálculos estadísticamente inválidos (por ejemplo, reportar una fila 'Promedio: 25 años' en la Tabla 1 cuando el instrumento del Anexo solo recogió intervalos de edad como '18-25, 26-35, 36 a más' y no la edad exacta).\n"
        "  * Exige reemplazar en todas las tablas e interpretaciones 'p = 0,000' o '0.000' por '< 0,001' y escribir '(ρ = 0,xxx; p < 0,001)'.\n"
        "  * Exige uniformizar todos los números y porcentajes con un decimal y coma decimal (ej. 50,0 %; 33,3 %; 100,0 % en lugar de mezclar '50', '33.3', '100.0').\n\n"
        "5. CAPÍTULO IV (DISCUSIÓN Y CONCLUSIONES), DECLARACIÓN DE IA Y REFERENCIAS BIBLIOGRÁFICAS\n"
        "- Discusión (si está presente): Detecta erratas ortográficas/tipográficas, exige formato '(ρ = 0,xxx; p < 0,001)' y verifica que no existan contradicciones lógicas entre lo que afirma la discusión y lo que realmente hallaron sus propios antecedentes citados (redactando en 'COLOCA:' el párrafo exacto que resuelve la contradicción).\n"
        "- Conclusiones (si están presentes): Verifica que cada conclusión responda en orden estricto a los objetivos (General, OE1, OE2, OE3...) e incorpore los valores estadísticos y porcentuales exactos obtenidos en Resultados.\n"
        "- Declaración de Uso de IA (si está presente): Indica precisar herramienta, modelo específico y proveedor (ej. 'ChatGPT, modelo GPT-4o, proveedor OpenAI').\n"
        "- Referencias Bibliográficas y Estilo Vancouver:\n"
        "  * Verifica que la lista final tenga numeración visible correlativa (1., 2., 3. ...) según orden de aparición en el texto.\n"
        "  * Corrige 'Disponible de:' por 'Disponible en:', fechas futuras erróneas y posición de las llamadas numéricas en todo el manuscrito (el número entre paréntesis debe ir ANTES del punto final y sin espacios dobles, ej. '...durante el embarazo (1).').\n"
        "  * Lista todas las fuentes nuevas incorporadas en las correcciones anteriores con su referencia COMPLETA en estilo Vancouver (Autores completos, Título completo exacto, Revista, Año;Volumen(Número):Páginas y DOI real verificado https://doi.org/...), incluyendo SIEMPRE la nota obligatoria: '(Verifica, lee y corrobora personalmente la coincidencia de título, autores y texto completo en el enlace antes de incorporarlo a tu versión final)'.\n\n"
        "6. ANEXOS (CONSISTENCIA, OPERACIONALIZACIÓN, INSTRUMENTO Y CONSENTIMIENTO)\n"
        "- Audita exhaustivamente cada tabla y texto de los anexos:\n"
        "  * Anexo 1 (Matriz de Consistencia): Verifica que incluya todos los problemas específicos y que las dimensiones listadas coincidan con lo realmente evaluado en los objetivos y resultados.\n"
        "  * Anexo 2 (Matriz de Operacionalización de Variables): Contrasta los baremos o puntajes del Anexo 2 contra la metodología y el cuestionario del Anexo 3 (por ejemplo, detectar si en el Anexo 2 colocaron 'Nivel alto 18-20 / Nivel medio 14-17 / Nivel bajo 10-13' cuando el cuestionario real es de 10 ítems con baremo de 0 a 10 puntos). Verifica también que estén incluidas las variables sociodemográficas si corresponden al OE1.\n"
        "  * Anexo 3 (Instrumento de Recolección de Datos): Detecta plurales incorrectos en títulos ('SIGNOS DE ALARMAS' -> 'SIGNOS DE ALARMA'), ítems con opciones de respuesta duplicadas (indicando reemplazar la opción repetida por el distractor original del autor del instrumento) y encabezados desordenados o sucios en la tabla Likert ('Totalmente en desacuerdo (1)', 'En desacuerdo (2)', 'Ni de acuerdo ni en desacuerdo (3)', 'De acuerdo (4)', 'Totalmente de acuerdo (5)').\n"
        "  * Anexo de Consentimiento Informado: Adapta la concordancia gramatical de género a la población real estudiada (ej. si son gestantes, usar 'fui informada', 'para usted' en vez de 'informado (a)' o 'para él (ella)') y corrige la relación de principios bioéticos ('Confidencialidad y privacidad (principio de respeto a la autonomía)' en lugar de asociar confidencialidad con justicia).\n\n"
        "FORMATO ESTRICTO E INQUEBRANTABLE PARA CADA CORRECCIÓN:\n"
        "Cada hallazgo debe redactarse EXACTAMENTE con esta estructura de 3 líneas/bloques para que el sistema lo convierta automáticamente en tablas ejecutivas dentro del archivo Word:\n"
        "CORRECCIÓN [Número de 2 dígitos: 01, 02, 03...] — [Subtítulo claro de la sección, tabla o anexo y tema] [Agregar '[CRÍTICO]' al final del título si la observación compromete la aprobación metodológica, clínica, estadística o normativa]\n"
        "DONDE DICE: [Cita textual exacta entre comillas de lo que dice el documento actual, o descripción precisa de la fila/celda de la tabla con sus valores actuales. Si hay varios errores en esa misma sección, numéralos como 1., 2., 3.]\n"
        "COLOCA: [Instrucción directa, asertiva e imperativa con el texto exacto corregido listo para copiar y pegar, numerando 1., 2., 3. en correspondencia exacta con DONDE DICE, e incluyendo entre paréntesis la breve explicación pedagógica de por qué se aplica esa mejora. REGLA INQUEBRANTABLE: Siempre que en 'COLOCA:' propongas una teoría científica, un modelo teórico, un antecedente o un texto que cite datos o autores nuevos, incluye al final de ese mismo bloque 'COLOCA:' la línea 'Fuente verificada para agregar a tu bibliografía (Vancouver): Apellido1 Iniciales, Apellido2 Iniciales. Título completo exacto del artículo. Revista Científica. Año;Volumen(Número):Páginas. Disponible en: https://doi.org/10.xxxx/xxxx (Verifica, lee y corrobora personalmente la coincidencia de título, autores y texto completo en el enlace antes de incorporarlo a tu versión final)'.]\n\n"
        "REGLAS ESTRICTAS DE NEGOCIO Y ESTILO:\n"
        "- PROHIBIDO ENTREGAR SOLO EL DOI SIN REFERENCIA COMPLETA: NUNCA pongas únicamente el código o enlace DOI suelto. Para que el usuario pueda corroborar el artículo por sí mismo, TODA fuente propuesta debe mostrar obligatoriamente Autores completos, Título completo exacto del artículo, Revista, Año;Volumen(Número):Páginas y el enlace DOI.\n"
        "- VALIDACIÓN OBLIGATORIA DE MATCH TÍTULO + AUTORES ANTES DE PROPONER UN ARTÍCULO: Antes de decidir incluir un artículo científico como propuesta, selecciónalo del 'BANCO DE LITERATURA CIENTÍFICA PRE-VERIFICADA' provisto en el prompt o verifica con google_search que el Título exacto del artículo y sus Autores coincidan al 100% con el registro del DOI. Queda terminantemente prohibido inventar estudios o mezclar autores con títulos o DOIs ajenos.\n"
        "- PROHIBIDO COLOCAR NÚMEROS DE PÁGINA: NUNCA coloques '(Página X)' ni 'Pág. X' en los títulos ni en el cuerpo de las correcciones, ya que no son referencia exacta. Ubica al estudiante indicando la sección exacta, el subtítulo, el número de tabla o el anexo (ej. 'CORRECCIÓN 01 — Carátula y Título', 'CORRECCIÓN 04 — 1.1 Realidad Problemática [CRÍTICO]', 'CORRECCIÓN 10 — Tabla 1 de Resultados [CRÍTICO]').\n"
        "- PROHIBIDO COLOCAR CALIFICACIÓN O PUNTAJE: No asignes notas sobre 20 ni sobre 100.\n"
        "- PROHIBIDO PROPONER TEORÍAS O TEXTOS SIN FUENTE VERIFICADA: Jamás propongas insertar una teoría, modelo conceptual, cifra epidemiológica o antecedente sin entregar inmediatamente en el mismo bloque 'COLOCA:' su referencia bibliográfica completa en estilo Vancouver con su DOI real verificado o enlace oficial directo.\n"
        "- PROHIBIDO SOLICITAR TRÁMITE DE COMITÉ DE ÉTICA: Nunca pidas al estudiante presentar su trabajo a un Comité de Ética ni adjuntar cartas de aprobación del CIEI; entrégale directamente el párrafo ético normativo listo para insertar.\n"
        "- CONFIDENCIALIDAD INSTITUCIONAL: NUNCA menciones las siglas UPN ni 'Universidad Privada del Norte' en ninguna parte del reporte (usa 'tu universidad' o 'la institución'), ni inventes nombres propios de asesores que no figuren en el documento subido.\n"
        "- TEXTO LIMPIO SIN MARKDOWN: No uses asteriscos (**), ni numerales (#), ni tablas Markdown (|---|) en tu respuesta; usa únicamente las etiquetas en mayúsculas 'CORRECCIÓN 01 — ...', 'DONDE DICE:' y 'COLOCA:' para que el generador nativo de Word construya las tablas.\n\n"
        "Tu tarea final es MUY IMPORTANTE: Al puro final de tu respuesta, debes agregar un bloque JSON estricto con la siguiente estructura (este bloque es oculto para el sistema y nunca lo ve el estudiante, no agregues nada después):\n"
        "```json\n"
        "{\n"
        '  "project_title": "[Extrae el título exacto de la tesis/proyecto del documento. Si no está, pon \\"Título no especificado\\"]",\n'
        '  "authors": "[Extrae el nombre del autor o autores del documento. Si no está, pon \\"Autor no especificado\\"]"\n'
        "}\n"
        "```"
    )

    lit_bank_section = (
        f"=================================================================\n"
        f"BANCO DE LITERATURA CIENTÍFICA PRE-VERIFICADA EN OPENALEX/CROSSREF PARA ESTA TESIS:\n"
        f"(ESTOS ARTÍCULOS SON 100% REALES Y YA TIENEN MATCH VERIFICADO DE TÍTULO + AUTORES + REVISTA + DOI. "
        f"ÚSALOS PRIORITARIAMENTE CUANDO SUGIERAS ANTECEDENTES O REFERENCIAS TEÓRICAS COPIANDO SUS DATOS EXACTOS):\n"
        f"{verified_lit_bank}\n\n"
    ) if verified_lit_bank else ""

    prompt = (
        f"NORMATIVA Y RÚBRICA INSTITUCIONAL DE REFERENCIA:\n{institutional_normative}\n\n"
        f"{lit_bank_section}"
        f"=================================================================\n"
        f"DATOS DE LA SOLICITUD A AUDITAR:\n"
        f"- Nivel académico declarado: {academic_level}\n"
        f"- Etapa del trabajo declarada: {work_stage}\n\n"
        f"DOCUMENTO COMPLETO DEL PROYECTO / TESIS DEL ESTUDIANTE (INCLUYE TEXTO, TABLAS EXTRAÍDAS Y ARCHIVOS VISUALES ADJUNTOS):\n\n{document_text}\n\n"
        f"=================================================================\n"
        f"INSTRUCCIÓN FINAL OBLIGATORIA DE FORMATO Y AUDITORÍA DE TABLAS:\n"
        f"1. AUDITORÍA OBLIGATORIA DEL TEXTO DENTRO DE LAS TABLAS: Está TERMINANTEMENTE PROHIBIDO obviar o pasar por alto el contenido de las tablas (tanto las marcadas como [TABLA DETECTADA #...] como cualquier tabla visible en el PDF o imágenes adjuntas). Debes leer fila por fila y columna por columna TODAS las tablas del documento (Tablas de Resultados 1, 2, 3, 4..., Matriz de Consistencia del Anexo 1, Matriz de Operacionalización de Variables del Anexo 2, Instrumentos/Cuestionarios en tabla del Anexo 3, Cronograma y Presupuesto) y generar bloques 'CORRECCIÓN XX' específicos citando textualmente en 'DONDE DICE:' los valores numéricos, porcentajes, baremos, dimensiones, encabezados o ítems que aparecen DENTRO de las celdas de dichas tablas.\n"
        f"2. Genera AHORA la 'HOJA DE RUTA DE CORRECCIONES IMPERATIVAS — PERSPECTA SALUD' recorriendo el documento sección por sección y capítulo por capítulo hasta el último anexo, SIN colocar números de página '(Página X)' ni 'Pág. X'. "
        f"Inicia con 'Instrucción general y aspectos logrados:' (indicando qué está bien y qué NO debe modificar, incluyendo la validación de la prueba estadística si ya es correcta). "
        f"Luego estructura ABSOLUTAMENTE CADA CORRECCIÓN con estas 3 etiquetas exactas en líneas separadas:\n"
        f"CORRECCIÓN 01 — [Título de la Sección, Tabla o Anexo] [CRÍTICO]\n"
        f"DONDE DICE: [Cita textual exacta o valores actuales de las celdas de la tabla con error/omisión]\n"
        f"COLOCA: [Texto exacto corregido listo para copiar e insertar + breve explicación entre paréntesis + Si propones una teoría, modelo, dato o antecedente nuevo, incluye obligatoriamente ahí mismo la referencia Vancouver COMPLETA (NUNCA solo el DOI): 'Fuente verificada para agregar a tu bibliografía (Vancouver): Apellido1 Iniciales, Apellido2 Iniciales. Título completo exacto del artículo. Revista Científica. Año;Volumen(Número):Páginas. Disponible en: https://doi.org/10.xxxx/xxxx (Verifica, lee y corrobora personalmente la coincidencia de título, autores y texto completo en el enlace antes de incorporarlo a tu versión final)']\n"
        f"(PROHIBIDO entregar solo el DOI sin autores/título, PROHIBIDO colocar números de página y PROHIBIDO usar el formato antiguo de 'OBSERVACIÓN / EXPLICACIÓN / ACCIÓN RECOMENDADA'; usa EXCLUSIVAMENTE 'CORRECCIÓN XX', 'DONDE DICE:' y 'COLOCA:')."
    )

    models_to_try = [
        'gemini-3.1-pro-preview',
        'gemini-pro-latest',
        'gemini-3.8-flash',
        'gemini-3.7-flash',
        'gemini-flash-latest'
    ]
    
    last_error = ""
    contents_payload = [prompt, *all_media_parts] if all_media_parts else prompt

    for model_name in models_to_try:
        for attempt in range(3): # Reintenta 3 veces por cada modelo
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents_payload,
                    config=types.GenerateContentConfig(
                        system_instruction=sys_instruction,
                        tools=[{"google_search": {}}],
                        max_output_tokens=32768,
                        temperature=0.2,
                    )
                )
                
                raw_report = response.text
                
                # Corrección de Mojibake (Si la API devuelve texto mal decodificado en ISO-8859-1 en lugar de UTF-8)
                try:
                    raw_report = raw_report.encode('latin-1').decode('utf-8')
                except UnicodeError:
                    pass
                
                
                metadata = {
                    "score": 0,
                    "has_ethics": False,
                    "has_grave_obs_before_methodology": True
                }
                report_clean = raw_report
                
                json_match = re.search(r'```json\s*(\{.*?\})\s*```', raw_report, re.DOTALL)
                if json_match:
                    try:
                        metadata = json.loads(json_match.group(1))
                        report_clean = raw_report[:json_match.start()].strip()
                    except:
                        pass

                # Verificación y enriquecimiento automático de DOIs y fuentes en Crossref / doi.org
                try:
                    report_clean = _verify_and_enrich_references(report_clean)
                except Exception:
                    pass

                return {
                    "report": report_clean,
                    "metadata": metadata
                }
                
            except Exception as e:
                last_error = str(e)
                # Si el modelo no existe o está bloqueado para esta API Key (404), no perdemos tiempo reintentando, pasamos al siguiente modelo.
                if "404" in last_error or "NOT_FOUND" in last_error or "is no longer available" in last_error:
                    break 
                
                # Si es un error de saturación temporal (503), esperamos un poco y reintentamos con el mismo modelo
                time.sleep(5 * (attempt + 1))
                
    return {"report": f"Error Crítico: Servidores de IA saturados. Se intentó conectar con todos los modelos disponibles sin éxito. Último error: {last_error}", "metadata": {}}
