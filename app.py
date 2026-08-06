"""
app.py - Version web (navigateur) de l'outil de remplissage des rapports GTG.
Reutilise directement la logique testee de fill_gtg_report.py.

Chaque utilisateur entre sa PROPRE cle API gratuite (Gemini, et en option
Mistral en secours) - rien n'est stocke cote serveur entre les sessions.
"""

import io
import json
import os
import re
import tempfile
import time
import zipfile

import openpyxl
import streamlit as st

from fill_gtg_report import (
    HEURES,
    QuotaExhaustedError,
    build_extraction_prompt,
    build_sensor_map,
    extract_from_image_gemini,
    extract_from_image_mistral,
    norm_tag,
)

try:
    from google import genai
except ImportError:
    genai = None


st.set_page_config(page_title="Remplissage rapport GTG", page_icon="🔧", layout="centered")
st.title("🔧 Remplissage automatique des rapports GTG")
st.caption(
    "Uploadez votre rapport Excel et vos photos d'écran de tournée — "
    "l'outil lit les valeurs des capteurs et remplit le rapport automatiquement."
)

with st.expander("ℹ️ Comment ça marche / confidentialité"):
    st.markdown(
        "- Chaque personne utilise **sa propre clé API gratuite** (Gemini, et en option Mistral).\n"
        "- Rien n'est conservé sur le serveur après votre session : vos fichiers et clés "
        "restent dans votre navigateur/session uniquement.\n"
        "- Clé Gemini gratuite : https://aistudio.google.com/apikey\n"
        "- Clé Mistral gratuite (optionnelle, en secours si Gemini est saturé) : "
        "https://console.mistral.ai/api-keys"
    )

# ---------------------------------------------------------------------------
# Etat de session
# ---------------------------------------------------------------------------
defaults = {
    "sensor_map": None, "excel_bytes": None, "excel_name": None,
    "matched": [], "unmatched": [], "seen_tags": set(), "processed_idx": 0,
    "photo_names": [], "run_state": None, "final_bytes": None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ---------------------------------------------------------------------------
# 1) Fichier Excel
# ---------------------------------------------------------------------------
st.header("1. Modèle de rapport Excel")
st.caption(
    "Uploadez UN SEUL modèle Excel (même structure pour toutes vos turbines). "
    "Un fichier rempli séparé sera généré automatiquement pour chaque turbine détectée."
)
excel_file = st.file_uploader("Modèle Excel (.xlsx)", type=["xlsx"])

if excel_file is not None and excel_file.name != st.session_state.excel_name:
    with st.spinner("Analyse de la structure du classeur..."):
        excel_bytes = excel_file.read()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(excel_bytes)
            tmp_path = tmp.name
        try:
            sensor_map = build_sensor_map(tmp_path)
        finally:
            os.unlink(tmp_path)
        st.session_state.sensor_map = sensor_map
        st.session_state.excel_bytes = excel_bytes
        st.session_state.excel_name = excel_file.name
        # reset une eventuelle tournee precedente
        st.session_state.matched, st.session_state.unmatched = [], []
        st.session_state.seen_tags, st.session_state.processed_idx = set(), 0
        st.session_state.run_state, st.session_state.final_bytes = None, None

if st.session_state.sensor_map:
    st.success(f"**{st.session_state.excel_name}** chargé — {len(st.session_state.sensor_map)} capteur(s)/champ(s) détecté(s).")

st.divider()

# ---------------------------------------------------------------------------
# 2) Parametres
# ---------------------------------------------------------------------------
st.header("2. Paramètres de la tournée")
c1, c2 = st.columns(2)
with c1:
    heure = st.selectbox("Heure du relevé", HEURES)
with c2:
    auto_turbine = st.checkbox(
        "🔎 Détecter automatiquement la turbine sur chaque photo",
        value=True,
        help="Coché : vous pouvez mélanger les photos de plusieurs turbines (GTG 811, 812, 813...) "
             "en un seul lot, l'outil lit le numéro affiché sur chaque écran et trie tout seul. "
             "Décoché : toutes les photos uploadées sont considérées comme appartenant à UNE seule "
             "turbine, que vous précisez ci-dessous.",
    )
turbine = None
if not auto_turbine:
    turbine = st.text_input("Numéro de turbine (ex: 812)", "")

st.subheader("Clé API (Gemini — gratuite)")
gemini_key = st.text_input("Votre clé Gemini", type="password",
                            help="Obtenez une clé gratuite sur https://aistudio.google.com/apikey")

use_fallback = st.checkbox("Activer un secours automatique avec Mistral si Gemini est saturé")
mistral_key = None
if use_fallback:
    mistral_key = st.text_input("Votre clé Mistral", type="password",
                                 help="Obtenez une clé gratuite sur https://console.mistral.ai/api-keys")

st.divider()

# ---------------------------------------------------------------------------
# 3) Photos
# ---------------------------------------------------------------------------
st.header("3. Photos des écrans")
photos = st.file_uploader(
    "Toutes les photos de cette tournée (peu importe l'ordre)",
    type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True,
)
if photos:
    st.write(f"📷 {len(photos)} photo(s) sélectionnée(s).")

st.divider()

# ---------------------------------------------------------------------------
# 4) Lancer
# ---------------------------------------------------------------------------
st.header("4. Lancer le remplissage")

can_run = bool(st.session_state.sensor_map and photos and gemini_key and heure)
if not can_run:
    st.info("Complétez les étapes 1 à 3 ci-dessus pour activer le bouton.")

run_clicked = st.button("▶️ Remplir le rapport", disabled=not can_run, type="primary")
resume_clicked = False
if st.session_state.run_state == "stopped":
    st.warning(
        f"⏸️ Arrêté après {st.session_state.processed_idx}/{len(st.session_state.photo_names)} photo(s) "
        "— quota probablement épuisé. Changez de clé ci-dessus si besoin, puis :"
    )
    resume_clicked = st.button("🔁 Reprendre là où ça s'est arrêté")

log_box = st.container(height=250) if hasattr(st, "container") else st.container()
progress_bar = st.progress(0.0)


def run_batch(resume=False):
    sensor_map = st.session_state.sensor_map
    prompt = build_extraction_prompt(turbine or None, auto_detect_turbine=auto_turbine)
    client = genai.Client(api_key=gemini_key)

    photo_list = photos if not resume else st.session_state.get("_photo_cache")
    st.session_state["_photo_cache"] = photo_list
    st.session_state.photo_names = [p.name for p in photo_list]

    if not resume:
        st.session_state.matched, st.session_state.unmatched = [], []
        st.session_state.seen_tags, st.session_state.processed_idx = set(), 0

    # matched/unmatched entries now carry a 6th/5th field: turbine id (or None)
    start_idx = st.session_state.processed_idx
    total = len(photo_list)
    active_backend = "gemini"

    for i in range(start_idx, total):
        photo = photo_list[i]
        log_box.write(f"**[{i + 1}/{total}]** {photo.name} (via {active_backend}) ...")
        photo.seek(0)
        suffix = os.path.splitext(photo.name)[1] or ".jpg"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(photo.read())
            tmp_path = tmp.name

        try:
            try:
                if active_backend == "gemini":
                    result = extract_from_image_gemini(client, tmp_path, "gemini-3.5-flash", prompt)
                    time.sleep(3)
                else:
                    result = extract_from_image_mistral(tmp_path, "pixtral-large-latest", prompt, mistral_key)
            except QuotaExhaustedError as qe:
                if use_fallback and mistral_key and active_backend == "gemini":
                    log_box.warning(f"⚠️ Gemini saturé ({qe}). Bascule automatique sur Mistral...")
                    active_backend = "mistral"
                    result = extract_from_image_mistral(tmp_path, "pixtral-large-latest", prompt, mistral_key)
                else:
                    st.session_state.processed_idx = i
                    st.session_state.run_state = "stopped"
                    st.error(
                        f"⚠️ Quota épuisé après {i}/{total} photo(s) ({qe}). "
                        "Changez de clé API ci-dessus puis cliquez sur *Reprendre*."
                    )
                    return
        finally:
            os.unlink(tmp_path)

        title = result.get("screen_title") or "(titre non lu)"
        screen_turbine = result.get("turbine") if auto_turbine else turbine
        readings = result.get("readings", [])
        log_box.write(f"    → {len(readings)} valeur(s) lue(s) ({title}"
                       + (f", turbine {screen_turbine}" if screen_turbine else "") + ")")

        for reading in readings:
            tag = norm_tag(reading.get("tag"))
            value = reading.get("value")
            reading_turbine = (reading.get("turbine") if auto_turbine else None) or screen_turbine
            if tag is None or value is None:
                continue
            if auto_turbine and not reading_turbine:
                st.session_state.unmatched.append((tag, value, title, "turbine non identifiée", None))
                continue
            st.session_state.seen_tags.add(tag)
            if tag in sensor_map:
                sheet_name = sensor_map[tag]['sheet']
                coord = sensor_map[tag]['cells'].get(heure)
                if coord is None:
                    st.session_state.unmatched.append((tag, value, title, "heure introuvable", reading_turbine))
                    continue
                coord_display = "+".join(coord) if isinstance(coord, list) else coord
                st.session_state.matched.append((tag, value, sheet_name, coord_display, title, reading_turbine))
            else:
                st.session_state.unmatched.append((tag, value, title, "tag absent du classeur", reading_turbine))

        st.session_state.processed_idx = i + 1
        progress_bar.progress((i + 1) / total)

    # --- construit un classeur separe par turbine detectee, a partir du meme modele ---
    by_turbine = {}
    for tag, value, sheet_name, coord_display, title, tb in st.session_state.matched:
        key = tb or "turbine_inconnue"
        by_turbine.setdefault(key, []).append((sheet_name, coord_display, value))

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        base_name = (st.session_state.excel_name or "rapport").replace(".xlsx", "")
        for tb, entries in by_turbine.items():
            wb = openpyxl.load_workbook(io.BytesIO(st.session_state.excel_bytes))
            for sheet_name, coord_display, value in entries:
                ws = wb[sheet_name]
                for c in coord_display.split('+'):
                    ws[c] = value
            buf = io.BytesIO()
            wb.save(buf)
            fname = f"{base_name}_GTG{tb}_{heure}.xlsx" if tb != "turbine_inconnue" else f"{base_name}_turbine_INCONNUE_{heure}.xlsx"
            zf.writestr(fname, buf.getvalue())

    st.session_state.final_bytes = zip_buf.getvalue()
    st.session_state.turbines_found = sorted(by_turbine.keys())
    st.session_state.run_state = "done"


if run_clicked:
    st.session_state["_photo_cache"] = photos
    run_batch(resume=False)

if resume_clicked:
    run_batch(resume=True)

# ---------------------------------------------------------------------------
# 5) Resultats
# ---------------------------------------------------------------------------
if st.session_state.run_state == "done":
    st.success("✅ Tournée terminée !")
    matched = st.session_state.matched
    unmatched = st.session_state.unmatched
    missing = sorted(set(st.session_state.sensor_map.keys()) - st.session_state.seen_tags)
    turbines_found = st.session_state.get("turbines_found", [])

    if auto_turbine:
        st.subheader(f"🏭 {len(turbines_found)} turbine(s) détectée(s) : {', '.join(turbines_found)}")

    st.subheader(f"✅ {len(matched)} valeur(s) écrite(s)")
    with st.expander("Voir le détail"):
        for tag, value, sheet, coord, title, tb in matched:
            tb_label = f"[GTG {tb}] " if tb else ""
            st.write(f"{tb_label}`{tag}` = {value} → {sheet}!{coord}  _(depuis: {title})_")

    if unmatched:
        st.subheader(f"⚠️ {len(unmatched)} valeur(s) à vérifier")
        with st.expander("Voir le détail"):
            for tag, value, title, reason, tb in unmatched:
                tb_label = f"[GTG {tb}] " if tb else ""
                st.write(f"{tb_label}`{tag}` = {value}  ({reason}, depuis: {title})")

    if missing:
        st.subheader(f"❓ {len(missing)} capteur(s)/champ(s) jamais vus (toutes turbines confondues)")
        with st.expander("Voir la liste"):
            st.write(", ".join(missing))

    zip_name = f"{(st.session_state.excel_name or 'rapport').replace('.xlsx', '')}_rempli_{heure}.zip"
    st.download_button(
        f"⬇️ Télécharger les {len(turbines_found)} rapport(s) remplis (.zip)",
        data=st.session_state.final_bytes,
        file_name=zip_name,
        mime="application/zip",
        type="primary",
    )
