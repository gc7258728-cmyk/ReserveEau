# -*- coding: utf-8 -*-
"""ReserveEau: bilan volumique discret et simulations pedagogiques, sans service externe."""
from __future__ import annotations

import copy
import hashlib
import html
import io
import json
import math
from datetime import date, timedelta
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
)
from reportlab.graphics.shapes import Drawing, Line, String, PolyLine
from scipy.stats import binomtest


TEAL = "#087E83"
CORAL = "#BD443F"
BLUE = "#2863B2"
INK = "#202D33"
GREEN = "#3A6C2C"
MAX_DAYS = 180
MAX_VOLUME = 10_000_000.0
NAV = [
    ("Tableau de bord", "dashboard"),
    ("Mon scénario", "tune"),
    ("Bilan quotidien", "table_chart"),
    ("Comparer les cuves", "compare_arrows"),
    ("Consommation et risques", "query_stats"),
    ("Réapprovisionnements", "local_shipping"),
    ("Rapports et sauvegardes", "download"),
    ("Méthode et tests", "school"),
]


def default_config():
    return dict(
        name="Mon bâtiment", start=date.today().isoformat(), capacity=3000.0,
        initial=3000.0, people=10, daily=50.0, days=14, threshold=20.0,
        low=40.0, high=60.0, deliveries=[], origin="Hypothèses personnelles",
    )


def example_config():
    c = default_config()
    c.update(name="Résidence des 10 étudiants", origin="Exemple fictif", days=14)
    return c


def number(value, name, low=0.0, high=MAX_VOLUME, integer=False):
    if isinstance(value, (bool, np.bool_)) or value is None:
        raise ValueError(f"{name} : un nombre est requis.")
    try:
        n = float(value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError(f"{name} : valeur numérique invalide.") from exc
    if not math.isfinite(n) or not low <= n <= high or (integer and n != int(n)):
        raise ValueError(f"{name} : valeur attendue entre {low:g} et {high:g}"
                         + (", entière." if integer else "."))
    return int(n) if integer else n


def validate_deliveries(rows, days):
    if not isinstance(rows, list) or len(rows) > MAX_DAYS:
        raise ValueError("Le plan doit comporter au plus 180 lignes.")
    grouped = {}
    for row in rows:
        if not isinstance(row, dict) or set(row) != {"jour", "volume_l"}:
            raise ValueError("Chaque livraison nécessite jour et volume_l.")
        d = number(row["jour"], "Jour de livraison", 1, days, True)
        v = number(row["volume_l"], "Volume livré (L)", 0.01)
        grouped[d] = grouped.get(d, 0.0) + v
        if grouped[d] > MAX_VOLUME:
            raise ValueError("Le total livré un même jour dépasse 10 000 000 L.")
    return [{"jour": k, "volume_l": v} for k, v in sorted(grouped.items())]


def validate_config(raw):
    if not isinstance(raw, dict) or set(raw) != set(default_config()):
        raise ValueError("Structure du scénario invalide ou incomplète.")
    c = copy.deepcopy(raw)
    for field in ("name", "origin"):
        if not isinstance(c[field], str) or not c[field].strip() or len(c[field]) > 100:
            raise ValueError("Nom ou origine invalide (1 à 100 caractères).")
        if any(ord(ch) < 32 or 0xD800 <= ord(ch) <= 0xDFFF for ch in c[field]):
            raise ValueError("Le nom et l’origine ne doivent pas contenir de caractères de contrôle.")
        c[field] = c[field].strip()
    try:
        d = date.fromisoformat(c["start"])
        if not date(2000, 1, 1) <= d <= date(2099, 1, 1):
            raise ValueError()
    except (TypeError, ValueError) as exc:
        raise ValueError("Date de départ invalide (2000 à 2099).") from exc
    for key, label, lo, hi, integer in [
        ("capacity", "Capacité", 1, MAX_VOLUME, False),
        ("initial", "Volume initial", 0, MAX_VOLUME, False),
        ("people", "Occupants", 1, 10000, True),
        ("daily", "Consommation habituelle", 0, 1000, False),
        ("days", "Horizon", 1, MAX_DAYS, True),
        ("threshold", "Seuil d'alerte", 0, 100, False),
        ("low", "Consommation minimale", 0, 1000, False),
        ("high", "Consommation maximale", 0, 1000, False),
    ]:
        c[key] = number(c[key], label, lo, hi, integer)
    if c["initial"] > c["capacity"]:
        raise ValueError("Le volume initial ne peut pas dépasser la capacité de la cuve.")
    if not c["low"] <= c["daily"] <= c["high"]:
        raise ValueError("La consommation doit respecter minimum ≤ habituelle ≤ maximum.")
    c["deliveries"] = validate_deliveries(c["deliveries"], c["days"])
    return c


def delivery_array(c):
    a = np.zeros(c["days"], dtype=float)
    for r in c["deliveries"]:
        a[r["jour"] - 1] += r["volume_l"]
    return a


def balance(c, demand=None):
    """Livraison en debut de jour, debordement, puis demande; pas de report du manque."""
    c = validate_config(c)
    if demand is None:
        demand = np.full(c["days"], c["people"] * c["daily"], dtype=float)
    demand = np.asarray(demand, dtype=float)
    if demand.shape != (c["days"],) or not np.isfinite(demand).all() or (demand < 0).any():
        raise ValueError("Une demande finie, positive ou nulle, est requise pour chaque jour.")
    supplies = delivery_array(c)
    level = c["initial"]
    rows = []
    for i, need in enumerate(demand):
        opening = level
        overflow = max(opening + supplies[i] - c["capacity"], 0.0)
        available = min(c["capacity"], opening + supplies[i])
        served = min(available, float(need))
        missing = max(float(need) - served, 0.0)
        level = max(available - served, 0.0)
        rows.append(dict(
            jour=i + 1, date=(date.fromisoformat(c["start"]) + timedelta(days=i)).isoformat(),
            debut_l=opening, livraison_l=supplies[i], debordement_l=overflow,
            demande_l=float(need), servie_l=served, manque_l=missing, fin_l=level,
            sous_seuil=level < c["capacity"] * c["threshold"] / 100,
        ))
    return pd.DataFrame(rows)


def summarize(df):
    missing = df["manque_l"] > 1e-8
    return dict(
        first_missing=int(df.loc[missing, "jour"].iloc[0]) if missing.any() else None,
        full_days=int((~missing).sum()), missing_days=int(missing.sum()),
        missing=float(df["manque_l"].sum()), overflow=float(df["debordement_l"].sum()),
        served=float(df["servie_l"].sum()), demand=float(df["demande_l"].sum()),
        final=float(df["fin_l"].iloc[-1]),
    )


def signature(c):
    return hashlib.sha256(json.dumps(c, sort_keys=True).encode()).hexdigest()


def probability_interval(hits, n):
    ci = binomtest(int(hits), int(n)).proportion_ci(confidence_level=0.95, method="exact")
    return float(ci.low), float(ci.high)


@st.cache_data(max_entries=6, ttl=1800, show_spinner=False)
def monte_carlo(c, n, seed):
    c = validate_config(c)
    n = number(n, "Nombre de simulations", 1000, 20000, True)
    seed = number(seed, "Graine", 0, 2**32 - 1, True)
    rng = np.random.default_rng(seed)
    levels = np.empty((n, c["days"]), dtype=np.float32)
    level = np.full(n, c["initial"], dtype=float)
    first = np.full(n, c["days"] + 1, dtype=np.int16)
    missing_days = np.zeros(n, dtype=np.int16)
    total_missing = np.zeros(n)
    total_demand = np.zeros(n)
    total_overflow = np.zeros(n)
    supplies = delivery_array(c)
    daily_shortage = []
    for j in range(c["days"]):
        per_person = (np.full(n, c["daily"]) if c["low"] == c["high"] else
                      rng.triangular(c["low"], c["daily"], c["high"], n))
        demand = per_person * c["people"]
        total_demand += demand
        total_overflow += np.maximum(level + supplies[j] - c["capacity"], 0)
        available = np.minimum(level + supplies[j], c["capacity"])
        missing = np.maximum(demand - available, 0)
        fails = missing > 1e-8
        first = np.where(fails & (first == c["days"] + 1), j + 1, first)
        missing_days += fails
        total_missing += missing
        level = np.maximum(available - demand, 0)
        levels[:, j] = level
        daily_shortage.append(float(fails.mean()))
    # Une seule matrice conservee; les trajectoires individuelles ne quittent pas la session.
    bands = np.quantile(levels, [0.05, 0.5, 0.95], axis=0)
    hits = int((first <= c["days"]).sum())
    lo, hi = probability_interval(hits, n)
    return dict(
        bands=bands, first=first, total_missing=total_missing,
        total_demand=total_demand, total_overflow=total_overflow,
        final=level, missing_days=missing_days, daily_shortage=np.array(daily_shortage),
        probability=hits / n, ci=(lo, hi), n=n, seed=seed,
        full_rate=float(1 - missing_days.sum() / (n * c["days"])),
    )


def comparison(c, capacities, fill_mode):
    rows, trajectories = [], []
    for capacity in capacities:
        v = number(capacity, "Capacité comparée", 1, MAX_VOLUME)
        cc = copy.deepcopy(c)
        cc["capacity"] = v
        if fill_mode == "Même taux de remplissage":
            cc["initial"] = v * c["initial"] / c["capacity"]
        elif c["initial"] > v:
            raise ValueError("À volume initial identique, chaque cuve doit contenir le volume initial.")
        df = balance(cc)
        ss = summarize(df)
        rows.append(dict(capacite_l=v, initial_l=cc["initial"],
                         premier_manque=ss["first_missing"], jours_couverts=ss["full_days"],
                         manque_l=ss["missing"], debordement_l=ss["overflow"], fin_l=ss["final"]))
        trajectories.append((f"{fmt(v)} L", df))
    return pd.DataFrame(rows), trajectories


def sensitivity(c):
    rows = []
    for label, key, factor in [
        ("Référence", "daily", 1), ("Consommation -10 %", "daily", .9),
        ("Consommation +10 %", "daily", 1.1),
        ("Capacité -10 %", "capacity", .9), ("Capacité +10 %", "capacity", 1.1),
    ]:
        cc = copy.deepcopy(c)
        cc[key] *= factor
        if key == "capacity":
            cc["initial"] *= factor
        else:
            cc["low"] *= factor
            cc["high"] *= factor
        # Limite d'interface seulement : les perturbations au-dela ne sont pas affichees.
        try:
            ss = summarize(balance(cc))
        except ValueError:
            continue
        rows.append(dict(variante=label, manque_l=ss["missing"], jours_couverts=ss["full_days"]))
    return pd.DataFrame(rows)


def read_delivery_csv(data, days):
    if len(data) > 1_000_000:
        raise ValueError("CSV trop volumineux : maximum 1 Mo.")
    try:
        df = pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding="utf-8-sig")
    except (ValueError, UnicodeError, pd.errors.ParserError) as exc:
        raise ValueError("CSV illisible : encodage UTF-8, séparateur virgule ou point-virgule.") from exc
    if set(df.columns) != {"jour", "volume_l"}:
        raise ValueError("Le CSV doit comporter exactement les colonnes jour et volume_l.")
    return validate_deliveries(df.to_dict("records"), days)


def read_backup(data):
    if len(data) > 100_000:
        raise ValueError("Sauvegarde trop volumineuse : maximum 100 Ko.")
    try:
        raw = json.loads(data)
    except (ValueError, UnicodeError) as exc:
        raise ValueError("Sauvegarde JSON illisible.") from exc
    if not isinstance(raw, dict) or set(raw) != {"format", "version", "scenario"}:
        raise ValueError("Sauvegarde RéserveEau attendue.")
    if raw["format"] != "reserveeau" or type(raw["version"]) is not int or raw["version"] != 1:
        raise ValueError("Version de sauvegarde non reconnue.")
    return validate_config(raw["scenario"])


def backup_bytes(c):
    return json.dumps(dict(format="reserveeau", version=1, scenario=c),
                      ensure_ascii=False, indent=2).encode("utf-8")


def csv_bytes(df):
    safe = df.copy()
    for col in safe.select_dtypes(include=["object", "str"]).columns:
        safe[col] = safe[col].map(lambda v: "'" + v if isinstance(v, str)
                                and v.lstrip().startswith(("=", "+", "-", "@")) else v)
    return safe.to_csv(index=False, sep=";", float_format="%.6f").encode("utf-8-sig")


def excel_bytes(c, df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        params = pd.DataFrame([dict(parametre=k, valeur=str(v)) for k, v in c.items()
                               if k != "deliveries"])
        for table, sheet in [(params, "Hypotheses"), (df, "Bilan quotidien"),
                             (pd.DataFrame(c["deliveries"], columns=["jour", "volume_l"]), "Livraisons")]:
            table.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for row in ws:
                for cell in row:
                    if isinstance(cell.value, str):
                        cell.data_type = "s"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(45, max(14, max(
                    len(str(cell.value or "")) for cell in col) + 2))
    return buffer.getvalue()


def pdf_bytes(c, df):
    """Rapport instantane : hypotheses, courbe vectorielle, bilan complet, sans simulation implicite."""
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("REtitle", fontName="Helvetica-Bold", fontSize=25,
                             leading=30, textColor=colors.HexColor(TEAL), spaceAfter=12))
    styles.add(ParagraphStyle("REsmall", fontSize=8, leading=11))
    styles["Heading2"].keepWithNext = True
    styles["Heading2"].textColor = colors.HexColor(TEAL)
    doc = SimpleDocTemplate(buffer, pagesize=(21*cm, 29.7*cm),
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=1.6*cm, bottomMargin=1.7*cm, title="RéserveEau - Bilan")
    ss = summarize(df)
    story = [Paragraph("RéserveEau", styles["REtitle"]),
             Paragraph(escape(c["name"]), styles["Heading2"]),
             Paragraph(f"{escape(c['origin'])} | Rapport du {date.today().strftime('%d/%m/%Y')}", styles["Normal"]),
             Spacer(1, 12)]
    facts = [
        ["Capacité", f"{fmt(c['capacity'])} L", "Volume initial", f"{fmt(c['initial'])} L"],
        ["Occupants", str(c["people"]), "Consommation", f"{fmt(c['daily'])} L/pers./j"],
        ["Horizon", f"{c['days']} jours", "Départ", c["start"]],
    ]
    t = Table(facts, colWidths=[3.2*cm, 5.4*cm, 3.3*cm, 5.5*cm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EDF6F5")),
                           ("FONTSIZE", (0, 0), (-1, -1), 9), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 9)]))
    story += [t, Spacer(1, 10), Paragraph("Résultats du bilan déterministe", styles["Heading2"]),
              Paragraph(f"Demande : {fmt(ss['demand'])} L. Eau servie : {fmt(ss['served'])} L. "
                        f"Manque : {fmt(ss['missing'])} L. Débordement : {fmt(ss['overflow'])} L. "
                        f"Volume final : {fmt(ss['final'])} L.", styles["Normal"]),
              Paragraph(f"Jours entièrement couverts : {ss['full_days']}/{c['days']}. "
                        + (f"Premier manque : J{ss['first_missing']}." if ss['first_missing']
                           else "Aucun manque sur l'horizon étudié."), styles["Normal"])]
    chart = Drawing(490, 175)
    x0, y0, w, h = 42, 30, 428, 125
    chart.add(Line(x0, y0, x0+w, y0, strokeColor=colors.grey))
    chart.add(Line(x0, y0, x0, y0+h, strokeColor=colors.grey))
    chart.add(String(3, y0+h-3, f"{fmt(c['capacity'])} L", fontSize=8))
    chart.add(String(22, y0-2, "0", fontSize=8))
    pts = []
    values = [c["initial"]] + df.fin_l.tolist()
    for i, v in enumerate(values):
        pts += [x0 + w*i/c["days"], y0 + h*v/c["capacity"]]
    chart.add(PolyLine(pts, strokeColor=colors.HexColor(TEAL), strokeWidth=2))
    chart.add(String(x0, 10, "J0", fontSize=8))
    chart.add(String(x0+w-20, 10, f"J{c['days']}", fontSize=8))
    story += [Spacer(1, 10), chart,
              Paragraph("Bilan quotidien (litres)", styles["Heading2"])]
    header = ["Jour", "Début", "Apport", "Débord.", "Demande", "Servie", "Manque", "Fin"]
    data = [header] + [[str(int(r.jour))] + [fmt(v) for v in [r.debut_l, r.livraison_l,
            r.debordement_l, r.demande_l, r.servie_l, r.manque_l, r.fin_l]] for r in df.itertuples()]
    t = Table(data, repeatRows=1, colWidths=[1.0*cm] + [2.35*cm]*7)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(TEAL)),
                           ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                           ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                           ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                           ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
                           ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F0F5F5")]),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
    story += [t, Spacer(1, 12), Paragraph("Hypothèses et limites", styles["Heading2"]),
              Paragraph("Livraisons en début de journée, débordement avant consommation, "
                        "demande uniforme par jour. Le manque n'est pas reporté au lendemain. "
                        "Les données sont des hypothèses, pas des mesures. Une fin de journée à zéro "
                        "ne signifie pas nécessairement un manque durant cette journée. "
                        "Ce bilan ne traite ni pression, ni pompage, ni potabilité, ni réserve incendie. "
                        "Il ne constitue pas un dimensionnement réglementaire.", styles["Normal"]),
              Paragraph("Les courbes sont déterministes. Les simulations probabilistes, si réalisées, "
                        "sont exportées séparément en CSV.", styles["REsmall"])]
    def footer(canvas, document):
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor(INK))
        canvas.drawString(1.5*cm, 1.0*cm, "RéserveEau | Outil pédagogique")
        canvas.drawRightString(19.5*cm, 1.0*cm, f"Page {document.page}")
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()


def fmt(v, decimals=0):
    return f"{float(v):,.{decimals}f}".replace(",", " ").replace(".", ",")


CSS = """
<style>
:root{--re-ink:#202D33;--re-teal:#087E83;--re-line:#D6E1E1}
.stApp{background:#F3F7F7;color:var(--re-ink);font-family:'Segoe UI',Arial,sans-serif;letter-spacing:0}
[data-testid=stAppViewContainer]{background:#F3F7F7}
[data-testid=stHeader]{background:rgba(243,247,247,.96)}
[data-testid=stMainBlockContainer]{max-width:1500px;padding-top:4.25rem;padding-bottom:3rem}
[data-testid=stCaptionContainer],[data-testid=stCaptionContainer] p{color:#52666B;opacity:1!important}
[data-testid=stSidebar]{background:#E4F0EE;border-right:1px solid #B9D4D0}
[data-testid=stSidebar] button{border-radius:5px;min-height:42px;text-align:left}
[data-testid=stSidebar] h2{font-size:23px!important}
.re-brand{display:flex;align-items:center;gap:14px;padding:12px 0 20px;border-bottom:3px solid #DC645C;margin-bottom:10px}
.re-mark{width:45px;height:45px;flex:0 0 45px;background:#087E83;border-radius:7px;display:grid;place-items:center;color:white;font-weight:700;font-size:22px}
.re-brand h1{font-size:34px;line-height:1.14;margin:0;color:#202D33;letter-spacing:0;padding:0}
.re-brand p{margin:4px 0 0;font-size:13px;color:#52666B}
.re-meta{display:flex;flex-wrap:wrap;gap:9px;align-items:center;margin:4px 0 14px;color:#43595E;font-size:13px}
.re-tag{padding:4px 8px;background:#DDECE8;border-left:3px solid #087E83}
.re-tag.fictive{background:#FAE7E3;border-color:#BD443F;color:#81312D}
.re-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:8px 0 20px}
.re-kpi{border:1px solid #D6E1E1;border-top:4px solid var(--accent);border-radius:6px;background:white;padding:15px 16px;min-width:0;min-height:122px}
.re-kpi small{display:block;color:#52666B;font-size:13px;margin-bottom:8px}
.re-kpi strong{display:block;color:#202D33;font-size:25px;line-height:1.2;overflow-wrap:anywhere}
.re-kpi span{display:block;font-size:12px;color:#52666B;margin-top:7px}
h2{font-size:24px!important;letter-spacing:0!important}h3{font-size:19px!important;letter-spacing:0!important}
[data-testid=stForm]{background:white;border:1px solid #CCDCDC;border-radius:6px}
[data-testid=stBaseButton-primaryFormSubmit],[data-testid=stBaseButton-primary]{background:#087E83;border-color:#087E83;color:white}
[data-testid=stBaseButton-primaryFormSubmit]:hover,[data-testid=stBaseButton-primary]:hover{background:#06676B;color:white;border-color:#06676B}
button{border-radius:5px!important}button p{white-space:normal!important;overflow-wrap:anywhere}
[data-baseweb=input],[data-baseweb=select]>div,[data-baseweb=textarea]{background:#FFFFFF;color:#202D33}
[data-testid=stNumberInput] input,[data-testid=stTextInput] input{color:#202D33}
[data-testid=stExpander]{background:white;border-radius:6px}
.re-tank-wrap{background:#E5F1F0;border-top:3px solid #087E83;padding:20px 16px 15px;min-height:365px;text-align:center;border-radius:0}
.re-tank-wrap h3{margin:0 0 12px;padding:0;font-size:18px;color:#202D33}
.re-tank{position:relative;width:172px;height:220px;border:3px solid #41777B;border-radius:18px 18px 8px 8px;margin:0 auto;overflow:hidden;background:#FAFCFC}
.re-water{position:absolute;bottom:0;left:0;right:0;background:#1D9EA8;border-top:3px solid #05686E;transition:height .5s ease}
.re-ticks{position:absolute;inset:0;background:repeating-linear-gradient(to top,transparent 0,transparent 43px,#41777B55 43px,#41777B55 44px);pointer-events:none}
.re-level{position:absolute;inset:0;display:grid;place-content:center;color:#132F36;text-shadow:0 1px 2px #fff;font-size:29px;font-weight:750;line-height:1.3}
.re-level small{font-size:14px;font-weight:600;background:#ffffffdd;padding:2px 6px;border-radius:3px}
.re-tank-footer{font-size:13px;color:#385A60;margin:14px 0 0}
.re-balance{display:flex;flex-wrap:wrap;gap:12px;background:#202D33;color:white;padding:16px 20px;margin:15px 0}
.re-balance div{flex:1 1 130px}.re-balance small{display:block;font-size:12px;color:#CAD7D8}.re-balance strong{font-size:19px;display:block;margin-top:5px}
[data-testid=stPlotlyChart]{background:white;border-radius:4px;overflow:hidden}
@media(max-width:1100px){.re-kpi strong{font-size:22px}.re-kpi{padding:12px}}
@media(max-width:700px){[data-testid=stMainBlockContainer]{padding:4.25rem 1rem 2rem}.re-kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.re-kpi strong{font-size:22px}.re-kpi{min-height:114px}.re-brand h1{font-size:29px}.re-brand{gap:10px}.re-tank-wrap{min-height:355px}.re-brand p{font-size:12px}}
@media(prefers-reduced-motion:reduce){.re-water{transition:none}}
</style>
"""


def kpis(items):
    st.html('<div class="re-kpis">' + "".join(
        f'<div class="re-kpi" style="--accent:{color}"><small>{html.escape(label)}</small>'
        f'<strong>{html.escape(str(value))}</strong><span>{html.escape(detail)}</span></div>'
        for label, value, detail, color in items) + "</div>")


def style_chart(fig, y="Volume (L)", height=365):
    fig.update_layout(
        height=height, template="plotly_white", paper_bgcolor="white", plot_bgcolor="white",
        font=dict(family="Arial, sans-serif", color=INK, size=12),
        margin=dict(l=12, r=16, t=20, b=16),
        legend=dict(orientation="h", yanchor="top", y=-.21, x=0, font=dict(size=11)),
        hovermode="x unified", xaxis_title="Jour", yaxis_title=y,
    )
    fig.update_xaxes(automargin=True, gridcolor="#EDF1F1", rangemode="tozero")
    fig.update_yaxes(automargin=True, gridcolor="#EDF1F1", rangemode="tozero")
    return fig


def line_figure(c, series):
    fig = go.Figure()
    palette = [TEAL, CORAL, BLUE, GREEN]
    for i, (label, df) in enumerate(series):
        start = c["initial"] if len(series) == 1 else float(df.debut_l.iloc[0])
        fig.add_trace(go.Scatter(x=[0] + df.jour.tolist(), y=[start] + df.fin_l.tolist(),
                                name=label, mode="lines+markers", line=dict(color=palette[i % 4], width=3),
                                marker=dict(size=4), hovertemplate="J%{x} : %{y:,.1f} L<extra>%{fullData.name}</extra>"))
    return style_chart(fig)


def chart(fig, key, alt):
    st.plotly_chart(fig, key=key, theme=None, config=dict(displayModeBar=False, responsive=True), alt=alt)


def table(df, key=None):
    st.dataframe(df, hide_index=True, key=key, height=min(450, 38+35*len(df)),
                 alt="Tableau des résultats de RéserveEau")


def download_csv(label, df, filename, key):
    st.download_button(label, csv_bytes(df), filename, "text/csv", key=key,
                       icon=":material/download:", on_click="ignore")


def save_current(c):
    st.session_state.workspaces[st.session_state.workspace] = validate_config(c)
    st.session_state.revision += 1
    st.session_state.pop("simulation", None)
    st.session_state.pop("reports", None)


def set_example():
    st.session_state.workspaces["Exemple fictif"] = example_config()
    st.session_state.workspace = "Exemple fictif"
    st.session_state.revision += 1
    st.session_state.pop("simulation", None)


def dashboard(c, df, key):
    ss = summarize(df)
    demand = c["people"] * c["daily"]
    auto = f"{fmt(c['initial']/demand, 1)} j" if demand else "Sans limite*"
    kpis([
        ("Réserve au départ", f"{fmt(c['initial'])} L", f"Cuve de {fmt(c['capacity'])} L", TEAL),
        ("Demande quotidienne", f"{fmt(demand)} L", f"{c['people']} occupants × {fmt(c['daily'])} L", BLUE),
        ("Autonomie sans apport", auto, "Rapport volume initial / demande", GREEN),
        ("Premier jour avec manque", f"J{ss['first_missing']}" if ss['first_missing'] else "Aucun", f"Sur {c['days']} jours étudiés", CORAL),
    ])
    left, right = st.columns([1, 2], gap="large")
    with left:
        view = st.slider("Niveau à la fin du jour", 0, c["days"], 0, key=key+"view")
        level = c["initial"] if view == 0 else float(df.fin_l.iloc[view - 1])
        percent = 100*level/c["capacity"]
        st.html(f'<div class="re-tank-wrap"><h3>État de la réserve · J{view}</h3>'
                f'<div class="re-tank" role="img" aria-label="Cuve remplie à {percent:.1f} pour cent">'
                f'<div class="re-water" style="height:{percent:.4f}%"></div><div class="re-ticks"></div>'
                f'<div class="re-level">{fmt(percent)} %<small>{fmt(level)} litres</small></div></div>'
                f'<p class="re-tank-footer">Capacité : {fmt(c["capacity"])} L · Schéma non à l’échelle</p></div>')
    with right:
        st.subheader("Trajectoire de la réserve")
        fig = line_figure(c, [("Fin de journée", df)])
        fig.add_hline(y=c["capacity"]*c["threshold"]/100, line_dash="dot", line_color=CORAL)
        chart(fig, key+"curve", "Volume en fin de journée et seuil d'alerte")
        st.caption(f"Trait pointillé : seuil de {fmt(c['threshold'])} % de la capacité. "
                   "J0 = volume initial ; les autres points sont les fins de journée.")
    st.html('<div class="re-balance">' + "".join(
        f'<div><small>{label}</small><strong>{fmt(val)} L</strong></div>' for label, val in [
            ("Eau demandée", ss["demand"]), ("Eau servie", ss["served"]),
            ("Demande non satisfaite", ss["missing"]), ("Eau débordée", ss["overflow"]),
        ]) + '</div>')
    if ss["missing_days"]:
        st.warning(f"Demande partiellement ou totalement non satisfaite sur {ss['missing_days']} jour(s). "
                   f"{ss['full_days']} journée(s) sont entièrement couvertes.", icon=":material/water_loss:")
    else:
        st.success("La demande est entièrement satisfaite sur l’horizon étudié.", icon=":material/check_circle:")
    if demand == 0:
        st.caption("* Demande nulle dans ce modèle ; aucune autonomie physique n'est certifiée.")
    st.subheader("Prochains apports du scénario")
    if c["deliveries"]:
        table(pd.DataFrame(c["deliveries"]).rename(columns={"jour": "Jour", "volume_l": "Volume (L)"}))
    else:
        st.caption("Aucun réapprovisionnement prévu.")


def scenario_page(c, key):
    st.subheader("Hypothèses du bâtiment")
    with st.form(key+"form"):
        name = st.text_input("Nom du scénario", c["name"], max_chars=100)
        a, b = st.columns(2)
        with a:
            capacity = st.number_input("Capacité de la cuve (L)", 1.0, MAX_VOLUME, c["capacity"], 100.0)
            initial = st.number_input("Volume disponible au départ (L)", 0.0, MAX_VOLUME, c["initial"], 100.0)
            people = st.number_input("Nombre d’occupants", 1, 10000, c["people"])
            daily = st.number_input("Consommation habituelle (L/personne/jour)", 0.0, 1000.0, c["daily"], 5.0)
        with b:
            start = st.date_input("Premier jour étudié", date.fromisoformat(c["start"]),
                                  min_value=date(2000, 1, 1), max_value=date(2099, 1, 1))
            days = st.number_input("Durée étudiée (jours)", 1, MAX_DAYS, c["days"])
            threshold = st.slider("Seuil d’alerte (% de capacité)", 0.0, 100.0, c["threshold"], 5.0)
            low = st.number_input("Consommation minimale simulée (L/personne/jour)", 0.0, 1000.0, c["low"], 5.0)
            high = st.number_input("Consommation maximale simulée (L/personne/jour)", 0.0, 1000.0, c["high"], 5.0)
        submitted = st.form_submit_button("Enregistrer le scénario", icon=":material/save:", type="primary")
    if submitted:
        cc = dict(c, name=name, capacity=capacity, initial=initial, people=people, daily=daily,
                  start=start.isoformat(), days=days, threshold=threshold, low=low, high=high)
        try:
            save_current(cc)
        except ValueError as exc:
            st.error(str(exc))
        else:
            st.session_state.flash = "Scénario enregistré dans cette session."
            st.rerun()
    st.caption("Les volumes sont des hypothèses de calcul. Les consommations minimale et maximale "
               "ne modifient pas le bilan déterministe. Les livraisons doivent rester dans la durée étudiée.")


def daily_page(c, df, key):
    st.subheader("De l’eau disponible à l’eau réellement servie")
    ss = summarize(df)
    kpis([("Jours entièrement couverts", str(ss["full_days"]), f"Sur {c['days']} jours", GREEN),
          ("Jours avec manque", str(ss["missing_days"]), "Manque strictement positif", CORAL),
          ("Volume final", f"{fmt(ss['final'])} L", "Fin du dernier jour", TEAL),
          ("Débordement cumulé", f"{fmt(ss['overflow'])} L", "Eau non stockable à l’arrivée", BLUE)])
    labels = dict(jour="Jour", date="Date", debut_l="Début (L)", livraison_l="Livraison (L)",
                  debordement_l="Débordement (L)", demande_l="Demande (L)", servie_l="Servie (L)",
                  manque_l="Manque (L)", fin_l="Fin (L)", sous_seuil="Sous le seuil")
    table(df.rename(columns=labels).round(2))
    fig = go.Figure()
    for col, label, color in [("servie_l", "Demande satisfaite", TEAL), ("manque_l", "Manque", CORAL)]:
        fig.add_trace(go.Bar(x=df.jour, y=df[col], name=label, marker_color=color))
    fig.update_layout(barmode="stack")
    chart(style_chart(fig, "Eau (L)"), key+"bars", "Part satisfaite et non satisfaite de la demande quotidienne")
    st.caption("Bilan : début + livraison = fin + eau servie + débordement. "
               "Demande = eau servie + manque. Le manque n’est pas reporté au lendemain.")
    download_csv("Bilan CSV", df, "reserveeau-bilan.csv", key+"csv")


def comparison_page(c, key):
    st.subheader("Trois cuves, les mêmes besoins")
    with st.form(key+"form"):
        a, b = st.columns(2)
        ca = a.number_input("Capacité alternative A (L)", 1.0, MAX_VOLUME, min(MAX_VOLUME, c["capacity"]*5/3), 100.0)
        cb = b.number_input("Capacité alternative B (L)", 1.0, MAX_VOLUME, min(MAX_VOLUME, c["capacity"]*7/3), 100.0)
        mode = st.selectbox("Conditions de départ", ["Même taux de remplissage", "Même volume initial"])
        submit = st.form_submit_button("Comparer les trois cuves", type="primary", icon=":material/compare_arrows:")
    if submit:
        st.session_state[key+"comparison"] = (ca, cb, mode)
    values = st.session_state.get(key+"comparison")
    if values:
        try:
            result, trajectories = comparison(c, [c["capacity"], values[0], values[1]], values[2])
        except ValueError as exc:
            st.error(str(exc))
            return
        chart(line_figure(c, trajectories), key+"plot", "Comparaison des volumes de trois cuves")
        display = result.copy()
        display["premier_manque"] = display["premier_manque"].map(
            lambda x: "Aucun" if pd.isna(x) else f"J{int(x)}")
        table(display.rename(columns={"capacite_l":"Capacité (L)", "initial_l":"Volume initial (L)",
              "premier_manque":"Premier manque", "jours_couverts":"Jours couverts",
              "manque_l":"Manque (L)", "debordement_l":"Débordement (L)", "fin_l":"Fin (L)"}))
        st.caption(values[2] + ". Consommation et livraisons identiques ; une plus grande cuve "
                   "ne crée pas d’eau. À taux de remplissage identique, elle contient davantage d’eau au départ.")
        download_csv("Comparaison CSV", result, "reserveeau-comparaison.csv", key+"csv")


def risk_page(c, key):
    st.subheader("Et si la consommation variait ?")
    with st.form(key+"form"):
        a, b = st.columns(2)
        n = a.select_slider("Nombre de scénarios simulés", [1000, 5000, 10000, 20000], value=10000)
        seed = b.number_input("Graine reproductible", 0, 2**32-1, 42)
        submit = st.form_submit_button("Simuler les consommations", type="primary", icon=":material/play_arrow:")
    st.caption(f"Loi triangulaire : minimum {fmt(c['low'])}, valeur habituelle {fmt(c['daily'])}, "
               f"maximum {fmt(c['high'])} L/personne/jour. Un tirage par bâtiment et par jour, "
               "multiplié par le nombre d’occupants. Les jours sont supposés indépendants ; livraisons certaines.")
    if submit:
        with st.spinner("Calcul des trajectoires…"):
            st.session_state.simulation = (signature(c), monte_carlo(c, n, seed))
    stored = st.session_state.get("simulation")
    if not stored or stored[0] != signature(c):
        return
    r = stored[1]
    kpis([
        ("Risque d’au moins un manque", f"{fmt(100*r['probability'], 2)} %", f"Sur {c['days']} jours", CORAL),
        ("Journées entièrement servies", f"{fmt(100*r['full_rate'], 2)} %", "Toutes les journées simulées", GREEN),
        ("Manque cumulé médian", f"{fmt(np.median(r['total_missing']))} L", "Quantile 50 % des trajectoires", BLUE),
        ("Volume final médian", f"{fmt(np.median(r['final']))} L", f"{r['n']:,} scénarios".replace(",", " "), TEAL),
    ])
    st.caption(f"Intervalle binomial exact à 95 % du risque : [{fmt(100*r['ci'][0], 3)} % ; "
               f"{fmt(100*r['ci'][1], 3)} %]. Il décrit l’incertitude de simulation, pas celle du modèle.")
    fig = go.Figure()
    x = list(range(1, c["days"]+1))
    fig.add_trace(go.Scatter(x=x, y=r["bands"][2], mode="lines", line=dict(width=0),
                             name="Quantile 95 %", showlegend=False, hovertemplate="J%{x} · Q95 : %{y:.1f} L<extra></extra>"))
    fig.add_trace(go.Scatter(x=x, y=r["bands"][0], fill="tonexty", fillcolor="#C8E5E3", mode="lines",
                             line=dict(width=0), name="Bande Q05–Q95",
                             hovertemplate="J%{x} · Q05 : %{y:.1f} L<extra></extra>"))
    fig.add_trace(go.Scatter(x=x, y=r["bands"][1], name="Médiane", line=dict(color=TEAL, width=3)))
    chart(style_chart(fig), key+"bands", "Médiane et quantiles ponctuels 5 à 95 pour cent du stock")
    st.caption("Quantiles calculés séparément chaque jour : cette bande n’est pas une garantie sur une trajectoire complète.")
    prob = go.Figure(go.Bar(x=x, y=100*r["daily_shortage"], name="Manque le jour considéré", marker_color=CORAL))
    prob.update_yaxes(range=[0, 100])
    chart(style_chart(prob, "Probabilité (%)", 310), key+"prob", "Probabilité de demande non satisfaite pour chaque jour")
    first = r["first"].astype(float)
    first[first > c["days"]] = np.nan
    output = pd.DataFrame(dict(simulation=np.arange(1, r["n"]+1), premier_jour_manque=first,
                               jours_avec_manque=r["missing_days"], demande_totale_l=r["total_demand"],
                               manque_total_l=r["total_missing"], debordement_total_l=r["total_overflow"],
                               volume_final_l=r["final"], graine=r["seed"]))
    download_csv("Simulations CSV", output, "reserveeau-simulations.csv", key+"csv")
    st.subheader("Sensibilité autour du bilan déterministe")
    sens = sensitivity(c)
    f = go.Figure(go.Bar(x=sens.manque_l, y=sens.variante, orientation="h",
                        marker_color=[INK, TEAL, CORAL, BLUE, GREEN]))
    f = style_chart(f, "Variante", 310)
    f.update_layout(xaxis_title="Manque cumulé (L)", showlegend=False)
    chart(f, key+"sensitivity", "Effet de variations de consommation et de capacité sur le manque")
    st.caption("Une hypothèse varie à la fois. Le taux de remplissage initial est conservé lorsque "
               "la capacité change. Cette comparaison n’est pas une preuve de causalité.")


def supplies_page(c, key):
    st.subheader("Plan de livraison")
    with st.form(key+"add", clear_on_submit=True):
        a, b = st.columns(2)
        day = a.number_input("Jour de livraison", 1, c["days"], min(4, c["days"]))
        volume = b.number_input("Volume livré (L)", 0.01, MAX_VOLUME, 1000.0, 100.0)
        add = st.form_submit_button("Ajouter la livraison", icon=":material/add:", type="primary")
    if add:
        try:
            save_current(dict(c, deliveries=c["deliveries"] + [dict(jour=day, volume_l=volume)]))
        except ValueError as exc:
            st.error(str(exc))
        else:
            st.session_state.flash = "Livraison enregistrée. Les apports d’un même jour sont regroupés."
            st.rerun()
    if c["deliveries"]:
        with st.form(key+"edit"):
            edited = st.data_editor(pd.DataFrame(c["deliveries"]), hide_index=True,
                                    num_rows="dynamic", key=key+"editor",
                                    column_config={"jour": st.column_config.NumberColumn("Jour", min_value=1, max_value=c["days"], step=1, required=True),
                                                   "volume_l": st.column_config.NumberColumn("Volume (L)", min_value=.01, max_value=MAX_VOLUME, required=True)},
                                    alt="Livraisons modifiables")
            edit = st.form_submit_button("Enregistrer le tableau", icon=":material/save:")
        if edit:
            try:
                rows = edited.dropna(how="all").to_dict("records")
                save_current(dict(c, deliveries=rows))
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.rerun()
    else:
        st.caption("Aucune livraison dans ce scénario.")
    with st.expander("Importer un plan CSV", icon=":material/upload:"):
        st.download_button("Modèle de livraisons CSV", b"jour;volume_l\n3;1000\n6;1000\n", "modele-livraisons.csv",
                           "text/csv", icon=":material/download:", on_click="ignore")
        up = st.file_uploader("Plan CSV", type=["csv"], key=key+"csvup")
        confirm = st.checkbox("Remplacer le plan actuel par ce CSV", key=key+"confirm")
        if st.button("Importer les livraisons", disabled=not (up and confirm), icon=":material/upload:"):
            try:
                save_current(dict(c, deliveries=read_delivery_csv(up.getvalue(), c["days"])))
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.rerun()
    st.subheader("Tester un rythme régulier")
    with st.form(key+"periodic"):
        a, b, d = st.columns(3)
        first = a.number_input("Première livraison (jour)", 1, c["days"], min(4, c["days"]))
        every = b.number_input("Intervalle (jours)", 1, MAX_DAYS, 3)
        amount = d.number_input("Quantité par livraison (L)", 0.01, MAX_VOLUME, 1500.0, 100.0)
        trial = st.form_submit_button("Évaluer ce rythme", icon=":material/analytics:", type="primary")
    if trial:
        st.session_state[key+"proposal"] = [dict(jour=j, volume_l=amount)
                                            for j in range(first, c["days"]+1, every)]
    proposal = st.session_state.get(key+"proposal")
    if proposal is not None:
        cc = dict(c, deliveries=proposal)
        ss = summarize(balance(cc))
        st.caption(f"Plan proposé : {len(proposal)} livraison(s), {fmt(sum(x['volume_l'] for x in proposal))} L "
                   f"livrés ; {fmt(ss['missing'])} L de manque et {fmt(ss['overflow'])} L de débordement.")
        chart(line_figure(c, [("Plan actuel", balance(c)), ("Plan proposé", balance(cc))]),
              key+"proposalplot", "Comparaison du plan de livraison actuel et proposé")
        if st.button("Appliquer ce plan de livraison", icon=":material/check:"):
            save_current(cc)
            st.session_state.flash = "Le plan proposé remplace les anciennes livraisons."
            st.rerun()


def reports_page(c, df, key):
    st.subheader("Résultats et sauvegarde du scénario")
    if st.button("Préparer le rapport et le classeur", type="primary", icon=":material/description:"):
        st.session_state.reports = (signature(c), date.today().isoformat(), pdf_bytes(c, df), excel_bytes(c, df))
    reports = st.session_state.get("reports")
    if reports and reports[:2] == (signature(c), date.today().isoformat()):
        with st.container(horizontal=True):
            st.download_button("Rapport PDF", reports[2], "reserveeau-rapport.pdf", "application/pdf",
                               icon=":material/picture_as_pdf:", on_click="ignore")
            st.download_button("Classeur Excel", reports[3], "reserveeau-bilan.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               icon=":material/table_chart:", on_click="ignore")
    with st.container(horizontal=True):
        download_csv("Bilan CSV", df, "reserveeau-bilan.csv", key+"csv")
        st.download_button("Sauvegarde JSON", backup_bytes(c), "reserveeau-scenario.json", "application/json",
                           icon=":material/save:", on_click="ignore")
    st.caption("Les saisies restent dans la session du navigateur. Le JSON conserve les hypothèses et les "
               "livraisons, pas les résultats de simulation. Le PDF et le classeur présentent le bilan déterministe.")
    st.subheader("Restaurer un scénario")
    up = st.file_uploader("Sauvegarde RéserveEau (JSON)", type=["json"], key=key+"restore")
    confirm = st.checkbox("Remplacer uniquement le scénario de cet espace", key=key+"confirm")
    if st.button("Restaurer le scénario", disabled=not (up and confirm), icon=":material/restore:"):
        try:
            save_current(read_backup(up.getvalue()))
        except ValueError as exc:
            st.error(str(exc))
        else:
            st.session_state.flash = "Scénario restauré. L’autre espace est inchangé."
            st.rerun()


def method_page():
    st.subheader("Le modèle, simplement")
    st.markdown("**1. L’eau arrive.** On ajoute les livraisons au volume de départ de la journée. "
                "L’excédent au-delà de la capacité déborde.\n\n"
                "**2. Les occupants consomment.** La demande vaut le nombre d’occupants multiplié par "
                "la consommation par personne. On ne peut servir que l’eau disponible.\n\n"
                "**3. On fait le bilan.** Le stock restant devient le volume de départ du lendemain. "
                "Une demande non satisfaite n’est pas reportée.")
    st.latex(r"D=Nq\quad;\quad V_{disp}=\min(C,V_{debut}+A)")
    st.latex(r"E=\min(D,V_{disp})\quad;\quad M=D-E\quad;\quad V_{fin}=V_{disp}-E")
    st.caption("C : capacité ; A : apport ; D : demande ; E : eau servie ; M : manque. Tous les volumes sont en litres.")
    st.subheader("Tests rapides à reproduire")
    table(pd.DataFrame([
        {"Test":"Sans livraison", "Saisie":"3 000 L, 10 personnes, 50 L/pers./j, 14 jours", "Attendu":"500 L/j ; J6 à zéro ; premier manque J7 ; manque total 4 000 L"},
        {"Test":"Livraison", "Saisie":"Même cas + 1 000 L au début de J3", "Attendu":"2 500 L à la fin de J3 ; premier manque J9 ; manque 3 000 L"},
        {"Test":"Débordement", "Saisie":"Même cas + 1 000 L au début de J1", "Attendu":"1 000 L débordés ; 2 500 L en fin de J1 ; premier manque J7"},
        {"Test":"Comparaison", "Saisie":"Cuves pleines 3 000, 5 000, 7 000 L ; aucun apport ; 14 jours", "Attendu":"6, 10 et 14 jours couverts ; premier manque J7, J11, aucun"},
        {"Test":"Simulation constante", "Saisie":"Minimum = habituelle = maximum = 50 ; exemple sans apport", "Attendu":"100 % des trajectoires avec manque ; premier manque J7"},
        {"Test":"Plan régulier", "Saisie":"1 500 L à J4, puis tous les 3 jours ; exemple de 14 jours", "Attendu":"4 livraisons, 6 000 L apportés ; aucun manque ; 0 débordement ; fin 2 000 L"},
    ]))
    st.subheader("Ce que les probabilités signifient")
    st.markdown("La loi triangulaire utilise un minimum, une valeur habituelle (le **mode**) et un maximum. "
                "Sa moyenne vaut leur somme divisée par trois : elle n’est donc pas toujours égale à la valeur habituelle. "
                "Les tirages quotidiens sont indépendants, mais chaque tirage représente une consommation "
                "moyenne commune au bâtiment. Les livraisons ne sont pas aléatoires dans ce modèle.")
    st.subheader("Limites")
    st.markdown("Aucune mesure réelle n’est ajoutée automatiquement. Ni les variations horaires, ni les fuites, "
                "ni l’évaporation, ni les retards de livraison ne sont modélisés. Une livraison annoncée "
                "le matin peut rendre le bilan optimiste si elle arrive en réalité le soir. "
                "L’autonomie sans apport est un rapport théorique à consommation constante. "
                "Un pas journalier ne permet pas de dater une panne à l’heure près. "
                "L’application ne vérifie ni pression, ni pompage, ni potabilité, ni réserve incendie, "
                "et ne remplace pas une étude de dimensionnement.")


def main():
    st.set_page_config(page_title="RéserveEau", page_icon=":material/water_drop:", layout="wide")
    st.html(CSS)
    st.session_state.setdefault("workspaces", {"Mon scénario": default_config(), "Exemple fictif": example_config()})
    st.session_state.setdefault("workspace", "Mon scénario")
    st.session_state.setdefault("revision", 0)
    st.session_state.setdefault("page", "Tableau de bord")
    with st.sidebar:
        st.header("RéserveEau", icon=":material/water_drop:")
        st.segmented_control("Espace de travail", ["Mon scénario", "Exemple fictif"],
                              key="workspace", selection_mode="single", required=True)
        for title, icon in NAV:
            if st.button(title, icon=f":material/{icon}:", width="stretch", key="nav_"+icon,
                         type="primary" if st.session_state.page == title else "secondary"):
                st.session_state.page = title
                st.rerun()
        st.caption("RéserveEau 1.0 · Génie civil\n\nOutil pédagogique · Sans connexion")
    c = validate_config(st.session_state.workspaces[st.session_state.workspace])
    key = st.session_state.workspace + str(st.session_state.revision) + "_"
    st.html('<header class="re-brand"><div class="re-mark" aria-hidden="true">R</div><div>'
            '<h1>RéserveEau</h1><p>EAU DISPONIBLE · BESOINS · AUTONOMIE</p></div></header>')
    tag = "fictive" if c["origin"] == "Exemple fictif" else ""
    st.html(f'<div class="re-meta"><span class="re-tag {tag}">{html.escape(c["origin"])}</span>'
            f'<span>{html.escape(c["name"])}</span><span>· {c["days"]} jours à partir du '
            f'{date.fromisoformat(c["start"]).strftime("%d/%m/%Y")}</span></div>')
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))
    if st.session_state.page == "Tableau de bord":
        st.button("Charger l’exemple fictif", icon=":material/science:", on_click=set_example)
    df = balance(c)
    page = st.session_state.page
    if page != "Tableau de bord":
        st.header(page)
    if page == "Tableau de bord":
        dashboard(c, df, key)
    elif page == "Mon scénario":
        scenario_page(c, key)
    elif page == "Bilan quotidien":
        daily_page(c, df, key)
    elif page == "Comparer les cuves":
        comparison_page(c, key)
    elif page == "Consommation et risques":
        risk_page(c, key)
    elif page == "Réapprovisionnements":
        supplies_page(c, key)
    elif page == "Rapports et sauvegardes":
        reports_page(c, df, key)
    elif page == "Méthode et tests":
        method_page()


if __name__ == "__main__":
    main()
