"""Gestione Strumenti e Attrezzature - Streamlit + Google Sheets."""
import hmac
import io
import re
import time
from datetime import date, datetime
from html import escape
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from dateutil.relativedelta import relativedelta
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Strumenti e Attrezzature", page_icon="🔧", layout="wide")

COLS = {
    "Strumenti": ["uid", "ID", "Descrizione", "Modello", "Marca", "Seriale", "Padre", "Figlio",
                  "Tipologia", "In carico a", "Data assegnazione", "Assegnato fino al", "Stato",
                  "Note stato", "Data arrivo", "Ultima calibrazione", "Scadenza (mesi)",
                  "Prossima calibrazione", "Certificato"],
    "Tecnici": ["Nome"],
    "Richieste": ["id", "Data", "Tecnico", "Tipo", "uid", "Attrezzatura", "Messaggio", "Stato"],
    "Note": ["uid", "Data", "Autore", "Testo"],
}
STATI = ["OK", "Rotto", "In taratura", "Da verificare", "Non indicato"]
TIPI = ["Strumento", "Attrezzo", "Veicolo"]
MAG = "Mag18"
TZ = ZoneInfo("Europe/Rome")


# ----------------------------------------------------------------- utilità
def now():
    return datetime.now(TZ)


def today():
    return now().date()


def parse(s):
    try:
        return date.fromisoformat(str(s)[:10])
    except ValueError:
        return None


def fmt(s):
    d = parse(s)
    return d.strftime("%d/%m/%Y") if d else ""


def days_to_cal(r):
    """Giorni alla prossima taratura (None se non applicabile o strumento rotto)."""
    d = parse(r["Prossima calibrazione"])
    if r["Stato"] == "Rotto" or not d:
        return None
    return (d - today()).days


def sem(d):
    if d is None or pd.isna(d):
        return ""
    return "🔴" if d < 0 else "🟠" if d <= 30 else "🟡" if d <= 90 else "🟢"


def is_mag(s):
    s = str(s).strip().lower()
    return s == "" or s.startswith("mag")


def mask_assigned(df):
    """Strumenti in carico a un tecnico (non in magazzino)."""
    return ~df["In carico a"].map(is_mag)


def mask_temp(df):
    """Assegnazioni temporanee: in carico a un tecnico con data di fine."""
    return mask_assigned(df) & df["Assegnato fino al"].str.strip().ne("")


def add_temp_cols(t):
    """Aggiunge Dal / Fino al / Durata / Residuo alle assegnazioni temporanee."""
    t = t.copy()
    a = t["Data assegnazione"].map(parse)
    b = t["Assegnato fino al"].map(parse)
    t["Dal"] = t["Data assegnazione"].map(fmt)
    t["Fino al"] = t["Assegnato fino al"].map(fmt)
    t["Durata"] = [f"{(y - x).days} gg" if x and y else "" for x, y in zip(a, b)]

    def resid(y):
        if not y:
            return ""
        n = (y - today()).days
        return f"scaduta da {-n} gg" if n < 0 else "scade oggi" if n == 0 else f"{n} gg"
    t["Residuo"] = [resid(y) for y in b]
    return t


ID_RE = re.compile(r"^\s*([A-Za-z])\s*[-_./ ]?\s*(\d+)")


def next_ids(df, letters="SAGK"):
    """Per ogni lettera: (ultimo numero usato, prossimo libero, cifre)."""
    out = {}
    for L in letters:
        nums, width = [], 3
        for v in df["ID"]:
            m = ID_RE.match(v)
            if m and m.group(1).upper() == L:
                nums.append(int(m.group(2)))
                width = max(width, len(m.group(2)))
        last = max(nums) if nums else 0
        out[L] = (last, last + 1, width)
    return out


# ------------------------------------------------------------ Google Sheets
@st.cache_resource
def book():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return gspread.authorize(creds).open_by_key(st.secrets["SHEET_ID"])


@st.cache_data(ttl=15, show_spinner=False)
def load(tab):
    rows = book().worksheet(tab).get_all_records(numericise_ignore=["all"], default_blank="")
    return pd.DataFrame(rows, columns=COLS[tab]).fillna("").astype(str)


def append(tab, rec):
    book().worksheet(tab).append_row([rec.get(c, "") for c in COLS[tab]], value_input_option="RAW")
    load.clear()


def update_row(tab, key, changes):
    """Aggiorna le celle di una riga, cercando la chiave nella prima colonna."""
    ws = book().worksheet(tab)
    cell = ws.find(str(key), in_column=1)
    if cell is None:
        raise ValueError("Riga non trovata: ricarica la pagina.")
    cols = COLS[tab]
    data = [{"range": gspread.utils.rowcol_to_a1(cell.row, cols.index(k) + 1), "values": [[v]]}
            for k, v in changes.items()]
    ws.batch_update(data, value_input_option="RAW")
    load.clear()


def techs():
    return [n for n in load("Tecnici")["Nome"].tolist() if n]


def notes_text():
    n = load("Note")
    return n.groupby("uid")["Testo"].apply(" ".join) if not n.empty else pd.Series(dtype=str)


# ------------------------------------------------------------------- login
def login():
    st.title("🔧 Strumenti e Attrezzature")
    role = st.radio("Accedi come", ["Tecnico", "Ufficio"], horizontal=True)
    with st.form("login"):
        nome = None
        if role == "Tecnico":
            nome = st.selectbox("Il tuo nome", techs(), index=None, placeholder="Scegli il tuo nome")
        pw = st.text_input("Password", type="password")
        ok = st.form_submit_button("Accedi")
    if ok:
        if not hmac.compare_digest(pw, str(st.secrets["passwords"][role.lower()])):
            st.error("Password errata.")
        elif role == "Tecnico" and not nome:
            st.error("Scegli il tuo nome dall’elenco.")
        else:
            st.session_state.me = {"role": role.lower(), "name": nome or "Ufficio"}
            st.session_state.show_popup = role == "Ufficio"
            r = load("Richieste")
            st.session_state.last_new = int((r["Stato"] == "nuova").sum())
            st.rerun()


@st.dialog("🔔 Notifiche")
def popup():
    r = load("Richieste")
    n = r[r["Stato"] == "nuova"]
    if n.empty:
        st.write("Nessuna nuova notifica.")
    else:
        st.write(f"**{len(n)}** da leggere:")
        for tec, c in n.groupby("Tecnico").size().items():
            st.write(f"👤 **{tec}**: {c} {'messaggio' if c == 1 else 'messaggi'}")
    if st.button("Chiudi"):
        st.rerun()


@st.fragment(run_every=30)
def watcher():
    """Ogni 30 secondi controlla se ci sono nuove richieste (solo ufficio)."""
    n = int((load("Richieste")["Stato"] == "nuova").sum())
    if n > st.session_state.get("last_new", 0):
        st.toast(f"{n} richieste da leggere: apri la sezione Richieste", icon="🔔")
    st.session_state.last_new = n


# --------------------------------------------------------------- inventario
def decorate(df):
    df = df.copy()
    df["_d"] = pd.to_numeric(df.apply(days_to_cal, axis=1), errors="coerce")
    df["Taratura"] = df.apply(lambda r: (sem(r["_d"]) + " " + fmt(r["Prossima calibrazione"])).strip(), axis=1)
    return df


def inventory(df, key, me):
    if df.empty:
        st.info("Nessuna voce.")
        return
    df = decorate(df)
    q = st.text_input("🔍 Cerca (ID, descrizione, marca, seriale, note…)", key=f"q_{key}")
    c1, c2, c3, c4 = st.columns(4)
    tipo = c1.selectbox("Tipo", ["Tutti"] + TIPI, key=f"t_{key}")
    stato = c2.selectbox("Stato", ["Tutti"] + STATI, key=f"s_{key}")
    carico = c3.selectbox("In carico a", ["Tutti"] + sorted(x for x in df["In carico a"].unique() if x), key=f"c_{key}")
    tar = c4.selectbox("Taratura", ["Tutte", "Scadute", "Scadute + entro 30 gg"], key=f"x_{key}")
    if tipo != "Tutti":
        df = df[df["Tipologia"] == tipo]
    if stato != "Tutti":
        df = df[df["Stato"] == stato]
    if carico != "Tutti":
        df = df[df["In carico a"] == carico]
    if tar == "Scadute":
        df = df[df["_d"] < 0]
    elif tar != "Tutte":
        df = df[df["_d"] <= 30]
    if q.strip():
        hay = (df["ID"] + " " + df["Descrizione"] + " " + df["Modello"] + " " + df["Marca"] + " " +
               df["Seriale"] + " " + df["In carico a"] + " " + df["uid"].map(notes_text()).fillna("")).str.lower()
        for w in q.lower().split():
            df = df[hay.loc[df.index].str.contains(w, regex=False)]
    df = df.reset_index(drop=True)
    st.caption(f"{len(df)} risultati · seleziona una riga per aprire la scheda")
    cols = ["ID", "Descrizione", "Marca", "Modello", "In carico a", "Stato", "Taratura"]
    ev = st.dataframe(df[cols], hide_index=True, on_select="rerun", selection_mode="single-row", key=f"tbl_{key}")
    if ev.selection.rows:
        detail(df.iloc[ev.selection.rows[0]], me, key)


def table_with_detail(df, cols, key, me):
    """Tabella selezionabile: cliccando una riga si apre la scheda."""
    df = df.reset_index(drop=True)
    if df.empty:
        st.info("Nessuna voce.")
        return
    ev = st.dataframe(df[cols], hide_index=True, on_select="rerun", selection_mode="single-row", key=f"tbl_{key}")
    if ev.selection.rows:
        detail(df.iloc[ev.selection.rows[0]], me, key)


def detail(r, me, key):
    uid = r["uid"]
    st.divider()
    st.subheader(f'{r["ID"]} — {r["Descrizione"]}')
    c1, c2, c3 = st.columns(3)
    c1.markdown(f'**Marca/Modello:** {r["Marca"]} {r["Modello"]}  \n**Seriale:** {r["Seriale"] or "—"}  \n**Tipo:** {r["Tipologia"]}')
    c2.markdown(f'**In carico a:** {r["In carico a"] or "—"}  \n**Dal:** {fmt(r["Data assegnazione"]) or "—"}  \n**Fino al:** {fmt(r["Assegnato fino al"]) or "—"}')
    c3.markdown(f'**Stato:** {r["Stato"]}  \n**Prossima taratura:** {r["Taratura"] or "—"}  \n**Certificato:** {r["Certificato"] or "—"}')

    # ---- storico note (accumulabili)
    notes = load("Note")
    notes = notes[notes["uid"] == uid]
    st.markdown(f"**📝 Storico note ({len(notes)})**")
    for n in notes.iloc[::-1].itertuples():
        with st.container(border=True):
            st.caption(f'{n.Data or "Importata da Excel"} · {n.Autore}')
            st.write(n.Testo)
    with st.form(f"nota_{key}_{uid}", clear_on_submit=True):
        testo = st.text_area("Aggiungi nota")
        if st.form_submit_button("Aggiungi nota") and testo.strip():
            append("Note", {"uid": uid, "Data": now().strftime("%Y-%m-%d %H:%M"), "Autore": me["name"], "Testo": testo.strip()})
            st.rerun()

    if me["role"] == "tecnico":
        if st.button("📨 Richiedi questa attrezzatura", key=f"rq_{key}_{uid}"):
            append("Richieste", {"id": f"r{int(time.time() * 1000)}", "Data": now().strftime("%Y-%m-%d %H:%M"),
                                 "Tecnico": me["name"], "Tipo": "Richiesta attrezzatura", "uid": uid,
                                 "Attrezzatura": f'{r["ID"]} · {r["Descrizione"]}', "Messaggio": "", "Stato": "nuova"})
            st.success("Richiesta inviata all’ufficio.")
        return

    # ---- solo ufficio: assegnazione e modifica
    with st.expander("👤 Assegnazione temporanea"):
        a1, a2 = st.columns(2)
        tec = a1.selectbox("Tecnico", techs(), index=None, key=f"as_{key}_{uid}")
        fino = a2.date_input("Fino al (facoltativo)", value=None, format="DD/MM/YYYY", key=f"fi_{key}_{uid}")
        b1, b2 = st.columns(2)
        if b1.button("Assegna", key=f"ab_{key}_{uid}") and tec:
            update_row("Strumenti", uid, {"In carico a": tec, "Data assegnazione": today().isoformat(),
                                          "Assegnato fino al": fino.isoformat() if fino else ""})
            st.rerun()
        if b2.button("Rientro in magazzino", key=f"rb_{key}_{uid}"):
            update_row("Strumenti", uid, {"In carico a": MAG, "Data assegnazione": today().isoformat(), "Assegnato fino al": ""})
            st.rerun()
    with st.expander("✏️ Modifica scheda"):
        with st.form(f"mod_{key}_{uid}"):
            ch = tool_fields(r)
            if st.form_submit_button("Salva modifiche"):
                update_row("Strumenti", uid, ch())
                st.rerun()


def tool_fields(r=None):
    """Campi del modulo strumento. Restituisce una funzione che legge i valori."""
    r = r if r is not None else {c: "" for c in COLS["Strumenti"]}
    idx = lambda lst, v: lst.index(v) if v in lst else 0
    a, b = st.columns(2)
    i = a.text_input("ID (es. S-075)", r["ID"])
    tipo = b.selectbox("Tipologia", TIPI, index=idx(TIPI, r["Tipologia"]))
    d = st.text_input("Descrizione", r["Descrizione"])
    a, b = st.columns(2)
    mod = a.text_input("Modello", r["Modello"])
    marca = b.text_input("Marca", r["Marca"])
    sn = st.text_input("Numero seriale", r["Seriale"])
    hold = [MAG] + techs()
    a, b = st.columns(2)
    car = a.selectbox("In carico a", hold, index=idx(hold, r["In carico a"] or MAG))
    stato = b.selectbox("Stato", STATI, index=idx(STATI, r["Stato"] or "OK"))
    a, b = st.columns(2)
    uc = a.date_input("Ultima calibrazione", value=parse(r["Ultima calibrazione"]), min_value=date(2000, 1, 1), format="DD/MM/YYYY")
    ms = b.number_input("Validità certificato (mesi)", min_value=0, step=1,
                        value=int(float(r["Scadenza (mesi)"])) if r["Scadenza (mesi)"] else 0)
    cert = st.text_input("Certificato di taratura", r["Certificato"])

    def read():
        pc = (uc + relativedelta(months=int(ms))).isoformat() if uc and ms else r["Prossima calibrazione"]
        return {"ID": i.strip(), "Tipologia": tipo, "Descrizione": d.strip(), "Modello": mod.strip(),
                "Marca": marca.strip(), "Seriale": sn.strip(), "In carico a": car, "Stato": stato,
                "Ultima calibrazione": uc.isoformat() if uc else "",
                "Scadenza (mesi)": str(int(ms)) if ms else "", "Prossima calibrazione": pc, "Certificato": cert.strip()}
    return read


def show_next_ids():
    df = load("Strumenti")
    with st.container(border=True):
        st.markdown("**🔢 Numerazioni disponibili** — usa questi codici per non creare ID doppi")
        for col, (L, (last, nxt, w)) in zip(st.columns(4), next_ids(df).items()):
            col.metric(f"Prossimo {L}", f"{L}-{nxt:0{w}d}")
            col.caption(f"Ultimo usato: {L}-{last:0{w}d}" if last else "Nessun codice presente")


def page_new(me):
    st.subheader("➕ Nuovo strumento / attrezzo")
    box = st.container()  # riempito dopo il salvataggio, così mostra già il numero aggiornato
    with st.form("nuovo", clear_on_submit=True):
        read = tool_fields()
        ok = st.form_submit_button("Aggiungi")
    if ok:
        v = read()
        esistenti = load("Strumenti")["ID"].str.strip().str.upper()
        if not v["ID"] or not v["Descrizione"]:
            st.error("ID e descrizione sono obbligatori.")
        elif (esistenti == v["ID"].upper()).any():
            st.error("Esiste già uno strumento con questo ID.")
        else:
            v.update({"uid": f"n{int(time.time() * 1000)}", "Data assegnazione": today().isoformat(),
                      "Data arrivo": today().isoformat()})
            append("Strumenti", v)
            st.success(f'Aggiunto {v["ID"]}.')
    with box:
        show_next_ids()


# ----------------------------------------------------------------- pagine
def set_view(v):
    st.session_state.dv = v


def kpi_sets(df):
    d = df["_d"]
    return {
        "voci": ("Voci", df),
        "scadute": ("Tarature scadute", df[d < 0].sort_values("_d")),
        "30gg": ("Scadenza 30 gg", df[(d >= 0) & (d <= 30)].sort_values("_d")),
        "assegnati": ("Assegnati", df[mask_assigned(df)]),
        "temp": ("Assegnazioni temporanee", df[mask_temp(df)]),
        "rotti": ("Rotti", df[df["Stato"] == "Rotto"]),
    }


COLS_STD = ["ID", "Descrizione", "Marca", "Modello", "In carico a", "Stato", "Taratura"]
COLS_PERM = ["ID", "Descrizione", "Marca", "Modello", "Seriale", "Stato", "Taratura"]
COLS_TEMP = ["ID", "Descrizione", "Marca", "Modello", "In carico a", "Dal", "Fino al", "Durata", "Residuo", "Stato", "Taratura"]


def tech_sheet_html(name, perm, temp):
    """Pagina HTML autonoma, impostata per la stampa su A4."""
    def tar(r):
        t = fmt(r["Prossima calibrazione"])
        return (f'<span class="sc">{t} (scaduta)</span>' if pd.notna(r["_d"]) and r["_d"] < 0 else t)

    def table(df, heads, getters):
        if df.empty:
            return "<p><i>Nessuno.</i></p>"
        h = "".join(f"<th>{escape(x)}</th>" for x in heads)
        rows = ""
        for _, r in df.iterrows():
            rows += "<tr>" + "".join(f"<td>{g(r)}</td>" for g in getters) + "</tr>"
        return f"<table><thead><tr>{h}</tr></thead><tbody>{rows}</tbody></table>"

    c = lambda k: (lambda r: escape(str(r[k])))
    mm = lambda r: escape(f'{r["Marca"]} {r["Modello"]}'.strip())
    t_perm = table(perm, ["ID", "Descrizione", "Marca / Modello", "Seriale", "Stato", "Prossima taratura"],
                   [c("ID"), c("Descrizione"), mm, c("Seriale"), c("Stato"), tar])
    t_temp = table(temp, ["ID", "Descrizione", "Marca / Modello", "Dal", "Fino al", "Residuo", "Stato", "Prossima taratura"],
                   [c("ID"), c("Descrizione"), mm, c("Dal"), c("Fino al"), c("Residuo"), c("Stato"), tar])
    return f"""<!doctype html><html lang="it"><head><meta charset="utf-8">
<title>Assegnazioni {escape(name)}</title><style>
@page {{ size: A4 portrait; margin: 12mm; }}
body {{ font-family: Arial, Helvetica, sans-serif; font-size: 11px; color: #111; margin: 0; padding: 14px; }}
h1 {{ font-size: 18px; margin: 0 0 2px; }}
h2 {{ font-size: 13px; margin: 16px 0 6px; border-bottom: 1px solid #333; padding-bottom: 2px; }}
.meta {{ color: #555; margin-bottom: 6px; }}
table {{ width: 100%; border-collapse: collapse; margin-bottom: 8px; }}
thead {{ display: table-header-group; }}
th, td {{ border: 1px solid #999; padding: 4px 5px; text-align: left; vertical-align: top; }}
th {{ background: #eee; }}
tr {{ page-break-inside: avoid; }}
.sc {{ color: #b00020; font-weight: bold; }}
.firma {{ margin-top: 36px; display: flex; justify-content: space-between; }}
.firma div {{ width: 45%; border-top: 1px solid #333; padding-top: 3px; }}
button {{ padding: 6px 14px; font-size: 13px; cursor: pointer; margin-bottom: 8px; }}
@media print {{ .noprint {{ display: none; }} body {{ padding: 0; }} }}
</style></head><body>
<button class="noprint" onclick="window.print()">🖨 Stampa (A4)</button>
<h1>Strumenti e attrezzi in carico a {escape(name)}</h1>
<div class="meta">Stampato il {today().strftime("%d/%m/%Y")} · Assegnati: {len(perm)} · Temporanei: {len(temp)}</div>
<h2>Assegnati ({len(perm)})</h2>{t_perm}
<h2>Assegnazioni temporanee ({len(temp)})</h2>{t_temp}
<div class="firma"><div>Data</div><div>Firma {escape(name)}</div></div>
</body></html>"""


def tech_page(df, name, me):
    mine = df[df["In carico a"] == name]
    is_t = mine["Assegnato fino al"].str.strip().ne("")
    perm = mine[~is_t].sort_values("ID")
    temp = add_temp_cols(mine[is_t]).sort_values("Fino al")
    st.subheader(f"👤 {name}")
    m = st.columns(3)
    m[0].metric("In carico in totale", len(mine))
    m[1].metric("Assegnati", len(perm))
    m[2].metric("Temporanei", len(temp))

    html = tech_sheet_html(name, perm, temp)
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", name)
    st.download_button("⬇ Scarica pagina (stampabile A4)", html.encode("utf-8"),
                       f"assegnazioni_{safe}_{today().isoformat()}.html", "text/html")
    with st.expander("🖨 Anteprima e stampa A4"):
        components.html(html, height=650, scrolling=True)

    st.markdown(f"**Assegnati ({len(perm)})**")
    table_with_detail(perm, COLS_PERM, f"tp_{safe}", me)
    st.markdown(f"**Assegnazioni temporanee ({len(temp)})**")
    table_with_detail(temp, COLS_TEMP, f"tt_{safe}", me)


def dash_view(df, view, me):
    st.button("← Torna al riepilogo", on_click=set_view, args=(None,), key="back_dash")
    kind, arg = view
    if kind == "tech":
        tech_page(df, arg, me)
    elif kind == "req":
        r = load("Richieste")
        r = r[r["Stato"] == "nuova"].iloc[::-1]
        st.subheader(f"Nuove richieste ({len(r)})")
        if r.empty:
            st.info("Nessuna nuova richiesta.")
        else:
            st.dataframe(r[["Data", "Tecnico", "Tipo", "Attrezzatura", "Messaggio"]], hide_index=True)
            st.caption("Per rispondere o evadere le richieste apri la sezione 📨 Richieste.")
    else:
        label, sub = kpi_sets(df)[arg]
        st.subheader(f"{label} ({len(sub)})")
        if arg == "temp":
            st.caption("Strumento, a chi è assegnato, da quando e fino a quando. Seleziona una riga per aprire la scheda.")
            table_with_detail(add_temp_cols(sub).sort_values("Fino al"), COLS_TEMP, "dv_temp", me)
        else:
            st.caption("Seleziona una riga per aprire la scheda.")
            table_with_detail(sub, COLS_STD, f"dv_{arg}", me)


def page_dash(me):
    df = decorate(load("Strumenti"))
    view = st.session_state.get("dv")
    if view:
        dash_view(df, view, me)
        return
    sets = kpi_sets(df)
    nuove = int((load("Richieste")["Stato"] == "nuova").sum())
    items = [(k, lab, len(sub), ("kpi", k)) for k, (lab, sub) in sets.items()]
    items.append(("req", "Nuove richieste", nuove, ("req", None)))
    st.caption("Clicca su un numero per vedere il dettaglio.")
    for col, (k, lab, n, v) in zip(st.columns(len(items)), items):
        col.caption(lab)
        col.button(str(n), key=f"kpi_{k}", on_click=set_view, args=(v,), use_container_width=True)

    st.markdown("**👤 Tecnici** — clicca per vedere strumenti e attrezzi assegnati")
    tl = techs()
    if tl:
        cols = st.columns(4)
        for i, t in enumerate(tl):
            sub = df[df["In carico a"] == t]
            nt = int(mask_temp(sub).sum())
            lab = f"{t} · {len(sub)}" + (f" ({nt} temp.)" if nt else "")
            cols[i % 4].button(lab, key=f"tk_{i}", on_click=set_view, args=(("tech", t),), use_container_width=True)

    st.markdown("**Tarature da gestire** (strumenti rotti esclusi)")
    due = df[df["_d"] <= 30].sort_values("_d")
    st.dataframe(due[["ID", "Descrizione", "In carico a", "Taratura"]], hide_index=True)


def page_requests(me):
    r = load("Richieste")
    if me["role"] == "tecnico":
        tools = load("Strumenti")
        tools = tools[tools["Stato"] != "Rotto"]
        opts = {"— nessuna / altro —": ""} | {f"{t.ID} · {t.Descrizione}": t.uid for t in tools.itertuples()}
        with st.form("rq", clear_on_submit=True):
            st.markdown("**Nuova richiesta / nota all’ufficio**")
            tipo = st.selectbox("Tipo", ["Richiesta attrezzatura", "Nota"])
            sel = st.selectbox("Attrezzatura (facoltativa)", list(opts))
            msg = st.text_area("Messaggio")
            if st.form_submit_button("Invia all’ufficio"):
                if not msg.strip() and not opts[sel]:
                    st.error("Scrivi un messaggio o scegli un’attrezzatura.")
                else:
                    append("Richieste", {"id": f"r{int(time.time() * 1000)}", "Data": now().strftime("%Y-%m-%d %H:%M"),
                                         "Tecnico": me["name"], "Tipo": tipo, "uid": opts[sel],
                                         "Attrezzatura": "" if not opts[sel] else sel, "Messaggio": msg.strip(), "Stato": "nuova"})
                    st.rerun()
        r = r[r["Tecnico"] == me["name"]]
    st.markdown(f"**{'Le mie richieste' if me['role'] == 'tecnico' else 'Richieste dei tecnici'} ({len(r)})**")
    for q in r.iloc[::-1].itertuples():
        with st.container(border=True):
            st.markdown(f"**{q.Tecnico}** · {q.Tipo} · `{q.Stato}`")
            st.caption(q.Data)
            if q.Attrezzatura:
                st.write(f"🔧 {q.Attrezzatura}")
            if q.Messaggio:
                st.write(q.Messaggio)
            if me["role"] == "ufficio" and q.Stato != "evasa":
                b = st.columns(3)
                if q.Stato == "nuova" and b[0].button("Segna letta", key=f"l_{q.id}"):
                    update_row("Richieste", q.id, {"Stato": "letta"})
                    st.rerun()
                if q.uid and b[1].button(f"Assegna a {q.Tecnico}", key=f"a_{q.id}"):
                    update_row("Strumenti", q.uid, {"In carico a": q.Tecnico, "Data assegnazione": today().isoformat(), "Assegnato fino al": ""})
                    update_row("Richieste", q.id, {"Stato": "evasa"})
                    st.rerun()
                if b[2].button("Evasa", key=f"e_{q.id}"):
                    update_row("Richieste", q.id, {"Stato": "evasa"})
                    st.rerun()


# --------------------------------------------------- ricerca "AI" sui dati
def ai_answer(q):
    df = decorate(load("Strumenti"))
    l, d, used, extra = q.lower(), [], False, ""
    L = df
    who = next((n for n in [MAG] + techs() if re.search(r"\b" + re.escape(n.lower().split()[0]) + r"\b", l)), None)
    if who:
        used = True
        L = L[L["In carico a"].str.lower().str.contains(who.lower().split()[0], regex=False)]
        d.append(f"in carico a {who}")
    elif "magazzino" in l:
        used = True
        L = L[L["In carico a"].str.lower().str.startswith("mag")]
        d.append("in magazzino")
    tp = "Strumento" if "strument" in l else "Attrezzo" if "attrezz" in l else "Veicolo" if re.search("veicol|furgon", l) else ""
    if tp:
        used = True
        L = L[L["Tipologia"] == tp]
        d.append(f"di tipo {tp.lower()}")
    if "scadut" in l:
        used = True
        L = L[L["_d"] < 0]
        d.append("con taratura scaduta")
    elif re.search("scad|tarat|calibr", l):
        used = True
        n = 30
        m = re.search(r"(\d+)\s*(giorni|gg|g\b|mesi|mese)", l)
        if m:
            n = int(m.group(1)) * (30 if "mes" in m.group(2) else 1)
        elif "trimestr" in l:
            n = 90
        L = L[L["_d"] <= n].sort_values("_d")
        sc = int((L["_d"] < 0).sum())
        d.append(f"con taratura scaduta o in scadenza entro {n} giorni")
        extra = f"Di cui già scaduti: {sc}; in scadenza: {len(L) - sc}.\n\n"
    if re.search("rott|guast", l):
        used = True
        L = L[L["Stato"] == "Rotto"]
        d.append("rotti")
    if "verificare" in l:
        used = True
        L = L[L["Stato"] == "Da verificare"]
        d.append("da verificare")
    if not used:
        stop = {"quanti", "quali", "cerca", "trova", "cosa", "che", "chi", "ha", "il", "lo", "la", "le", "gli", "di", "del",
                "dei", "in", "per", "con", "sono", "mi", "dimmi", "strumenti", "attrezzi"}
        tk = [w for w in re.sub(r"[?.,!]", "", l).split() if len(w) > 1 and w not in stop]
        if not tk:
            return "Posso rispondere su scadenze di taratura, assegnazioni ai tecnici, stati e ricerche per ID, descrizione o marca."
        hay = (L["ID"] + " " + L["Descrizione"] + " " + L["Modello"] + " " + L["Marca"] + " " + L["Seriale"] + " " +
               L["In carico a"] + " " + L["uid"].map(notes_text()).fillna("")).str.lower()
        for w in tk:
            L = L[hay.loc[L.index].str.contains(w, regex=False)]
        d.append("che contengono “" + " ".join(tk) + "”")
    if L.empty:
        return "Nessuna voce trovata " + ", ".join(d) + "."
    righe = []
    for _, r in L.head(15).iterrows():
        extra_r = ", ROTTO" if r["Stato"] == "Rotto" else (f", taratura {fmt(r['Prossima calibrazione'])}" if r["Prossima calibrazione"] else "")
        righe.append(f"- **{r['ID']}** — {r['Descrizione']} {r['Modello']} ({r['In carico a'] or '—'}{extra_r})")
    coda = f"\n- …e altre {len(L) - 15}." if len(L) > 15 else ""
    return f"**{len(L)}** voci {', '.join(d)}.\n\n{extra}" + "\n".join(righe) + coda


def page_ai():
    st.caption("Risponde solo con i dati della tabella. Esempi: “Quanti strumenti sono in scadenza taratura?”, “Cosa ha Yarema?”, “Strumenti di Terry in scadenza”, “Quanti attrezzi sono rotti?”")
    hist = st.session_state.setdefault("ai", [])
    for role, txt in hist:
        st.chat_message(role).markdown(txt)
    q = st.chat_input("Fai una domanda sui dati…")
    if q:
        a = ai_answer(q)
        hist += [("user", q), ("assistant", a)]
        st.chat_message("user").markdown(q)
        st.chat_message("assistant").markdown(a)


# ------------------------------------------------------------------ export
def export_df():
    s = load("Strumenti").copy()
    n = load("Note")
    if n.empty:
        s["Note (storico)"] = ""
    else:
        n = n.assign(t=[(f"[{a} {b}] " if a else "") + c for a, b, c in zip(n["Data"].str[:10], n["Autore"], n["Testo"])])
        s["Note (storico)"] = s["uid"].map(n.groupby("uid")["t"].apply(" | ".join)).fillna("")
    return s.drop(columns=["uid"])


def page_export():
    st.subheader("⬇ Esporta dati aggiornati")
    s = export_df()
    st.download_button("Esporta CSV", s.to_csv(sep=";", index=False).encode("utf-8-sig"), "strumenti_attrezzature.csv", "text/csv")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        s.to_excel(w, sheet_name="Strumenti_e_Attrezzi", index=False)
        load("Richieste").to_excel(w, sheet_name="Richieste", index=False)
        load("Note").to_excel(w, sheet_name="Note", index=False)
    st.download_button("Esporta Excel", buf.getvalue(), "strumenti_attrezzature.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# -------------------------------------------------------------------- main
def main():
    me = st.session_state.get("me")
    if not me:
        login()
        return
    ufficio = me["role"] == "ufficio"
    if ufficio:
        watcher()
        if st.session_state.pop("show_popup", False):
            popup()
    nuove = int((load("Richieste")["Stato"] == "nuova").sum()) if ufficio else 0
    pages = ([("dash", "📊 Riepilogo"), ("inv", "📦 Inventario"), ("new", "➕ Nuovo"),
              ("req", f"📨 Richieste ({nuove})"), ("ai", "🤖 Cerca AI"), ("exp", "⬇ Esporta")] if ufficio else
             [("inv", "📦 Attrezzature"), ("mie", "🧰 Le mie"), ("req", "📨 Richieste e note"), ("ai", "🤖 Cerca AI")])
    with st.sidebar:
        st.markdown(f"**{'Ufficio' if ufficio else 'Tecnico: ' + me['name']}**")
        page = st.radio("Menu", [p[0] for p in pages], format_func=dict(pages).get)
        if ufficio and st.button("🔔 Notifiche"):
            popup()
        if st.button("🔄 Aggiorna dati"):
            load.clear()
            st.rerun()
        if st.button("Esci"):
            st.session_state.clear()
            st.rerun()
    st.title("🔧 Strumenti e Attrezzature")
    s = load("Strumenti")
    if page == "dash":
        page_dash(me)
    elif page == "inv":
        inventory(s, "inv", me)
    elif page == "mie":
        inventory(s[s["In carico a"] == me["name"]], "mie", me)
    elif page == "new":
        page_new(me)
    elif page == "req":
        page_requests(me)
    elif page == "ai":
        page_ai()
    elif page == "exp":
        page_export()


try:
    main()
except gspread.exceptions.APIError as e:
    st.error("Google Sheets ha rifiutato la richiesta (troppe richieste ravvicinate o permessi mancanti). Attendi un minuto e premi “Aggiorna dati”.")
    st.caption(str(e)[:300])
except gspread.exceptions.SpreadsheetNotFound:
    st.error("Foglio Google non trovato: controlla SHEET_ID e che il foglio sia condiviso con l’account di servizio.")
except gspread.exceptions.WorksheetNotFound as e:
    st.error(f"Manca una scheda nel foglio Google: {e}. Servono Strumenti, Tecnici, Richieste e Note.")
except KeyError as e:
    st.error(f"Manca un’impostazione nei Secrets di Streamlit: {e}")
