"""Gestione Strumenti e Attrezzature - Streamlit + Google Sheets."""
import hmac
import io
import re
import time
from datetime import date, datetime
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
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


def page_new(me):
    st.subheader("➕ Nuovo strumento / attrezzo")
    with st.form("nuovo", clear_on_submit=True):
        read = tool_fields()
        ok = st.form_submit_button("Aggiungi")
    if ok:
        v = read()
        if not v["ID"] or not v["Descrizione"]:
            st.error("ID e descrizione sono obbligatori.")
        elif (load("Strumenti")["ID"] == v["ID"]).any():
            st.error("Esiste già uno strumento con questo ID.")
        else:
            v.update({"uid": f"n{int(time.time() * 1000)}", "Data assegnazione": today().isoformat(),
                      "Data arrivo": today().isoformat()})
            append("Strumenti", v)
            st.success(f'Aggiunto {v["ID"]}.')


# ----------------------------------------------------------------- pagine
def page_dash():
    df = decorate(load("Strumenti"))
    k = st.columns(6)
    k[0].metric("Voci", len(df))
    k[1].metric("Tarature scadute", int((df["_d"] < 0).sum()))
    k[2].metric("Scadenza 30 gg", int(((df["_d"] >= 0) & (df["_d"] <= 30)).sum()))
    k[3].metric("Assegnati", int((df["In carico a"].str.strip().ne("") & ~df["In carico a"].str.lower().str.startswith("mag")).sum()))
    k[4].metric("Rotti", int((df["Stato"] == "Rotto").sum()))
    r = load("Richieste")
    k[5].metric("Nuove richieste", int((r["Stato"] == "nuova").sum()))
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
        page_dash()
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
