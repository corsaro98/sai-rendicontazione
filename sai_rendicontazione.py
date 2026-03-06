"""
SAI Rendicontazione — App Streamlit (senza pdfplumber, compatibile Python 3.14)
"""
import io, re, zipfile
from datetime import datetime
import pandas as pd
import pypdf
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

st.set_page_config(page_title="SAI · Rendicontazione", page_icon="📋", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background: #F5F4F0; }
.block-container { padding: 2rem 3rem; max-width: 1400px; }
.hdr { background: #1C1C1C; border-radius: 14px; padding: 1.75rem 2.25rem; margin-bottom: 2rem; }
.hdr h1 { color: #FFF; font-size: 1.5rem; margin: 0 0 .25rem; }
.hdr p  { color: #888; font-size: .82rem; margin: 0; font-family: 'DM Mono', monospace; }
.badge { background: #C5F135; color: #1C1C1C; font-size: .62rem; font-family: 'DM Mono', monospace;
         padding: .2rem .5rem; border-radius: 4px; letter-spacing: .06em; vertical-align: middle; margin-left: .4rem; }
.mrow { display: flex; gap: .9rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.mbox { flex: 1; min-width: 110px; background: white; border: 1px solid #E5E5E1; border-radius: 10px; padding: 1rem 1.25rem; }
.mlbl { font-size: .65rem; font-family: 'DM Mono', monospace; text-transform: uppercase; letter-spacing: .07em; color: #999; margin-bottom: .2rem; }
.mval { font-size: 1.8rem; font-weight: 600; color: #1C1C1C; line-height: 1; }
.th { display: grid; grid-template-columns: 40px 90px 1fr 120px 100px 90px; gap: .4rem; padding: .55rem .9rem;
      background: #EFEFEB; border-radius: 8px 8px 0 0; border: 1px solid #E5E5E1;
      font-family: 'DM Mono', monospace; font-size: .62rem; text-transform: uppercase; letter-spacing: .07em; color: #999; }
.tb { border: 1px solid #E5E5E1; border-top: none; border-radius: 0 0 8px 8px; background: white; overflow: hidden; margin-bottom: 1.5rem; }
.tr { display: grid; grid-template-columns: 40px 90px 1fr 120px 100px 90px; gap: .4rem; padding: .6rem .9rem;
      border-bottom: 1px solid #F3F3F0; align-items: start; font-size: .8rem; }
.tr:last-child { border-bottom: none; }
.tr:hover { background: #FAFAF8; }
.fn { font-family: 'DM Mono', monospace; font-size: .68rem; background: #F5F4F0; border: 1px solid #E0E0DA;
      border-radius: 5px; padding: .22rem .45rem; color: #444; word-break: break-all; line-height: 1.4; }
.mono { font-family: 'DM Mono', monospace; font-size: .75rem; color: #555; }
.av { font-size: .68rem; color: #B85C00; margin-top: .2rem; }
.bdg { display: inline-block; padding: .18rem .55rem; border-radius: 20px; font-size: .65rem; font-family: 'DM Mono', monospace; font-weight: 500; }
.ba { background: #E6F9D8; color: #2E7D0C; }
.bm { background: #FFF5D0; color: #956800; }
.bb { background: #FFE8CC; color: #B85C00; }
.be { background: #FDECEA; color: #B71C1C; }
.bi { background: #EFEFEB; color: #888; }
div.stButton > button { background: #1C1C1C !important; color: white !important; border: none !important;
    border-radius: 8px !important; padding: .65rem 1.5rem !important; font-weight: 500 !important; width: 100%; }
.stDownloadButton button { background: #C5F135 !important; color: #1C1C1C !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important; width: 100%; }
.stProgress > div > div { background: #C5F135 !important; }
hr { border: none; border-top: 1px solid #E5E5E1; margin: 1.5rem 0; }
.slbl { font-size: .65rem; font-family: 'DM Mono', monospace; text-transform: uppercase; letter-spacing: .08em; color: #999; margin-bottom: .35rem; }
</style>
""", unsafe_allow_html=True)

# ── FUNZIONI CORE ─────────────────────────────────────────────────────────────

def sanitize(s):
    s = str(s).strip()
    s = s.replace("/", "-")
    s = re.sub(r'[\\*?:"<>|]', "", s)
    return re.sub(r"\s+", " ", s)[:80]

def parse_importo(s):
    try: return float(str(s).replace(".", "").replace(",", ".").strip())
    except: return None

def token_sim(a, b):
    a_t = set(re.split(r"[\s\-_/]+", a.upper())) - {""}
    b_t = set(re.split(r"[\s\-_/]+", b.upper())) - {""}
    if not a_t or not b_t: return 0
    return int(100 * len(a_t & b_t) / max(len(a_t), len(b_t)))

def _numeri_fattura(causale):
    m = re.search(r"nr\.?\s+([\w]+(?:/[\w]+){4,})", causale, re.I)
    if m: return [p for p in m.group(1).split("/") if p.strip()]
    m = re.search(r"nr\.?\s+(FPR\s+[\w/]+)", causale, re.I)
    if m: return [m.group(1).strip()]
    raw = re.findall(r"nr\.?\s+([\w/.\-]+?)(?=\s+del\b|\s+nr\.?\b|\s+e\s+ricevute|\s*$|\s*-\s)", causale, re.I)
    if not raw: raw = re.findall(r"nr\.?\s+([\w/.\-]+)", causale, re.I)
    return [n.strip().rstrip(".-") for n in raw
            if n.strip().rstrip(".-") and not re.match(r"\d{2}/\d{2}/\d{4}", n)
            and n.lower() not in ("del", "e", "e.", "al")]

def estrai_bonifici(pdf_bytes):
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    out = []
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""

        # Beneficiario: il nome precede "Beneficiario:" nella stessa riga
        m = re.search(r"([A-Z][^\n]+?)Beneficiario:", text)
        ben = m.group(1).strip() if m else ""
        ben = re.split(r"\s*-\s*(?:LEI|Persona)", ben)[0].strip()

        # Importo: primo numero EUR nella pagina
        m = re.search(r"([\d\.,]+)\s*EUR", text)
        imp_str = m.group(1) if m else ""

        # Data addebito: cerca data prima di EUR (formato DD.MM.YYYY)
        m = re.search(r"(\d{2}\.\d{2}\.\d{4})\s+[\d\.,]+\s*EUR", text)
        data_add = m.group(1) if m else ""
        if not data_add:
            m = re.search(r"Data di addebito:\s*\n?\s*(\d{2}\.\d{2}\.\d{4})", text)
            data_add = m.group(1) if m else ""
        m2 = re.search(r"Data creazione:\s*(\d{2}\.\d{2}\.\d{4})", text)
        data_pag = data_add or (m2.group(1) if m2 else "")

        # Causale: riga dopo "140 caratteri)"
        m = re.search(r"140 caratteri\)\s*\n(.+?)(?:\n|$)", text)
        causale = m.group(1).strip() if m else ""

        numeri = _numeri_fattura(causale)
        m = re.search(r"\bdel\s+(\d{2}/\d{2}/\d{4})", causale, re.I)
        data_fatt = m.group(1) if m else ""

        out.append(dict(
            pagina=i+1, beneficiario=ben, importo=parse_importo(imp_str),
            importo_str=imp_str, data_pagamento=data_pag,
            causale=causale, numeri_fattura=numeri, data_fattura=data_fatt,
        ))
    return out

def carica_registro(excel_bytes):
    df = pd.read_excel(io.BytesIO(excel_bytes), header=6, skiprows=[7])
    df.columns = ["_","N","Natura","Data_Doc","N_Documento","Modalita_Pagamento",
                  "Data_Pagamento","Cod_Spesa","Descrizione","Importo_Totale",
                  "Finanziamento","Importo_Imputato","Coop"]
    df = df[df["N"].notna() & (df["N"] != "N.")].copy()
    df["N"] = df["N"].astype(str).str.strip()
    df["N_Documento"] = df["N_Documento"].astype(str).str.strip()
    df["Desc_norm"] = df["Descrizione"].astype(str).str.upper().str.strip()
    df["Importo_Totale"] = pd.to_numeric(df["Importo_Totale"], errors="coerce")
    return df

def _ndoc_match(ndoc, nf, nf_norm):
    nd = re.sub(r"^FPR\s*", "", ndoc, flags=re.I).strip()
    return ndoc == nf or ndoc == nf_norm or nd == nf_norm

def abbina(bon, reg):
    mb = reg["Modalita_Pagamento"].astype(str).str.upper().str.contains("BONI", na=False)
    mv = reg["Modalita_Pagamento"].isna() | reg["Modalita_Pagamento"].astype(str).str.strip().isin(["","nan","None"])
    sub = reg[mb | mv].copy()
    if bon["numeri_fattura"]:
        trovati = []
        for nf in bon["numeri_fattura"]:
            nf = nf.strip()
            nfn = re.sub(r"^FPR\s*", "", nf, flags=re.I).strip()
            hit = sub[sub["N_Documento"].apply(lambda x: _ndoc_match(str(x).strip(), nf, nfn))]
            trovati.extend(hit["N"].tolist())
        seen = set(); trovati = [n for n in trovati if not (n in seen or seen.add(n))]
        if trovati:
            diff = abs(reg[reg["N"].isin(trovati)]["Importo_Totale"].sum() - (bon["importo"] or 0))
            return trovati, "fattura", "alta" if diff < 1.0 else "media"
    if bon["importo"]:
        fi = sub[abs(sub["Importo_Totale"] - bon["importo"]) < 0.02]
        if len(fi) == 1: return fi["N"].tolist(), "importo", "media"
    bc = re.sub(r"\b(NUOVO|NEW|CONTO|SRL|SRLS|SPA|SNC|SAS)\b","", bon["beneficiario"].upper()).strip()
    bs, bn = 0, None
    for _, row in sub.iterrows():
        s = token_sim(bc, str(row["Desc_norm"]))
        if s > bs: bs, bn = s, row["N"]
    if bs >= 75 and bn:
        r = reg[reg["N"] == bn]
        ok = not (not r.empty and bon["importo"] and abs(r.iloc[0]["Importo_Totale"] - bon["importo"]) > 5)
        return [bn], "nome_fuzzy", ("alta" if bs >= 90 else "media") if ok else "bassa"
    return [], "nessuno", "nessuna"

_IP = [r"ricarica\s+(cassa|fondo)", r"giroconto", r"pocket\s+money",
       r"\bvitto\s+(gen|feb|mar|apr|mag|giu|lug|ago|set|ott|nov|dic)"]
_IB = ["carta prepagata", "i girasoli scs"]

def is_interno(b):
    c = b["causale"].lower(); bn = b["beneficiario"].lower()
    if any(p in bn for p in _IB): return True
    for p in _IP:
        if re.search(p, c): return True
    return False

def build_nome(n_reg, ben, numeri, data_fatt):
    parti = [sanitize(n_reg) if n_reg else "???", "BONIFICO", sanitize(ben),
             f"FATT. N. {sanitize(' + '.join(numeri))}" if numeri else "FATT. N. -"]
    if data_fatt: parti.append(data_fatt.replace("/","-").replace(".","-"))
    return " - ".join(parti) + ".pdf"

def estrai_pagina(pdf_bytes, idx):
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        writer = pypdf.PdfWriter()
        writer.add_page(reader.pages[idx])
        buf = io.BytesIO(); writer.write(buf); return buf.getvalue()
    except: return None

def compila_registro(excel_bytes, risultati, reg_df):
    wb = load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    COL_N=2; COL_NDOC=5; COL_MOD=6; COL_DATA=7
    n_to_row = {}
    for row in ws.iter_rows(min_row=8, max_row=ws.max_row):
        c = row[COL_N-1]; v = str(c.value or "").strip()
        if v and v not in ("N.","nan"): n_to_row[v] = c.row
    verde = PatternFill("solid", fgColor="D4F5C8")
    vfont = Font(color="1A6B0A")
    compilate = 0
    for ris in risultati:
        if ris["confidence"] not in ("alta","media","bassa"): continue
        if not ris["n_registro"]: continue
        dt = None
        if ris["data_pagamento"]:
            try: dt = datetime.strptime(ris["data_pagamento"], "%d.%m.%Y")
            except: pass
        for n_reg in ris["n_registro"]:
            ex_row = n_to_row.get(str(n_reg))
            if ex_row is None: continue
            mod = False
            c_mod = ws.cell(row=ex_row, column=COL_MOD)
            if str(c_mod.value or "").strip() in ("","nan","None"):
                c_mod.value="BONIFICO"; c_mod.fill=verde; c_mod.font=vfont; mod=True
            if dt:
                c_data = ws.cell(row=ex_row, column=COL_DATA)
                if str(c_data.value or "").strip() in ("","nan","None","NaT"):
                    c_data.value=dt; c_data.number_format="DD/MM/YYYY"
                    c_data.fill=verde; c_data.font=vfont; mod=True
            if mod: compilate += 1
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue(), compilate

# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hdr">
  <h1>SAI · Rendicontazione <span class="badge">BETA</span></h1>
  <p>Abbina i bonifici bancari al registro · Rinomina i PDF · Compila data pagamento nel registro</p>
</div>""", unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    st.markdown('<div class="slbl">📂 Registro spese (.xlsx)</div>', unsafe_allow_html=True)
    reg_file = st.file_uploader("Registro", type=["xlsx"], key="reg", label_visibility="collapsed")
    if reg_file: st.success(f"✓ {reg_file.name}")
with c2:
    st.markdown('<div class="slbl">🏦 PDF bonifici (Intesa Sanpaolo)</div>', unsafe_allow_html=True)
    pdf_file = st.file_uploader("PDF", type=["pdf"], key="pdf", label_visibility="collapsed")
    if pdf_file: st.success(f"✓ {pdf_file.name}  ({pdf_file.size//1024} KB)")

st.markdown("<hr>", unsafe_allow_html=True)

if reg_file and pdf_file:
    if st.button("▶  Avvia elaborazione", use_container_width=True):
        prog = st.progress(0, text="Lettura PDF…")
        pdf_bytes = pdf_file.read(); excel_bytes = reg_file.read()
        bonifici = estrai_bonifici(pdf_bytes)
        prog.progress(20, text="Lettura registro…")
        reg_df = carica_registro(excel_bytes)
        prog.progress(35, text="Abbinamento…")
        risultati = []
        for idx, bon in enumerate(bonifici):
            prog.progress(35+int(45*idx/max(len(bonifici),1)), text=f"Abbinamento {idx+1}/{len(bonifici)}…")
            if is_interno(bon):
                ris = {**bon, "n_registro":[], "metodo":"interno","confidence":"interno",
                       "nome_file":f"INTERNO - BONIFICO - {sanitize(bon['beneficiario'])}.pdf","avvisi":[]}
            else:
                nl, met, conf = abbina(bon, reg_df)
                avvisi = []
                if len(nl)>1 and met=="fattura":
                    s = reg_df[reg_df["N"].isin(nl)]["Importo_Totale"].sum()
                    if abs(s-(bon["importo"] or 0))>1.0: avvisi.append(f"Somma {s:.2f}€ ≠ {bon['importo']:.2f}€")
                ris = {**bon, "n_registro":nl, "metodo":met, "confidence":conf,
                       "nome_file":build_nome("+".join(nl),bon["beneficiario"],bon["numeri_fattura"],bon["data_fattura"]),
                       "avvisi":avvisi}
            risultati.append(ris)
        prog.progress(82, text="Compilazione registro…")
        reg_out, n_comp = compila_registro(excel_bytes, risultati, reg_df)
        prog.progress(93, text="Creazione ZIP…")
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for ris in risultati:
                pb2 = estrai_pagina(pdf_bytes, ris["pagina"]-1)
                if pb2:
                    folder = "interni" if ris["confidence"]=="interno" else \
                             "da_verificare" if ris["confidence"] in ("bassa","nessuna") else "abbinati"
                    zf.writestr(f"{folder}/{ris['nome_file']}", pb2)
        prog.progress(100, text="✓ Completato!")
        prog.empty()
        st.session_state.update(risultati=risultati, reg_out=reg_out,
                                zip_data=zip_buf.getvalue(), n_comp=n_comp, done=True)

if st.session_state.get("done"):
    R = st.session_state["risultati"]; n_comp = st.session_state["n_comp"]
    tot=len(R); alta=sum(1 for r in R if r["confidence"]=="alta")
    media=sum(1 for r in R if r["confidence"]=="media")
    interni=sum(1 for r in R if r["confidence"]=="interno")
    da_ver=sum(1 for r in R if r["confidence"] in ("bassa","nessuna"))
    st.markdown(f"""
<div class="mrow">
  <div class="mbox"><div class="mlbl">Bonifici PDF</div><div class="mval">{tot}</div></div>
  <div class="mbox"><div class="mlbl">Alta</div><div class="mval" style="color:#2E7D0C">{alta}</div></div>
  <div class="mbox"><div class="mlbl">Media</div><div class="mval" style="color:#956800">{media}</div></div>
  <div class="mbox"><div class="mlbl">Da verificare</div><div class="mval" style="color:#B71C1C">{da_ver}</div></div>
  <div class="mbox"><div class="mlbl">Interni</div><div class="mval" style="color:#888">{interni}</div></div>
  <div class="mbox"><div class="mlbl">Righe compilate</div><div class="mval" style="color:#2E7D0C">{n_comp}</div></div>
</div>""", unsafe_allow_html=True)

    st.markdown('<div class="slbl">📋 Dettaglio abbinamenti</div>', unsafe_allow_html=True)
    st.markdown("""<div class="th"><span>Pag.</span><span>Importo</span><span>Nome file PDF</span>
      <span>N° Registro</span><span>Metodo</span><span>Stato</span></div><div class="tb">""", unsafe_allow_html=True)

    _BADGE = {"alta":'<span class="bdg ba">✓ Alta</span>', "media":'<span class="bdg bm">⚠ Media</span>',
              "bassa":'<span class="bdg bb">⚠ Bassa</span>', "nessuna":'<span class="bdg be">✕ Non trovato</span>',
              "interno":'<span class="bdg bi">↩ Interno</span>'}
    _MET = {"fattura":"🔢 N° fatt.","importo":"💶 Importo","nome_fuzzy":"👤 Fuzzy","interno":"↩ —","nessuno":"— —"}

    for r in R:
        badge=_BADGE.get(r["confidence"],""); metodo=_MET.get(r["metodo"],r["metodo"])
        n_reg=", ".join(r["n_registro"]) if r["n_registro"] else "—"
        imp=f"€&nbsp;{r['importo_str']}" if r["importo_str"] else "—"
        av_html="".join(f'<div class="av">⚡ {a}</div>' for a in r.get("avvisi",[]))
        st.markdown(f"""<div class="tr">
          <span class="mono" style="color:#AAA">{r['pagina']:02d}</span>
          <span class="mono">{imp}</span>
          <span><div class="fn">{r['nome_file']}</div>{av_html}</span>
          <span class="mono">{n_reg}</span>
          <span style="font-size:.75rem;color:#777">{metodo}</span>
          {badge}</div>""", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)

    d1, d2 = st.columns(2)
    with d1:
        st.markdown('<div class="slbl">⬇ PDF rinominati (.zip)</div>', unsafe_allow_html=True)
        st.download_button("⬇  Scarica PDF rinominati (.zip)", data=st.session_state["zip_data"],
            file_name="SAI_PDF_rinominati.zip", mime="application/zip", use_container_width=True)
    with d2:
        st.markdown('<div class="slbl">⬇ Registro Excel aggiornato</div>', unsafe_allow_html=True)
        st.download_button("⬇  Scarica registro aggiornato (.xlsx)", data=st.session_state["reg_out"],
            file_name="Registro_aggiornato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
else:
    if not (reg_file and pdf_file):
        st.markdown("""<div style="text-align:center;padding:4rem 2rem;color:#C0C0BC">
          <div style="font-size:3.5rem;margin-bottom:1rem">📂</div>
          <div style="font-family:'DM Mono',monospace;font-size:.88rem">
            Carica il registro Excel e il PDF dei bonifici per iniziare</div></div>""", unsafe_allow_html=True)
