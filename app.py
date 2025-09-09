import streamlit as st
import pandas as pd
import numpy as np
import openpyxl as op
import io
import datetime as dt

st.set_page_config(page_title="Analisi Sottoscorta DPC", layout="wide")
st.title("📊 Analisi Sottoscorta - DPC")

st.write("Carica tutti i file CSV richiesti:")

# --- File uploader ---
file_contratti = st.file_uploader("Contratti", type=[".csv"])
file_ordini = st.file_uploader("Ordini", type=[".csv"])
file_anag = st.file_uploader("Anagrafica", type=[".csv"])
file_carichi = st.file_uploader("Carichi", type=[".csv"])
file_sottoscorta = st.file_uploader("Sottoscorte/Fabbisogni", type=[".csv"])
file_carenze = st.file_uploader("Carenze", type=[".csv"])

numero_gg_riordino=[i for i in range(50,76)]

if st.button("Esegui analisi"):
    if not all([file_contratti, file_ordini, file_anag, file_carichi, file_sottoscorta, file_carenze, 
                    st.selectbox("Selezionare autonomia sotto la quale riordinare",numero_gg_riordino)]):
        st.error("⚠️ Devi caricare tutti i file CSV!")
    else:
        # --- Creazione DataFrame ---
        df_c = pd.read_csv(file_contratti, sep=";", encoding="latin")
        df_o = pd.read_csv(file_ordini, sep=";", encoding="latin")
        df_anag = pd.read_csv(file_anag, sep=";", encoding="latin", dtype={"MinSan10": str})
        df_carichi = pd.read_csv(file_carichi, sep=";", encoding="latin")
        df_sott = pd.read_csv(file_sottoscorta, sep=";", encoding="latin", dtype={"Minsan": str})
        df_carenze = pd.read_csv(file_carenze, sep=";", encoding="latin", dtype={"Minsan": str})

        riordino= st.selectbox("Selezionare autonomia sotto la quale riordinare",numero_gg_riordino)
        # --- Prodotti da escludere ---
        prodEscl = [
            "042494070","042494029","044924025","043208091","043208038","045183050","045183148","043443136","043443047",
            "044229312","044229223","044229045","044229134","043145022","043145061","043375118","043375082","043375056",
            "043375029","046343113","046343253","046339089","046339026","046342085","046342022"
        ]

        # --- Formati DataFrame ---
        # anagrafica
        df_anag = df_anag.iloc[:, [0,1,2,10]].rename(columns={"MinSan10": "Minsan"})

        # contratti
        df_c = df_c.iloc[:, [0,1,2,3,4,5,6,8,9,10,11,22,23,24,26,27,37,38]]
        for data in ["Validità dal", "Validità al"]:
            df_c[data] = pd.to_datetime(df_c[data])
        for val in ["Importo", "Prezzo"]:
            df_c[val] = df_c[val].apply(lambda x: float(str(x).replace(",", ".")))
        for val in ["Qta", "Qta Ordinato"]:
            try:
                df_c[val] = df_c[val].astype(int)
            except:
                df_c[val] = df_c[val].apply(lambda x: int(str(x).split(",")[0]))

        Colonne = {
            "Tipo contratto":"TipoContratto",
            "Validità dal":"DataIn",
            "Validità al":"DataFin",
            "Stato":"StatoContratto",
            "Codice CIG":"CIG",
            "Descrizione CIG":"DescrizioneCIG",
            "Stato Riga":"StatoRiga",
            "Prodotto":"CodProd",
            "Descr.":"Prodotto",
            "Qta Ordinato":"QtaOrdinato"
        }

        df_c.rename(columns=Colonne, inplace=True)
        df_c.insert(13, "Minsan", pd.merge(df_c["CodProd"], df_anag[["Codice","Minsan"]],
                                           left_on="CodProd", right_on="Codice", how="left")["Minsan"])
        df_c.insert(18, "QtaResidua", df_c["Qta"] - df_c["QtaOrdinato"])

        # ordini
        df_o.insert(0, "Ordine", "DPC-" + df_o["Anno"].astype(str) + "-" + df_o["Num."].astype(str))
        df_o["Data ordine"] = pd.to_datetime(df_o["Data ordine"], format="%d/%m/%Y")
        df_o = df_o.iloc[:, [0,6,8,9,14,15,16,19,22]]
        df_o.rename(columns={"Data ordine":"Data","Qta/Val Rettificata":"Qta","Prezzo Unit.":"Prezzo"}, inplace=True)
        df_o.insert(6, "Minsan", pd.merge(df_o["Prodotto"], df_anag[["Codice","Minsan"]],
                                         left_on="Prodotto", right_on="Codice", how="left")["Minsan"])

        # carichi
        df_carichi["Data Attività"] = pd.to_datetime(df_carichi["Data Attività"].str.slice(0,10), format="%d/%m/%Y")
        df_carichi["Minsan"] = df_carichi["Prodotto"].str.slice(0,9)
        df_carichi["Prodotto"] = df_carichi["Prodotto"].str.slice(12)
        df_carichi = df_carichi[(~df_carichi["Riferimento Ordine Carico"].str.contains("DPC24")) &
                                (df_carichi["Riferimento Ordine Carico"].str.len() == 9)]
        df_carichi = df_carichi.iloc[:, [0,1,5,2,3,4]].rename(columns={
            "Data Attività":"Data",
            "Riferimento Ordine Carico":"Ordine",
            "Qta Movimentata":"Qta"
        })
        df_carichi["Ordine"] = "DPC-2025-" + df_carichi["Ordine"].str.slice(5).astype(int).astype(str)

        # sottoscorta
        df_sott = df_sott.fillna("")
        indici = []
        for i in range(2):
            lista = df_sott.iloc[i].tolist()
            indici.append(lista)
        new_indice = []
        for i in range(len(lista)):
            val1 = str(indici[0][i])
            val2 = str(indici[1][i])
            if val1 == "":
                new_indice.append(val2)
            elif val2 == "":
                new_indice.append(val1)
            else:
                new_indice.append(val1 + "-" + val2)
        df_sott.columns = new_indice
        df_sott.drop(axis=0, index=[0,1], inplace=True)
        df_sott.replace("", np.nan, inplace=True)
        df_sott.dropna(axis=1, how='all', inplace=True)
        df_sott = df_sott.iloc[:, [0,1,2,4,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,22,23,28,29]]
        df_sott.rename(columns={"Domanda Media Giornaliera":"Cmg","Giacenza Totale":"Giacenza"}, inplace=True)
        for val in ["Cmg","Giacenza"]:
            df_sott[val] = df_sott[val].apply(lambda x: float(str(x).replace(",", ".")))

        # --- Sezione elaborazione df ---
        # ORDINI
        df_o = df_o[(df_o["Stato"]!="Revocato") & (df_o["Qta"] != 0)]
        carichiOrd = df_carichi.groupby(["Minsan","Ordine"])["Qta"].sum().reset_index()
        df_o["QtaCaricata"] = pd.merge(df_o, carichiOrd, on=["Minsan","Ordine"], how="left", suffixes=["","Caricata"]).fillna(0)["QtaCaricata"].values
        df_o["DaCaricare"] = df_o["Qta"] - df_o["QtaCaricata"]
        OrdiniPendenti = df_o[df_o["DaCaricare"] > 0]
        DaCaricare = OrdiniPendenti.groupby("Minsan")["DaCaricare"].sum().reset_index()
        UltimoOrdPendente = OrdiniPendenti[["Minsan","Ordine","Data","Qta"]].sort_values(by=["Minsan","Data"], ascending=[True,False]).drop_duplicates(subset="Minsan")
        ProdDanno = df_o[df_o["Autorizzazione"]=="DPC/2025/1-2"][["Minsan","Fornitore"]].drop_duplicates()

        # CONTRATTI
        ContAperti = df_c[(df_c["StatoRiga"]=="Aperto") & (df_c["StatoContratto"]=="Aperto")]
        DispProd = ContAperti.groupby("Minsan")["QtaResidua"].sum().reset_index()
        ContValidi = df_c[df_c["DataFin"] >= pd.Timestamp.today().normalize()]
        ContrattiOrdinati = df_c.sort_values(by=["Minsan","StatoContratto","Anno","Numero"], ascending=[True,True,False,False])
        UltimoContratto = ContrattiOrdinati.drop_duplicates(subset="Minsan", keep="first")
        UltimoContratto["TipoGara"] = UltimoContratto["Descrizione"].apply(lambda x: "MEPA" if "mepa" in str(x).lower() else "")
        Mepa = UltimoContratto[UltimoContratto["TipoGara"]=="MEPA"]["Minsan"].unique()
        ProdGara = UltimoContratto[(UltimoContratto["TipoContratto"]=='CONTRATTO DPC IN GARA') &
                                   (UltimoContratto["Minsan"].isin(ContValidi["Minsan"]))]["Minsan"]
        ProdEconomia = UltimoContratto[(UltimoContratto["TipoContratto"]=='CONTRATTI DPC IN ECONOMIA') &
                                       (UltimoContratto["Minsan"].isin(ContValidi["Minsan"]))]["Minsan"]

        # SOTTOSCORTA
        GruppoEq = df_sott.groupby("GruppoEq")[["Cmg","Giacenza"]].sum().reset_index()
        df_sott = pd.merge(df_sott, GruppoEq, on="GruppoEq", suffixes=["Prod","GruppoEq"], how="left")
        df_sott = pd.merge(df_sott, DispProd, on="Minsan", how="left")
        df_sott["QtaResidua"] = df_sott["QtaResidua"].fillna(0)
        df_sott = pd.merge(df_sott, UltimoContratto[["Minsan", "Fornitore"]], on="Minsan", how="left")
        df_sott = pd.merge(df_sott, DaCaricare, on="Minsan", how="left")
        df_sott["DaCaricare"] = df_sott["DaCaricare"].fillna(0)
        df_sott = pd.merge(df_sott, UltimoOrdPendente[["Minsan","Ordine","Data"]], on="Minsan", how="left")
        df_sott = pd.merge(df_sott, df_carenze, on="Minsan", how="left")
        df_sott.loc[df_sott["GruppoEq"].isnull(), "CmgGruppoEq"] = df_sott["CmgProd"]
        df_sott.loc[df_sott["GruppoEq"].isnull(), "GiacenzaGruppoEq"] = df_sott["GiacenzaProd"]

        # Categoria prodotto
        CategorieProdotto = [
            df_sott["Minsan"].isin(prodEscl),
            (df_sott["Minsan"].isin(ProdGara)) & (df_sott["Minsan"].isin(ContValidi["Minsan"])),
            (df_sott["Minsan"].isin(ProdEconomia)) & (df_sott["Minsan"].isin(ContValidi["Minsan"])),
            (df_sott["Minsan"].isin(ProdDanno["Minsan"])),
            (~df_sott["Minsan"].isin(ContValidi["Minsan"])) & (df_sott["Minsan"].isin(df_c["Minsan"])),
            (~df_sott["Minsan"].isin(df_c["Minsan"]))
        ]
        Cat1 = ["IN ESAURIMENTO","GARA","ECONOMIA","DANNO","FUORI GARA","MAI CONTRATTUALIZZATO"]
        df_sott["TipoProd"] = np.select(CategorieProdotto, Cat1, default="Altro")
        df_sott["Mepa"] = df_sott["Minsan"].apply(lambda x: "si" if x in Mepa else "no")

        # Funzione Accordo Quadro
        def AccordoQuadro(eq):
            df_eq = df_sott[(df_sott["GruppoEq"]==eq) & (df_sott["TipoProd"]=="GARA")]
            if len(df_eq) > 1:
                minsan = df_eq["Minsan"].unique()
                scadenzaContratti = ContValidi[(ContValidi["TipoContratto"]=='CONTRATTO DPC IN GARA') &
                                               (ContValidi["Minsan"].isin(minsan))]["DataFin"].unique()
                return "AQ" if len(scadenzaContratti) == 1 else ""
            return ""
        
        df_sott.loc[df_sott["TipoProd"]=="GARA", "AccordoQuadro"] = df_sott[df_sott["TipoProd"]=="GARA"]["GruppoEq"].apply(AccordoQuadro)
        df_sott["TipoProd"] = df_sott.apply(lambda x: "AQ" if x["AccordoQuadro"]=="AQ" else x["TipoProd"], axis=1)
        df_sott.drop("AccordoQuadro", axis=1, inplace=True)
        df_sott.loc[df_sott["TipoProd"]=="AQ", "CmgGruppoEq"] = df_sott["CmgProd"]
        df_sott.loc[df_sott["TipoProd"]=="AQ", "GiacenzaGruppoEq"] = df_sott["GiacenzaProd"]

        df_sott["Autonomia"] = (df_sott["GiacenzaGruppoEq"] + df_sott["DaCaricare"]) / df_sott["CmgGruppoEq"]
        df_sott["Autonomia"] = df_sott["Autonomia"].replace([np.inf, np.nan], 9999).astype(int)

        df_sott.loc[df_sott["TipoProd"]=="DANNO", "Fornitore"] = pd.merge(
            df_sott.loc[df_sott["TipoProd"]=="DANNO"],
            ProdDanno,
            on="Minsan",
            how="left"
        )["Fornitore_y"].values

        def OrdineForn(forn):
            return len(df_sott[(df_sott["Fornitore"]==forn) & (df_sott["Autonomia"] <= 45)])

        df_sott.loc[df_sott["TipoProd"].isin(["GARA","DANNO","ECONOMIA","AQ"]), "QtaDaOrdinare"] = (
            df_sott.loc[df_sott["TipoProd"].isin(["GARA","DANNO","ECONOMIA","AQ"])].apply(
                lambda x: int((75*x["CmgGruppoEq"]-(x["GiacenzaGruppoEq"]+x["DaCaricare"]))) 
                if OrdineForn(x["Fornitore"]) >= 1 and x["Autonomia"] < riordino else 0,
                axis=1
            )
        )
        df_sott["QtaDaOrdinare"] = df_sott["QtaDaOrdinare"].fillna(0)

        # Stato riga
        CategorieProdotto2 = [
            df_sott["Minsan"].isin(ContAperti["Minsan"]),
            ~df_sott["Minsan"].isin(ContAperti["Minsan"])
        ]
        Cat2 = ["RIGA APERTA", "RIGA CHIUSA"]
        df_sott["StatoRiga"] = np.select(CategorieProdotto2, Cat2, default="Altro")

        df_sott.info()

        # Ridimensiono il df sottoscorta
        df_sott = df_sott.iloc[:, [5,6,7,8,9,10,11,12,13,14,15,16,17,18,0,1,26,21,22,27,28,29,30,19,23,20,24,33,31,32,35,25,34]]
        df_sott = df_sott.sort_values(by=["Fornitore","Descrizione","Conservazione","TipoProd"])

        df_sott_ord = df_sott[(df_sott["QtaDaOrdinare"]>0) & 
                              (df_sott["TipoProd"].isin(["GARA","AQ","ECONOMIA","DANNO"]))].sort_values(by=["Fornitore","Descrizione","Conservazione"])

        # --- Salvataggio Excel in memoria ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as ex:
            df_sott.to_excel(ex, sheet_name="sottoscorta webdpc", index=False)
            df_sott_ord.to_excel(ex, sheet_name="proposta d'ordine", index=False)

        # --- Formattazione con openpyxl ---
        output.seek(0)
        wb = op.load_workbook(output)
        for foglio in wb:
            n_col = foglio.max_column
            foglio.cell(1, n_col+1, value="Nuova Autonomia")
            contaR=1
    
            for riga in foglio:
                contaC=1
                for cella in riga:
                    cella.font=op.styles.Font(name="Calibri", size=12)
                    cella.border=op.styles.borders.Border(
                        top=op.styles.borders.Side(style='thin'),
                        left=op.styles.borders.Side(style='thin'),
                        right=op.styles.borders.Side(style='thin'),
                        bottom=op.styles.borders.Side(style='thin'))
                    cella.alignment=op.styles.Alignment(vertical="center")
                    if contaR==1:
                        cella.font=op.styles.Font(name="Calibri", size=12, bold=True)
                        cella.alignment=op.styles.Alignment(vertical="center",horizontal="center")
                    if contaR!=1 and contaC==n_col+1:
                        cella.value=f'=IFERROR(ROUND((AG{contaR}+AA{contaR}+T{contaR})/Y{contaR},0),"")'
                    contaC+=1
                contaR+=1    
    
            foglio.column_dimensions.group("a","n",hidden=True)

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.success("✅ Analisi completata!")
        st.download_button(
            label="📥 Scarica sottoscorta.xlsx",
            data=final_output,
            file_name="sottoscorta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )













