



import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import base64
import numpy as np

st.set_page_config(layout="wide", page_title="Nesting Programm Stange Version 1.5", page_icon="🌲")

def render_center_logo(png_path: str, width_px: int = 220):
    try:
        with open(png_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        st.markdown(
            f"<div style='text-align:center;margin-top:6px;margin-bottom:4px;'>"
            f"<img src='data:image/png;base64,{b64}' style='width:{width_px}px;border-radius:18px;box-shadow:0 6px 24px rgba(0,0,0,0.18);'/>"
            f"</div>",
            unsafe_allow_html=True
        )
    except Exception:
        st.markdown("<h3 style='text-align:center;'>🌲</h3>", unsafe_allow_html=True)

LOGO_PATH = "7a14a898-5611-46ed-b775-a15a4347ad93.png"
render_center_logo(LOGO_PATH, width_px=220)

st.markdown("<div style='text-align:center;'><h1 style='margin-bottom:0;'>Nesting Programm Stange Version 1.5</h1><p style='color:#666;font-style:italic;margin-top:2px;'>created by SPE – 'Machs dir selbst'</p></div>", unsafe_allow_html=True)
st.markdown("---")

# Parameters + Upload
col1, col2 = st.columns([3,2])
with col1:
    min_laenge = st.number_input("Minimale Kantholzlänge (m)", value=4.0)
    max_laenge = st.number_input("Maximale Kantholzlänge (m)", value=12.0)
    verschnitt = st.number_input("Zuschlag Verschnitt (m)", value=0.1)
    breiten_zuschlag = st.number_input("Zuschlag Breitenoptimierung (m)", value=0.041)
with col2:
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])

# Session Defaults
for k in ("master_df","et1_result","et2_result","et3_result"):
    if k not in st.session_state: st.session_state[k]=None
if "changed_uids" not in st.session_state: st.session_state.changed_uids=set()
if "assigned" not in st.session_state: st.session_state.assigned={"et1":set(),"et2":set(),"et3":set()}
if "source_sig" not in st.session_state: st.session_state.source_sig=None
if "change_log" not in st.session_state: st.session_state.change_log=[]
if "et1_undo_stack" not in st.session_state: st.session_state.et1_undo_stack=[]  # Undo-Stack für manuelle Breitenoptimierung

def safe_rerun():
    if hasattr(st,"rerun"): st.rerun()
    elif hasattr(st,"experimental_rerun"): st.experimental_rerun()

def color_status(v):
    if v=="✅": return "background-color:#b6f2bb;color:black;"
    if v=="🟠": return "background-color:#ffe29f;color:black;"
    if v=="❌": return "background-color:#ffb3b3;color:black;"
    return ""

def compute_status_symbol(lf, uid, dim_x=None):
    if uid in st.session_state.changed_uids: return "🟠"
    try:
        if dim_x is not None and not pd.isna(dim_x):
            return "✅" if float(min_laenge) <= float(dim_x) <= float(max_laenge) else "❌"
        return "✅" if float(lf) >= float(min_laenge) else "❌"
    except Exception:
        return "❌"

def compute_status_view(df):
    out=df.copy()
    length_col="Länge_f" if "Länge_f" in out.columns else "Länge"
    dim_col="Dimension X" if "Dimension X" in out.columns else None
    out["Unvollständig"]=[
        compute_status_symbol(out[length_col].iloc[i], out["UID"].iloc[i],
                              out[dim_col].iloc[i] if dim_col else None)
        for i in range(len(out))
    ]
    return out

def load_uploaded_to_master(uploaded):
    df=pd.read_excel(uploaded,dtype=str)
    for c in("Länge","Höhe","Breite"):
        if c not in df.columns: df[c]=pd.NA
        df[f"{c}_f"]=pd.to_numeric(df[c],errors="coerce")
    for c in("Pak","PB","Bezeichnung","GES"):
        if c not in df.columns: df[c]=""
    df["PB."]=df["PB"]
    df["UID"]=df["PB"].astype(str)+"_"+df["Bezeichnung"].astype(str)+"_"+df.index.astype(str)
    if "Pak" in df.columns:
        df["Pak_num"]=df["Pak"].astype(str).str.extract(r"(\d+)").astype("Int64")
    else:
        df["Pak_num"]=pd.NA
    return df

def sort_ui(df_pool, et_key):
    sort_cols=[c for c in["PB","Pak_num","GES","Bezeichnung","Höhe_f"] if c in df_pool.columns]
    chosen=st.multiselect(f"Sortierspalten {et_key}",sort_cols,default=[c for c in("PB","Pak_num") if c in sort_cols],key=f"sort_{et_key}")
    directions={s:st.selectbox(f"Sortierrichtung {et_key} {s}",["Absteigend","Aufsteigend"],key=f"dir_{et_key}_{s}") for s in chosen}
    if chosen:
        asc=[(directions[s]=="Aufsteigend") for s in chosen]
        return df_pool.sort_values(by=chosen,ascending=asc).reset_index(drop=True)
    return df_pool.reset_index(drop=True)

def pool_for_et(et_num:int):
    master=st.session_state.master_df
    assigned=st.session_state.assigned
    exclude=set().union(*[assigned[k] for k in assigned if k!=f"et{et_num}"])
    return master[~master["UID"].isin(exclude)].copy()

def sync_change(uid,new_len,new_breite):
    m=st.session_state.master_df["UID"]==uid
    if not m.any(): return
    old_len=st.session_state.master_df.loc[m,"Länge"].values[0]
    old_breite=st.session_state.master_df.loc[m,"Breite"].values[0]
    pak=st.session_state.master_df.loc[m,"Pak"].values[0]
    st.session_state.changed_uids.add(uid)
    st.session_state.master_df.loc[m,["Länge","Breite"]]=[new_len,new_breite]
    st.session_state.master_df.loc[m,["Länge_f","Breite_f"]]=[pd.to_numeric(new_len,errors="coerce"),
                                                              pd.to_numeric(new_breite,errors="coerce")]
    st.session_state.change_log.append({
        "Zeit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Aktion": "Edit Zeile",
        "UID": uid, "Pak": pak,
        "Alte Länge": old_len, "Alte Breite": old_breite,
        "Neue Länge": new_len, "Neue Breite": new_breite
    })

def reset_change(uid):
    logs=[x for x in st.session_state.change_log if x.get("UID")==uid and x.get("Aktion")=="Edit Zeile"]
    if not logs: return
    last=logs[-1]
    m=st.session_state.master_df["UID"]==uid
    if m.any():
        st.session_state.master_df.loc[m,["Länge","Breite"]]=[last["Alte Länge"],last["Alte Breite"]]
        st.session_state.master_df.loc[m,["Länge_f","Breite_f"]]=[pd.to_numeric(last["Alte Länge"],errors="coerce"),
                                                                  pd.to_numeric(last["Alte Breite"],errors="coerce")]
        if uid in st.session_state.changed_uids: st.session_state.changed_uids.remove(uid)
        # entferne nur den letzten Edit-Log für diese UID
        idx = len(st.session_state.change_log)-1
        while idx>=0:
            if st.session_state.change_log[idx].get("UID")==uid and st.session_state.change_log[idx].get("Aktion")=="Edit Zeile":
                st.session_state.change_log.pop(idx)
                break
            idx-=1
        st.success(f"🔁 {uid} zurückgesetzt."); safe_rerun()

# Datei laden
if uploaded_file:
    sig=(uploaded_file.name,uploaded_file.size)
    if st.session_state.source_sig!=sig:
        st.session_state.master_df=load_uploaded_to_master(uploaded_file)
        st.session_state.changed_uids=set()
        st.session_state.assigned={"et1":set(),"et2":set(),"et3":set()}
        st.session_state.et1_result=None; st.session_state.et2_result=None; st.session_state.et3_result=None
        st.session_state.change_log=[]; st.session_state.et1_undo_stack=[]
        st.session_state.source_sig=sig
        st.success("✅ Datei erfolgreich geladen – bereit zum Nesten")
else:
    st.stop()

st.markdown("---")

def nesting(df_nest,max_length,verschnitt_val,min_length,etappen_nr,
            breite_optimierung=False,standard_breite=None,breite_zuschlag_val=0.0):
    bins,groups,nesting_nr=[],[],1
    for _,row in df_nest.iterrows():
        l,b,h=row["Länge_f"],row["Breite_f"],row["Höhe_f"]
        l=float(l or 0); b=float(b or 0); h=float(h or 0)
        placed=False
        for i in range(len(bins)):
            curr=bins[i]
            if breite_optimierung and standard_breite:
                sum_b=sum(r["Breite_f"] for r in groups[i])+b+breite_zuschlag_val
                if sum_b<=standard_breite and (curr+l+verschnitt_val)<=max_length:
                    bins[i]+=l+verschnitt_val; groups[i].append(row); placed=True; break
            else:
                if abs(b-groups[i][0]["Breite_f"])<1e-6 and (curr+l+verschnitt_val)<=max_length:
                    bins[i]+=l+verschnitt_val; groups[i].append(row); placed=True; break
        if not placed:
            bins.append(l+verschnitt_val); groups.append([row])
    rows=[]
    for g in groups:
        total_l=sum((r["Länge_f"]+verschnitt_val) for r in g)
        max_b=max(r["Breite_f"] for r in g); max_h=max(r["Höhe_f"] for r in g)
        m2=round(total_l*max_b,3); m3=round(total_l*max_b*max_h/1000,3)
        for r in g:
            status="✅" if float(min_laenge)<=float(total_l)<=float(max_laenge) else "❌"
            if r["UID"] in st.session_state.changed_uids: status="🟠"
            rows.append({
                "Etappe":etappen_nr,"Nesting-Nr":nesting_nr,
                "PB":r.get("PB",""),"PB.":r.get("PB.",""),
                "GES":r.get("GES",""),"Pak":r.get("Pak",""),
                "Bauteil-Name":r.get("Bezeichnung",""),
                "Dimension Z":max_h,"Dimension Y":max_b,"Dimension X":round(total_l,3),
                "m2":m2,"m3":m3,"Breite":r["Breite"],"Länge":r["Länge"],
                "Unvollständig":status,"UID":r["UID"]
            })
        nesting_nr+=1
    return pd.DataFrame(rows)

def etappe_block(et_num:int,label,key):
    with st.expander(f"🧩 {label}", expanded=False):
        pool=pool_for_et(et_num)

        # 1) STANDARDBREITE wieder einführen NUR in Etappe 1:
        if et_num == 1:
            # wähle Standardbreite (diese wird aus dem Pool ausgeschlossen)
            try:
                std_opts = sorted(pd.to_numeric(pool["Breite_f"], errors="coerce").dropna().unique().tolist())
            except Exception:
                std_opts = []
            standard_breite = st.selectbox("Standardbreite (aus Pool ausschließen)", std_opts, index=0 if std_opts else None, key="std_breite_et1") if std_opts else None
            if standard_breite is not None:
                pool = pool[pd.to_numeric(pool["Breite_f"], errors="coerce") != float(standard_breite)].copy()
        # KEINE automatische Breitenoptimierung mehr:
        standard_breite = None
        breite_opt = False

        pool_sorted=sort_ui(pool,f"Et{et_num}")
        if st.button(f"▶️ Nesting {label} starten",key=f"btn_et{et_num}"):
            res=nesting(pool_sorted,max_laenge,verschnitt,min_laenge,et_num,
                        breite_optimierung=False,
                        standard_breite=None,
                        breite_zuschlag_val=0.0)
            st.session_state[key]=res
            st.session_state.assigned[f"et{et_num}"]=set(res["UID"])
            st.success(f"{label} abgeschlossen.")

        dfk=st.session_state.get(key)
        if dfk is not None and not dfk.empty:
            view=compute_status_view(dfk).rename(columns={"Dimension Z":"N. Z","Dimension Y":"N. Y","Dimension X":"N. X"})
            cols=["Etappe","Nesting-Nr","PB","GES","Pak","Bauteil-Name","N. Z","N. Y","N. X","Breite","Länge","Unvollständig","UID"]
            view=view[[c for c in cols if c in view.columns]]
            editable=st.data_editor(view,
                column_config={"Breite":st.column_config.NumberColumn("Breite",step=0.01),
                               "Länge":st.column_config.NumberColumn("Länge",step=0.01)},
                disabled=[c for c in view.columns if c not in ["Länge","Breite"]],
                use_container_width=True,key=f"edit_et{et_num}")
            # diffs
            base=dfk.set_index("UID")
            ed=editable.set_index("UID")
            common=ed.index.intersection(base.index)
            diff=(ed.loc[common,"Länge"]!=base.loc[common,"Länge"])|(ed.loc[common,"Breite"]!=base.loc[common,"Breite"])
            uids=list(common[diff])
            if uids:
                for uid in uids:
                    r=ed.loc[uid]
                    sync_change(uid, r["Länge"], r["Breite"])
                dfk.loc[dfk["UID"].isin(uids),["Länge","Breite"]]=ed.loc[uids,["Länge","Breite"]].values
                st.session_state[key]=dfk
                st.success(f"{len(uids)} Bauteil(e) angepasst (🟠)."); safe_rerun()

            # ------------------------------
            # 🪚 Manuelle Breitenoptimierung (nach Nesting, Etappe 1)
            # ------------------------------
            if et_num == 1:
                st.markdown("---")
                st.subheader("🪚 Manuelle Breitenoptimierung (nach Nesting)")
                st.caption("Kombiniere Breiten additiv innerhalb gleicher N. Z (Höhe): Gesamtbreite = Summe der gewählten Breiten + Zuschlag; Gesamtlänge = längstes Teil.")

                # 3) NUR gleiche N. Z dürfen kombiniert werden → erst N. Z wählen:
                try:
                    nz_unique = sorted(pd.to_numeric(dfk["Dimension Z"], errors="coerce").dropna().unique().tolist())
                except Exception:
                    nz_unique = []
                selected_nz = st.selectbox("N. Z (Höhe) auswählen:", nz_unique, index=0 if nz_unique else None, key="manual_nz_select")

                # Breitenliste nur für diese N. Z
                if selected_nz is not None:
                    rows_nz = dfk[pd.to_numeric(dfk["Dimension Z"], errors="coerce")==float(selected_nz)].copy()
                    try:
                        breite_values = pd.to_numeric(rows_nz["Breite"], errors="coerce").dropna().unique()
                    except Exception:
                        breite_values = np.array([])
                else:
                    rows_nz = dfk.copy()
                    breite_values = np.array([])

                if len(breite_values) == 0:
                    st.info("Keine gültigen Breitenwerte (für diese N. Z) gefunden.")
                else:
                    breiten_liste = sorted(breite_values.tolist())
                    auswahl = st.multiselect("Breiten auswählen (werden zusammen optimiert):", breiten_liste, key="manual_breiten_select")

                    zuschlag_val = st.number_input("Zuschlag Breitenoptimierung (m)", min_value=0.0, max_value=1.0, value=float(breiten_zuschlag), step=0.001, key="manual_breiten_zuschlag")

                    # 2) Undo-Button SOLL IMMER SICHTBAR SEIN
                    if st.button("↩️ Letzte manuelle Optimierung rückgängig machen", key="undo_manual_breiten_always"):
                        if st.session_state.et1_undo_stack:
                            st.session_state[key] = st.session_state.et1_undo_stack.pop()
                            # den letzten manuellen Protokolleintrag entfernen
                            idx = len(st.session_state.change_log) - 1
                            while idx >= 0:
                                if st.session_state.change_log[idx].get("Aktion") == "Manuelle Breitenoptimierung":
                                    st.session_state.change_log.pop(idx)
                                    break
                                idx -= 1
                            st.warning("Letzte manuelle Breitenoptimierung rückgängig gemacht.")
                            safe_rerun()
                        else:
                            st.info("Kein vorheriger Zustand vorhanden.")

                    if auswahl:
                        # Nimm ALLE Zeilen innerhalb N. Z mit einer der ausgewählten Breiten
                        selected_rows = rows_nz[pd.to_numeric(rows_nz["Breite"], errors="coerce").isin(auswahl)].copy()

                        # 4) Name inkl. Pak-Nummern bauen (ein Eintrag pro ZEILE, nicht nur pro Breitenwert)
                        pair_list = []
                        for _, rr in selected_rows.iterrows():
                            bval = float(pd.to_numeric(rr["Breite"], errors="coerce"))
                            pakv = str(rr.get("Pak","")).strip()
                            pair_list.append(f"{bval:.3f} Pak{pakv if pakv else '-'}")
                        name_pak = " + ".join(pair_list)

                        # Additive Breite = Summe aller ausgewählten ZEILEN + Zuschlag
                        neue_breite = round(pd.to_numeric(selected_rows["Breite"], errors="coerce").sum() + float(zuschlag_val), 3)

                        # Länge = max der Original-Längen der ausgewählten ZEILEN
                        neue_laenge = float(pd.to_numeric(selected_rows["Länge"], errors="coerce").max())

                        st.success("✅ Manuelle Breitenoptimierung berechnet")
                        st.write("**Ausgewählte Teile:**", ", ".join(selected_rows.get("Bauteil-Name","").astype(str)))
                        st.write(f"**Neue Gesamtbreite:** {neue_breite:.3f} m")
                        st.write(f"**Neue Gesamtlänge:** {neue_laenge:.3f} m (vom längsten Teil)")
                        st.write(f"**Kombination:** Optimiert ({name_pak})")

                        if st.button("Änderung übernehmen", key="apply_manual_breiten"):
                            # Undo-Stack: kompletten Zustand sichern
                            st.session_state.et1_undo_stack.append(dfk.copy())

                            # Protokoll-Eintrag
                            entry = {
                                "Zeit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Aktion": "Manuelle Breitenoptimierung",
                                "Breitenkombination": name_pak,
                                "Zuschlag": float(zuschlag_val),
                                "Neue Breite": neue_breite,
                                "Neue Laenge": neue_laenge,
                                "N.Z": float(selected_nz)
                            }
                            st.session_state.change_log.append(entry)

                            # Entferne ausgewählte Zeilen aus dfk (nur die in rows_nz + matching Breite)
                            remaining_mask = ~(
                                (pd.to_numeric(dfk["Dimension Z"], errors="coerce")==float(selected_nz)) &
                                (pd.to_numeric(dfk["Breite"], errors="coerce").isin(auswahl))
                            )
                            remaining = dfk[remaining_mask].copy()

                            # Neue kombinierte Zeile einfügen
                            new_row = {
                                "Etappe": 1,
                                "Nesting-Nr": (remaining["Nesting-Nr"].max() + 1) if "Nesting-Nr" in remaining.columns and not remaining.empty else 1,
                                "PB": "OPT",
                                "PB.": "",
                                "GES": "",
                                "Pak": "",  # bewusst leer, der Name enthält Pak-Infos
                                "Bauteil-Name": f"Optimiert ({name_pak})",
                                "Dimension Z": float(selected_nz),
                                "Dimension Y": neue_breite,
                                "Dimension X": neue_laenge,
                                "m2": round(neue_breite * neue_laenge, 3),
                                "m3": pd.NA,
                                "Breite": neue_breite,
                                "Länge": neue_laenge,
                                "Unvollständig": "🟠",  # markiert als geändert
                                "UID": f"MANOPT_{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
                            }
                            remaining = pd.concat([remaining, pd.DataFrame([new_row])], ignore_index=True)

                            st.session_state[key] = remaining
                            st.success("Änderung übernommen – Nestingliste aktualisiert.")
                            safe_rerun()
        else:
            st.info("Noch keine Nesting-Ergebnisse in dieser Etappe.")

# Etappen
etappe_block(1,"Etappe 1 – Sonderbreiten","et1_result")
etappe_block(2,"Etappe 2 – Nach Längen","et2_result")
etappe_block(3,"Etappe 3 – Restliches Nesting","et3_result")

st.markdown("---")

# Master-Übersicht
st.subheader("📋 Master-Übersicht (nur geänderte Bauteile)")
changed=st.session_state.master_df[st.session_state.master_df["UID"].isin(st.session_state.changed_uids)]
if changed.empty:
    st.info("Keine manuell angepassten Bauteile.")
else:
    st.dataframe(compute_status_view(changed), use_container_width=True)

st.markdown("---")

# Änderungslog
st.subheader("🕓 Änderungsprotokoll")
if not st.session_state.change_log:
    st.info("Noch keine Änderungen vorgenommen.")
else:
    df_log=pd.DataFrame(st.session_state.change_log)

    # Anzeige
    for i, row in df_log.iterrows():
        c1,c2=st.columns([8,1])
        with c1:
            aktion = row.get("Aktion","Edit Zeile")
            if aktion=="Edit Zeile":
                st.write(
                    f"**Pak {row.get('Pak','')}** | UID `{row.get('UID','')}` | ⏱️ {row['Zeit']}  \n"
                    f"**Länge:** {row.get('Alte Länge','?')} → **{row.get('Neue Länge','?')}** | "
                    f"**Breite:** {row.get('Alte Breite','?')} → **{row.get('Neue Breite','?')}**"
                )
            elif aktion=="Manuelle Breitenoptimierung":
                st.write(
                    f"**{aktion}** | ⏱️ {row['Zeit']}  \n"
                    f"**Kombination:** {row.get('Breitenkombination','')}  \n"
                    f"**Zuschlag:** {row.get('Zuschlag','')}  \n"
                    f"**Neue Breite:** {row.get('Neue Breite','')} | **Neue Länge:** {row.get('Neue Laenge','')} | **N. Z:** {row.get('N.Z','')}"
                )
        with c2:
            if aktion=="Edit Zeile" and "UID" in row and isinstance(row["UID"], str) and row["UID"]:
                if st.button("🔁 Reset", key=f"reset_{i}_{row['UID']}"):
                    reset_change(row["UID"])

    # Export als Excel Änderungslog
    ts=datetime.now().strftime("%Y-%m-%d_%H-%M")
    buf=BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w:
        df_log.to_excel(w,index=False,sheet_name="Änderungslog")
    st.download_button("📤 Änderungslog exportieren", data=buf.getvalue(),
                       file_name=f"aenderungslog_{ts}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")

# Gesamtübersicht + Export (mit Editier-Funktion)
st.header("📊 Gesamtübersicht")
results=[]
for key,et in (("et1_result",1),("et2_result",2),("et3_result",3)):
    dfk=st.session_state.get(key)
    if dfk is not None and not dfk.empty:
        t=dfk.copy(); t["Etappe"]=et; results.append(t)
if results:
    gesamt=pd.concat(results,ignore_index=True).drop_duplicates(subset=["UID"])
    gesamt=gesamt.sort_values(by=["Etappe","Nesting-Nr"]).reset_index(drop=True)

    gesamt_view=gesamt.rename(columns={"Dimension Z":"N. Z","Dimension Y":"N. Y","Dimension X":"N. X","m2":"N.m2","m3":"N.m3"})
    # 5) Gesamtübersicht editierbar: PB, GES, Bauteil-Name, Pak
    editable_cols = ["PB","GES","Bauteil-Name","Pak"]
    for c in ["Länge","Breite","N. Z","N. Y","N. X","UID","Etappe","Nesting-Nr","PB."]:
        if c not in gesamt_view.columns: gesamt_view[c]=np.nan

    # Anzeige-Set
    cols_display=["Etappe","Nesting-Nr","PB","GES","Pak","Bauteil-Name","N. Z","N. Y","N. X","Breite","Länge","Unvollständig","UID"]
    display_df = gesamt_view[[c for c in cols_display if c in gesamt_view.columns]].copy()

    # Editor
    ed = st.data_editor(
        display_df,
        use_container_width=True,
        disabled=[c for c in display_df.columns if c not in editable_cols],
        column_config={
            "PB": st.column_config.TextColumn("PB"),
            "GES": st.column_config.TextColumn("GES"),
            "Pak": st.column_config.TextColumn("Pak"),
            "Bauteil-Name": st.column_config.TextColumn("Bauteil-Name"),
        },
        key="gesamt_editor"
    )

    # Änderungen in Etappen-Resultate zurückschreiben per UID
    if not ed.equals(display_df):
        # finde geänderte Zeilen in editierbaren Feldern
        changed_rows = []
        for idx in ed.index:
            row_old = display_df.loc[idx, editable_cols]
            row_new = ed.loc[idx, editable_cols]
            if not row_old.equals(row_new):
                changed_rows.append(idx)
        if changed_rows:
            for idx in changed_rows:
                uid = ed.loc[idx, "UID"]
                for et_key in ("et1_result","et2_result","et3_result"):
                    df_src = st.session_state.get(et_key)
                    if df_src is not None and not df_src.empty and uid in df_src["UID"].values:
                        for col in editable_cols:
                            if col in df_src.columns:
                                df_src.loc[df_src["UID"]==uid, col] = ed.loc[idx, col]
                        st.session_state[et_key] = df_src
                        break
            st.success(f"{len(changed_rows)} Zeile(n) in der Gesamtübersicht übernommen.")

    # ---------------------------------------------------
    # ✅ NEUER EXPORTBLOCK – angepasst wie gewünscht
    # ---------------------------------------------------

    # B.m2 / B.m3 neu berechnen (auf Basis der m-Werte)
    gesamt_view["B.m2"]=round(pd.to_numeric(gesamt_view["Länge"],errors="coerce")*pd.to_numeric(gesamt_view["Breite"],errors="coerce"),3)
    gesamt_view["B.m3"]=round(gesamt_view["B.m2"]*pd.to_numeric(gesamt_view["N. Z"],errors="coerce")/1000,3)

    # Export-Kopie anlegen
    export_df = gesamt_view.copy()

    # Falls "Höhe" fehlt: aus N. Z übernehmen (bleibt UNVERÄNDERT)
    if "Höhe" not in export_df.columns and "N. Z" in export_df.columns:
        export_df["Höhe"] = export_df["N. Z"]

    # Nur diese Spalten in mm umrechnen
    for col in ["N. Y","N. X","Breite","Länge"]:
        if col in export_df.columns:
            export_df[col] = (
                pd.to_numeric(export_df[col], errors="coerce").mul(1000)
            ).round(0).astype("Int64")

    # N. Z und Höhe bleiben unverändert!

    # Spalten in exakt gewünschter Reihenfolge
    cols_export=["Etappe","Nesting-Nr","PB","N. Z","N. Y","N. X","N.m2","N.m3","GES","Bauteil-Name","Pak","PB.","Höhe","Breite","Länge","B.m2","B.m3","Unvollständig","UID"]
    gesamt_export=export_df[[c for c in cols_export if c in export_df.columns]]

    ts=datetime.now().strftime("%Y-%m-%d_%H-%M")
    buffer=BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as w:
        gesamt_export.to_excel(w, index=False, sheet_name="Ergebnis")
    st.download_button("📤 Export Gesamtergebnis", data=buffer.getvalue(),
                       file_name=f"nesting_gesamt_{ts}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Noch keine Nestings vorhanden.")



