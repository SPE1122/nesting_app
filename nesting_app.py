# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import base64

# ========= Meta =========
st.set_page_config(layout="wide", page_title="Nesting Programm Stange 2.0 + Etappe 0 + Statistik", page_icon="🌲")

def render_center_logo(png_path: str, width_px: int = 220):
    try:
        with open(png_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        st.markdown(
            "<div style='text-align:center;margin-top:6px;margin-bottom:4px;'>"
            f"<img src='data:image/png;base64,{b64}' style='width:{width_px}px;border-radius:18px;box-shadow:0 6px 24px rgba(0,0,0,0.18);'/>"
            "</div>",
            unsafe_allow_html=True
        )
    except Exception:
        st.markdown("<h3 style='text-align:center;'>🌲</h3>", unsafe_allow_html=True)

LOGO_PATH = "7a14a898-5611-46ed-b775-a15a4347ad93.png"
render_center_logo(LOGO_PATH, width_px=220)  # Breite bei Bedarf anpassen

st.markdown("<div style='text-align:center;'><h1 style='margin-bottom:0;'>Nesting Programm Stange 2.0</h1><p style='color:#666;font-style:italic;margin-top:2px;'>created by SPE – 'Machs dir selbst'</p></div>", unsafe_allow_html=True)
st.markdown(
    "<div style='text-align:center; margin: 4px 0 8px 0;'>"
    "Legende: ✅ innerhalb Vorgaben & unverändert &nbsp;|&nbsp; 🟠 manuell geändert &nbsp;|&nbsp; ❌ außerhalb Vorgaben &nbsp;|&nbsp; 🟣 Bestand (Lager-Rest)"
    "</div>",
    unsafe_allow_html=True
)
st.markdown("---")

# ========= Params + Upload =========
col1, col2 = st.columns([3,2])
with col1:
    min_laenge = st.number_input("Minimale Kantholzlänge (m)", value=4.0)
    max_laenge = st.number_input("Maximale Kantholzlänge (m)", value=12.0)
    verschnitt = st.number_input("Zuschlag Verschnitt (m, 1× pro Nest)", value=0.1, step=0.001, format="%.3f")
    breiten_zuschlag = st.number_input("Zuschlag Breitenoptimierung (m)", value=0.041, step=0.001, format="%.3f")
with col2:
    uploaded_file = st.file_uploader("Excel: Master (PB, Bezeichnung, Höhe, Breite, Länge, Pak, GES)", type=["xlsx"])
    up_reste = st.file_uploader("Excel: Lager-Reste (Nr, Höhe, Breite, Länge, [optional] Bezeichnung)", type=["xlsx"])

# ========= Session =========
defaults = {
    "master_df": None, "et1_result": None, "et2_result": None, "et3_result": None, "et_dyn": pd.DataFrame(),
    "changed_uids": set(), "assigned": {"et0": set(), "et1": set(), "et2": set(), "et3": set()},
    "source_sig": None, "change_log": [], "et1_undo_stack": [],
    "dyn_etappen": [],
    # Reste
    "rest_assignments": [], "rest_consumed_uids": set(), "lager_df": None,
}
for k,v in defaults.items():
    if k not in st.session_state: st.session_state[k]=v

def safe_rerun():
    if hasattr(st,"rerun"): st.rerun()
    elif hasattr(st,"experimental_rerun"): st.experimental_rerun()

# ========= Loaders =========
def load_master(uploaded):
    df=pd.read_excel(uploaded,dtype=str)
    for c in("PB","Bezeichnung","Höhe","Breite","Länge","Pak","GES"):
        if c not in df.columns: df[c]=""
    for c in("Höhe","Breite","Länge"):
        df[f"{c}_f"]=pd.to_numeric(df[c],errors="coerce")
    df["PB."]=df["PB"]
    try:
        df["Pak_num"]=df["Pak"].astype(str).str.extract(r"(\d+)").astype("Int64")
    except Exception:
        df["Pak_num"]=pd.NA
    df["UID"]=df["PB"].astype(str)+"_"+df["Bezeichnung"].astype(str)+"_"+df.index.astype(str)
    return df

def load_reste(uploaded):
    df = pd.read_excel(uploaded, dtype=str)
    rename = {"Nr":"Nr","nr":"Nr","NR":"Nr","Höhe":"Höhe","Hoehe":"Höhe","Z":"Höhe",
              "Breite":"Breite","Y":"Breite","Länge":"Länge","Laenge":"Länge","X":"Länge"}
    df = df.rename(columns={k:v for k,v in rename.items() if k in df.columns})
    for c in ("Nr","Höhe","Breite","Länge"):
        if c not in df.columns: df[c]=""
    for c in ("Höhe","Breite","Länge"):
        df[f"{c}_f"] = pd.to_numeric(df[c], errors="coerce")
    df["rest_laenge_f"] = pd.to_numeric(df.get("Länge_f"), errors="coerce")
    df["verbraucht"] = False
    return df

# ========= Status/Helper =========
def compute_status_symbol(lf, uid, dim_x=None):
    if uid in st.session_state.changed_uids: return "🟠"
    try:
        eps = 1e-9
        if dim_x is not None and not pd.isna(dim_x):
            return "✅" if (float(min_laenge)-eps) <= float(dim_x) <= (float(max_laenge)+eps) else "❌"
        return "✅" if float(lf) >= float(min_laenge)-eps else "❌"
    except Exception:
        return "❌"

def compute_status_view(df):
    out=df.copy()
    length_col="Länge_f" if "Länge_f" in out.columns else "Länge"
    dim_col="Dimension X" if "Dimension X" in out.columns else None
    out["Unvollständig"]=[
        compute_status_symbol(out[length_col].iloc[i], out["UID"].iloc[i], out[dim_col].iloc[i] if dim_col else None)
        for i in range(len(out))
    ]
    return out

def sort_ui(df_pool, et_key):
    df = df_pool.copy()
    cand = [c for c in ["PB","Pak","Pak_num","GES","Bezeichnung","Höhe_f","Länge_f","Dimension X","N. X"] if c in df.columns]
    sel = st.multiselect(f"Sortierspalten {et_key}", cand, default=["PB","Pak_num"] if "Pak_num" in cand else [])
    dirs = {s: st.selectbox(f"Richtung {et_key} {s}", ["Absteigend","Aufsteigend"], key=f"dir_{et_key}_{s}") for s in sel}
    if not sel: return df.reset_index(drop=True)
    by, asc = [], []
    for c in sel:
        mapped = 'Pak_num' if c=='Pak' and 'Pak_num' in df.columns else ('Dimension X' if c=='N. X' and 'Dimension X' in df.columns else c)
        by.append(mapped)
        asc.append(dirs[c]=="Aufsteigend")
    return df.sort_values(by=by, ascending=asc, kind="mergesort").reset_index(drop=True)

def pool_for_et(et_num:int):
    master=st.session_state.master_df
    exclude=set().union(*[v for k,v in st.session_state.assigned.items() if k!=f"et{et_num}"])
    return master[~master["UID"].isin(exclude)].copy()

def sync_change(uid,new_len,new_breite):
    m=st.session_state.master_df["UID"]==uid
    if not m.any(): return
    old_len=st.session_state.master_df.loc[m,"Länge"].values[0]
    old_breite=st.session_state.master_df.loc[m,"Breite"].values[0]
    pak=st.session_state.master_df.loc[m,"Pak"].values[0]
    st.session_state.changed_uids.add(uid)
    st.session_state.master_df.loc[m,["Länge","Breite"]]=[new_len,new_breite]
    st.session_state.master_df.loc[m,["Länge_f","Breite_f"]]=[pd.to_numeric(new_len,errors="coerce"), pd.to_numeric(new_breite,errors="coerce")]
    st.session_state.change_log.append({
        "Zeit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Aktion": "Edit Zeile",
        "UID": uid, "Pak": pak,
        "Alte Länge": old_len, "Alte Breite": old_breite,
        "Neue Länge": new_len, "Neue Breite": new_breite
    })

# ========= Reste (Etappe 0) =========
def compat_parts_for_rest(master_pool, rest_row):
    Z = float(pd.to_numeric(rest_row.get("Höhe_f", np.nan), errors="coerce") or 0.0)
    Y = float(pd.to_numeric(rest_row.get("Breite_f", np.nan), errors="coerce") or 0.0)
    X = float(pd.to_numeric(rest_row.get("rest_laenge_f", np.nan), errors="coerce") or 0.0)
    df = master_pool.copy()
    for c in ("Höhe_f","Breite_f","Länge_f"):
        df[c] = pd.to_numeric(df.get(c), errors="coerce")
    m = (df["Höhe_f"]==Z) & (df["Breite_f"]<=Y+1e-12) & (df["Länge_f"]<=X+1e-12)
    return df[m].copy()

def assign_rest(lager_df, rest_idx, uid, x_consume, Z, Y):
    prev = float(lager_df.loc[rest_idx, "rest_laenge_f"])
    lager_df.loc[rest_idx, "rest_laenge_f"] = round(prev - float(x_consume), 6)
    if lager_df.loc[rest_idx, "rest_laenge_f"] <= 1e-6:
        lager_df.loc[rest_idx, "verbraucht"] = True
    st.session_state.rest_assignments.append({
        "Zeit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Rest_Nr": lager_df.loc[rest_idx, "Nr"],
        "UID": uid, "Z": float(Z), "Y": float(Y), "X_verbraucht": float(x_consume),
        "rest_idx": int(rest_idx),
    })
    st.session_state.rest_consumed_uids.add(uid)
    st.session_state.assigned.setdefault("et0", set()).add(uid)

def undo_last_assignment():
    if not st.session_state.rest_assignments: return False
    last = st.session_state.rest_assignments.pop()
    idx = last["rest_idx"]; add_back = float(last["X_verbraucht"])
    st.session_state.lager_df.loc[idx, "rest_laenge_f"] = float(st.session_state.lager_df.loc[idx, "rest_laenge_f"]) + add_back
    st.session_state.lager_df["verbraucht"] = st.session_state.lager_df["rest_laenge_f"] <= 1e-6
    uid = last.get("UID")
    st.session_state.rest_consumed_uids.discard(uid)
    st.session_state.assigned.setdefault("et0", set()).discard(uid)
    return True

# ========= Nesting =========
def nesting(df_nest,max_length,verschnitt_val,min_length,etappen_nr,
            breite_optimierung=False,standard_breite=None,breite_zuschlag_val=0.0):
    sums, groups, nesting_nr = [], [], 1
    for _,row in df_nest.iterrows():
        l = float(pd.to_numeric(row["Länge_f"], errors="coerce") or 0)
        b = float(pd.to_numeric(row["Breite_f"], errors="coerce") or 0)
        h = float(pd.to_numeric(row["Höhe_f"], errors="coerce") or 0)
        placed = False
        for i in range(len(groups)):
            same_b = abs(b - groups[i][0]["Breite_f"]) < 1e-6
            if breite_optimierung and standard_breite:
                sum_b=sum(pd.to_numeric(r["Breite_f"],errors="coerce") for r in groups[i])+b+breite_zuschlag_val
                if sum_b<=standard_breite and (sums[i]+l+verschnitt_val)<=max_length:
                    sums[i]+=l; groups[i].append(row); placed=True; break
            else:
                if same_b and (sums[i]+l+verschnitt_val)<=max_length:
                    sums[i]+=l; groups[i].append(row); placed=True; break
        if not placed:
            sums.append(l); groups.append([row])

    rows=[]
    for idx,g in enumerate(groups):
        total_l=sums[idx]+(verschnitt_val if len(g)>0 else 0.0)
        max_b=max(pd.to_numeric(r["Breite_f"],errors="coerce") for r in g)
        max_h=max(pd.to_numeric(r["Höhe_f"],errors="coerce") for r in g)
        m2=round(total_l*max_b,3); m3=round(m2*max_h/1000,3)
        status_ok=(float(min_laenge)<=float(total_l)<=float(max_laenge))
        for r in g:
            status="✅" if status_ok else "❌"
            if r["UID"] in st.session_state.changed_uids: status="🟠"
            rows.append({
                "Etappe":etappen_nr,"Nesting-Nr":nesting_nr,
                "PB":r.get("PB",""),"PB.":r.get("PB.",""),"GES":r.get("GES",""),"Pak":r.get("Pak",""),
                "Bauteil-Name":r.get("Bezeichnung",""),
                "Dimension Z":max_h,"Dimension Y":max_b,"Dimension X":round(total_l,3),
                "m2":m2,"m3":m3,"Breite":r["Breite"],"Länge":r["Länge"],
                "Unvollständig":status,"UID":r["UID"]
            })
        nesting_nr+=1
    return pd.DataFrame(rows)

# ========= Statistik-Helfer =========
def compute_verschnitt_stats_split(df_res: pd.DataFrame, max_len: float) -> dict:
    if df_res is None or len(df_res)==0:
        return {"saw_m2":0.0,"saw_m3":0.0,"pct_saw":0.0,"teile_m2":0.0,"teile_m3":0.0,"gruppen":pd.DataFrame()}
    t = df_res.copy()
    for col in ["Dimension X","Dimension Y","Dimension Z","Länge","Breite"]:
        if col not in t.columns: t[col]=pd.NA
    t["nx"] = pd.to_numeric(t["Dimension X"], errors="coerce")
    t["ny"] = pd.to_numeric(t["Dimension Y"], errors="coerce")
    t["nz"] = pd.to_numeric(t["Dimension Z"], errors="coerce")
    t["l_i"] = pd.to_numeric(t["Länge"], errors="coerce")
    t["b_i"] = pd.to_numeric(t["Breite"], errors="coerce")
    grp = t.groupby("Nesting-Nr", as_index=False).agg(nx=("nx","max"), ny=("ny","max"), nz=("nz","max"))
    teile = t.assign(area=lambda df: df["l_i"]*df["b_i"],
                     sumL=lambda df: df["l_i"]).groupby("Nesting-Nr", as_index=False).agg(
        teile_m2=("area","sum"), sumL=("sumL","sum")
    )
    grp = grp.merge(teile, on="Nesting-Nr", how="left").fillna({"teile_m2":0.0,"sumL":0.0})
    grp["teile_m3"] = grp["teile_m2"] * grp["nz"] / 1000.0
    grp["saw_m2"] = (grp["nx"] - grp["sumL"]).clip(lower=0) * grp["ny"]
    grp["saw_m3"] = grp["saw_m2"] * grp["nz"] / 1000.0
    saw_m2 = float(grp["saw_m2"].sum()); saw_m3=float(grp["saw_m3"].sum())
    teile_m2=float(grp["teile_m2"].sum()); teile_m3=float(grp["teile_m3"].sum())
    denom = float((grp["nx"]*grp["ny"]).sum())
    pct_saw = (100.0 * saw_m2 / denom) if denom>0 else 0.0
    return {"saw_m2":round(saw_m2,3),"saw_m3":round(saw_m3,3),"pct_saw":round(pct_saw,2),
            "teile_m2":round(teile_m2,3),"teile_m3":round(teile_m3,3),"gruppen":grp}

# ========= Load =========
if uploaded_file:
    sig=(uploaded_file.name,uploaded_file.size)
    if st.session_state.source_sig!=sig:
        st.session_state.master_df=load_master(uploaded_file)
        st.session_state.changed_uids=set()
        st.session_state.assigned={"et0":set(), "et1":set(),"et2":set(),"et3":set()}
        st.session_state.et1_result=None; st.session_state.et2_result=None; st.session_state.et3_result=None; st.session_state.et_dyn=pd.DataFrame()
        st.session_state.change_log=[]; st.session_state.et1_undo_stack=[]
        st.session_state.rest_assignments=[]; st.session_state.rest_consumed_uids=set()
        st.session_state.source_sig=sig
        st.success("✅ Datei erfolgreich geladen – bereit zum Nesten")
if up_reste is not None:
    st.session_state.lager_df = load_reste(up_reste); st.success("Lager-Reste geladen.")

if st.session_state.master_df is None: st.stop()

st.markdown("---")

# ========= Etappe 0 UI =========
with st.expander("Etappe 0 – Lager-Reste (Rest auswählen ➜ passende Bauteile)", expanded=True):
    if st.session_state.lager_df is None or st.session_state.lager_df.empty:
        st.info("Keine Lager-Reste geladen.")
    else:
        df_r = st.session_state.lager_df.copy()
        for c in ("Höhe_f","Breite_f","Länge_f","rest_laenge_f"):
            if c in df_r.columns: df_r[c]=pd.to_numeric(df_r[c],errors="coerce").round(3)
        df_r["Status"]=np.where(df_r["verbraucht"],"✅ verbraucht","frei")
        df_r["🟣"]=""
        used_idx=set(a["rest_idx"] for a in st.session_state.rest_assignments)
        if used_idx: df_r.loc[list(used_idx),"🟣"]="🟣"
        cols_show=["🟣","Nr"]+(["Bezeichnung"] if "Bezeichnung" in df_r.columns else [])+["Höhe_f","Breite_f","rest_laenge_f","Länge_f","Status"]
        st.dataframe(df_r[[c for c in cols_show if c in df_r.columns]], use_container_width=True)

        show_only_free = st.checkbox("Nur freie Reste anzeigen", value=True, key="reste_only_free")
        df_pick = df_r[~df_r["verbraucht"]].copy() if show_only_free else df_r.copy()
        if df_pick.empty:
            st.warning("Kein freier Rest verfügbar."); selected_rest=None
        else:
            df_pick["label"]=df_pick.apply(lambda r: f"{'🟣 ' if r['🟣']=='🟣' else ''}Nr {r['Nr']} | Z={r['Höhe_f']:.0f} Y={r['Breite_f']:.3f} Rest_X={r['rest_laenge_f']:.3f}"+(f" | {r['Bezeichnung']}" if 'Bezeichnung' in df_pick.columns else ""),axis=1)
            sel=st.selectbox("Rest wählen", df_pick["label"]); selected_rest=df_pick.loc[df_pick["label"]==sel].iloc[0] if sel else None

        if selected_rest is not None:
            pool_master = st.session_state.master_df[~st.session_state.master_df["UID"].isin(st.session_state.assigned.get("et0", set()))].copy()
            compat = compat_parts_for_rest(pool_master, selected_rest)
            if compat.empty:
                st.info("Keine passenden Bauteile für diesen Rest.")
            else:
                compat["label"]=compat.apply(lambda r: f"{r['PB']} | Pak {r['Pak']} | {r['Bezeichnung']} | Z={int(r['Höhe_f'])} Y={r['Breite_f']:.3f} L={r['Länge_f']:.3f}",axis=1)
                picks=st.multiselect("Bauteile wählen (Mehrfach-Zuweisung)", compat["label"])
                chosen_rows=compat[compat["label"].isin(picks)].copy()
                if len(chosen_rows):
                    ridx=selected_rest.name; rest_len=float(selected_rest["rest_laenge_f"])
                    will_assign, skipped=[], []
                    for _, rr in chosen_rows.iterrows():
                        need=float(rr["Länge_f"])
                        if need <= rest_len + 1e-9: will_assign.append(rr); rest_len -= need
                        else: skipped.append(rr)
                    st.caption(f"Geplante Zuweisungen: {len(will_assign)} | Nicht passend (zu lang): {len(skipped)} | Rest danach ≈ {rest_len:.3f} m")
                    if st.button("➕ Zuweisen (alle passenden)"):
                        for rr in will_assign:
                            assign_rest(st.session_state.lager_df, ridx, rr["UID"], float(rr["Länge_f"]), float(rr["Höhe_f"]), float(rr["Breite_f"]))
                        st.success(f"{len(will_assign)} Bauteil(e) zugewiesen."); safe_rerun()
        if st.button("↩️ Letzte Zuweisung rückgängig", key="reste_undo_btn"):
            if undo_last_assignment(): st.warning("Rückgängig gemacht."); safe_rerun()
            else: st.info("Kein Eintrag.")

st.markdown("---")

# ========= Etappe 1/2 =========
def etappe_block(et_num:int,label,key):
    with st.expander(f"🧩 {label}", expanded=False):
        pool=pool_for_et(et_num)

        if et_num == 1:
            try:
                std_opts = sorted(pd.to_numeric(pool["Breite_f"], errors="coerce").dropna().unique().tolist())
            except Exception:
                std_opts = []
            standard_breite = st.selectbox("Standardbreite (aus Pool ausschließen)", options=[None] + std_opts, index=0, format_func=lambda v: "-" if v is None else f"{v}", key="std_breite_et1")
            if standard_breite is not None:
                pool = pool[pd.to_numeric(pool["Breite_f"], errors="coerce") != float(standard_breite)].copy()

        z_range = None
        if et_num == 2:
            c1, c2 = st.columns(2)
            with c1:
                zus_von = st.number_input("Zuschlag (von, m)", min_value=0.0, value=float(verschnitt), step=0.001, key="et2_zus_von")
            with c2:
                zus_bis = st.number_input("Zuschlag (bis, m)", min_value=zus_von, value=float(max(verschnitt, zus_von)), step=0.001, key="et2_zus_bis")
            z_range = (zus_von, zus_bis)

        pool_sorted=sort_ui(pool,f"Et{et_num}")
        if st.button(f"▶️ Nesting {label} starten",key=f"btn_et{et_num}"):
            z = verschnitt
            if et_num == 2 and z_range is not None: z = z_range[0]
            res=nesting(pool_sorted,max_laenge,z,min_laenge,et_num,breite_optimierung=False,standard_breite=None,breite_zuschlag_val=0.0)
            st.session_state[key]=res
            st.session_state.assigned[f"et{et_num}"]=set(res["UID"])
            st.success(f"{label} abgeschlossen.")

        dfk=st.session_state.get(key)
        if dfk is not None and not dfk.empty:
            view=compute_status_view(dfk).rename(columns={"Dimension Z":"N. Z","Dimension Y":"N. Y","Dimension X":"N. X"})
            cols=["Etappe","Nesting-Nr","PB","GES","Pak","Bauteil-Name","N. Z","N. Y","N. X","Breite","Länge","Unvollständig","UID"]
            view=view[[c for c in cols if c in view.columns]]
            column_cfg = {"Breite": st.column_config.NumberColumn("Breite", step=0.001, format="%.3f"),
                          "Länge":  st.column_config.NumberColumn("Länge",  step=0.01,  format="%.2f"),
                          "N. Y":   st.column_config.NumberColumn("N. Y",   format="%.3f")}
            editable=st.data_editor(view, column_config=column_cfg, disabled=[c for c in view.columns if c not in ["Länge","Breite"]], use_container_width=True, key=f"edit_et{et_num}")
            base=dfk.set_index("UID"); ed=editable.set_index("UID")
            common=ed.index.intersection(base.index)
            diff=(ed.loc[common,"Länge"]!=base.loc[common,"Länge"])|(ed.loc[common,"Breite"]!=base.loc[common,"Breite"])
            uids=list(common[diff])
            if uids:
                for uid in uids:
                    r=ed.loc[uid]; sync_change(uid, r["Länge"], r["Breite"])
                dfk.loc[dfk["UID"].isin(uids),["Länge","Breite"]]=ed.loc[uids,["Länge","Breite"]].values
                st.session_state[key]=dfk
                st.success(f"{len(uids)} Bauteil(e) angepasst (🟠)."); safe_rerun()

            # ---- Manuelle Breitenoptimierung (nur Etappe 1) ----
            if et_num == 1 and dfk is not None and not dfk.empty:
                st.markdown("---")
                st.subheader("🪚 Manuelle Breitenoptimierung (Etappe 1)")
                try:
                    nz_unique = sorted(pd.to_numeric(dfk["Dimension Z"], errors="coerce").dropna().unique().tolist())
                except Exception:
                    nz_unique = []
                selected_nz = st.selectbox("N. Z (Höhe) auswählen:", nz_unique, index=0 if nz_unique else None, key="manual_nz_select_et1")
                if selected_nz is not None:
                    rows_nz = dfk[pd.to_numeric(dfk["Dimension Z"], errors="coerce")==float(selected_nz)].copy()
                    try:
                        breite_values = pd.to_numeric(rows_nz["Breite"], errors="coerce").dropna().unique()
                    except Exception:
                        breite_values = np.array([])
                else:
                    rows_nz = dfk.copy(); breite_values = np.array([])

                if len(breite_values)==0:
                    st.caption("Keine gültigen Breitenwerte (für diese N. Z) gefunden.")
                else:
                    breiten_liste = sorted(breite_values.tolist())
                    auswahl = st.multiselect("Breiten auswählen (werden zusammen optimiert):", breiten_liste, key="manual_breiten_select_et1")
                    zuschlag_val = st.number_input("Zuschlag Breitenoptimierung (m)", min_value=0.0, max_value=1.0, value=float(breiten_zuschlag), step=0.001, key="manual_breiten_zuschlag_et1")

                    if auswahl:
                        selected_rows = rows_nz[pd.to_numeric(rows_nz["Breite"], errors="coerce").isin(auswahl)].copy()
                        pair_list = []
                        for _, rr in selected_rows.iterrows():
                            bval = float(pd.to_numeric(rr["Breite"], errors="coerce"))
                            pakv = str(rr.get("Pak",""))
                            pair_list.append(f"{bval:.3f} Pak{pakv if pakv else '-'}")
                        name_pak = " + ".join(pair_list)

                        neue_breite = round(pd.to_numeric(selected_rows["Breite"], errors="coerce").sum() + float(zuschlag_val), 3)
                        neue_laenge = float(pd.to_numeric(selected_rows["Länge"], errors="coerce").max())

                        st.write(f"**Neue Gesamtbreite:** {neue_breite:.3f} m  |  **Neue Gesamtlänge:** {neue_laenge:.3f} m")
                        if st.button("Änderung übernehmen (manuelle Breitenoptimierung)", key="apply_manual_breiten_et1"):
                            remaining_mask = ~(
                                (pd.to_numeric(dfk['Dimension Z'], errors='coerce')==float(selected_nz)) &
                                (pd.to_numeric(dfk['Breite'], errors='coerce').isin(auswahl))
                            )
                            remaining = dfk[remaining_mask].copy()
                            new_row = {
                                "Etappe": 1,
                                "Nesting-Nr": (remaining["Nesting-Nr"].max() + 1) if "Nesting-Nr" in remaining.columns and not remaining.empty else 1,
                                "PB": "OPT", "PB.": "", "GES": "", "Pak": "",
                                "Bauteil-Name": f"Optimiert ({name_pak})",
                                "Dimension Z": float(selected_nz),
                                "Dimension Y": neue_breite,
                                "Dimension X": neue_laenge,
                                "m2": round(neue_breite * neue_laenge, 3),
                                "m3": pd.NA,
                                "Breite": neue_breite,
                                "Länge": neue_laenge,
                                "Unvollständig": "🟠",
                                "UID": f"MANOPT_{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
                            }
                            remaining = pd.concat([remaining, pd.DataFrame([new_row])], ignore_index=True)
                            st.session_state[key] = remaining
                            st.success("Manuelle Breitenoptimierung übernommen."); safe_rerun()
        else:
            st.info("Noch keine Nesting-Ergebnisse in dieser Etappe.")

etappe_block(1,"Etappe 1 – Sonderbreiten","et1_result")
etappe_block(2,"Etappe 2 – Nach Längen","et2_result")

# ========= Dynamische Etappen =========
st.subheader("🧱 Dynamische Etappen")
with st.expander("➕ Etappe hinzufügen", expanded=False):
    name = st.text_input("Name der Etappe", value=f"Etappe {len(st.session_state.dyn_etappen)+3}")
    if st.button("Etappe anlegen"):
        st.session_state.dyn_etappen.append({"name": name, "key": f"dyn_{len(st.session_state.dyn_etappen)+1}"})
        st.success("Etappe hinzugefügt."); safe_rerun()

dyn_frames = []
for i, et in enumerate(st.session_state.dyn_etappen, start=1):
    lbl = et["name"]; key = et["key"]
    with st.expander(f"🧩 {lbl}", expanded=False):
        pool = pool_for_et(3+i)
        pool_sorted = sort_ui(pool, lbl)
        if st.button(f"▶️ Nesting {lbl} starten", key=f"btn_{key}"):
            res = nesting(pool_sorted, max_laenge, verschnitt, min_laenge, 3+i, breite_optimierung=False, standard_breite=None, breite_zuschlag_val=0.0)
            st.session_state[key] = res; st.success(f"{lbl} abgeschlossen.")
        dfk = st.session_state.get(key)
        if dfk is not None and not dfk.empty:
            view = compute_status_view(dfk).rename(columns={"Dimension Z":"N. Z","Dimension Y":"N. Y","Dimension X":"N. X"})
            cols = ["Etappe","Nesting-Nr","PB","GES","Pak","Bauteil-Name","N. Z","N. Y","N. X","Breite","Länge","Unvollständig","UID"]
            column_cfg = {"Breite": st.column_config.NumberColumn("Breite", step=0.001, format="%.3f"),
                          "Länge":  st.column_config.NumberColumn("Länge",  step=0.01,  format="%.2f"),
                          "N. Y":   st.column_config.NumberColumn("N. Y",   format="%.3f")}
            editable = st.data_editor(view[[c for c in cols if c in view.columns]], column_config=column_cfg, disabled=[c for c in view.columns if c not in ["Länge","Breite"]], use_container_width=True, key=f"edit_{key}")
        else:
            st.info("Noch keine Nesting-Ergebnisse in dieser Etappe.")
    if st.session_state.get(key) is not None and not st.session_state.get(key).empty:
        dyn_frames.append(st.session_state.get(key))

if dyn_frames:
    st.session_state.et_dyn = pd.concat(dyn_frames, ignore_index=True)
else:
    st.session_state.et_dyn = pd.DataFrame()

st.markdown("---")

# ========= Cluster-Analyse =========
st.subheader("🔎 Cluster-Analyse: gleiche Längen (Excel-Originaldaten)")
base = st.session_state.master_df.copy() if st.session_state.master_df is not None else pd.DataFrame()
if base.empty:
    st.info("Keine Daten.")
else:
    base["Länge_f"] = pd.to_numeric(base.get("Länge_f", pd.NA), errors="coerce")
    base["Länge_r"] = base["Länge_f"].round(3)
    cl = base.groupby("Länge_r", as_index=False).agg(Anzahl=("UID","count"))
    cl = cl.sort_values(["Anzahl","Länge_r"], ascending=[False, True]).reset_index(drop=True)
    st.dataframe(cl, use_container_width=True)
    if len(cl):
        sel = st.number_input("Detail-Cluster anzeigen (Länge_r)", value=float(cl["Länge_r"].iloc[0]), step=0.001, format="%.3f")
        detail = base[base["Länge_r"]==sel][["PB","Pak","Bezeichnung","Länge"]]
        if not detail.empty:
            st.caption(f"Teile mit Länge_r={sel:.3f}")
            st.dataframe(detail, use_container_width=True)

st.markdown("---")

# ========= Gesamtübersicht + Exporte =========
st.header("📊 Gesamtübersicht")
results=[]
for key,et in (("et1_result",1),("et2_result",2)):
    dfk=st.session_state.get(key)
    if dfk is not None and not dfk.empty:
        t=dfk.copy(); t["Etappe"]=et; results.append(t)
df_dyn = st.session_state.get("et_dyn")
if df_dyn is not None and not df_dyn.empty:
    t = df_dyn.copy()
    if "Etappe" not in t.columns: t["Etappe"] = 3
    results.append(t)
# Etappe 0 synthetisch
def build_et0_rows_for_total():
    if not st.session_state.rest_assignments: return pd.DataFrame()
    m = st.session_state.master_df.copy()
    for c in ("Höhe_f","Breite_f","Länge_f"): m[c]=pd.to_numeric(m[c], errors="coerce")
    rows=[]
    for rec in st.session_state.rest_assignments:
        uid=rec["UID"]; part=m[m["UID"]==uid]
        if part.empty: continue
        r=part.iloc[0]; z=float(r["Höhe_f"]); y=float(r["Breite_f"]); x=float(r["Länge_f"])
        rows.append({
            "Etappe":0,"Nesting-Nr":rec["Rest_Nr"],"PB":r.get("PB",""),"PB.":r.get("PB.",""),
            "GES":r.get("GES",""),"Pak":r.get("Pak",""),"Bauteil-Name":r.get("Bezeichnung",""),
            "Dimension Z":z,"Dimension Y":round(y,3),"Dimension X":round(x,3),
            "m2":round(x*y,3),"m3":round((x*y)*z/1000,3),"Breite":y,"Länge":x,"Unvollständig":"🟣","UID":uid
        })
    return pd.DataFrame(rows)
et0_rows = build_et0_rows_for_total()
if et0_rows is not None and not et0_rows.empty: results.insert(0, et0_rows)

if results:
    gesamt=pd.concat(results,ignore_index=True).drop_duplicates(subset=["UID"], keep="first")
    gesamt_view=gesamt.rename(columns={"Dimension Z":"N. Z","Dimension Y":"N. Y","Dimension X":"N. X","m2":"N.m2","m3":"N.m3"})

    for c in ["Länge","Breite","N. Z","N. Y","N. X","UID","Etappe","Nesting-Nr","PB.","PB","GES","Bauteil-Name","Pak","Unvollständig"]:
        if c not in gesamt_view.columns: gesamt_view[c]=np.nan

    assign_df = pd.DataFrame(st.session_state.get("rest_assignments", []))
    rest_map = dict(assign_df.groupby("UID")["Rest_Nr"].last()) if not assign_df.empty else {}
    for newc in ["Quelle","🟣","Rest_Nr"]:
        if newc not in gesamt_view.columns: gesamt_view[newc]=pd.NA
    if "UID" in gesamt_view.columns:
        m_bestand = gesamt_view["UID"].map(rest_map).notna() | (gesamt_view.get("Etappe", 9) == 0)
        gesamt_view.loc[m_bestand, "Quelle"] = "Bestand"
        gesamt_view.loc[m_bestand, "🟣"] = "🟣"
        gesamt_view["Rest_Nr"] = gesamt_view["UID"].map(rest_map)

    sortable = [c for c in ["Etappe","Nesting-Nr","PB","Pak","GES","Bauteil-Name","N. Z","N. Y","N. X","Breite","Länge"] if c in gesamt_view.columns]
    sel = st.multiselect("Sortierreihenfolge (Gesamtübersicht)", sortable, default=[])
    dirs = {s: st.selectbox(f"Richtung {s}", ["Absteigend","Aufsteigend"], key=f"dir_gesamt_{s}") for s in sel}
    if sel:
        by, asc = [], []
        for c in sel:
            mapped = 'Pak_num' if c=='Pak' and 'Pak_num' in gesamt_view.columns else ('Dimension X' if c=='N. X' and 'Dimension X' in gesamt_view.columns else c)
            by.append(mapped)
            asc.append(dirs[c]=="Aufsteigend")
        gesamt_view = gesamt_view.sort_values(by=by, ascending=asc, kind="mergesort").reset_index(drop=True)

    base_cols=["Etappe","Nesting-Nr","PB","GES","Pak","Bauteil-Name","N. Z","N. Y","N. X","Breite","Länge"]
    tail_cols=["Unvollständig","Quelle","🟣","Rest_Nr","UID"]
    display_df = gesamt_view[[c for c in base_cols if c in gesamt_view.columns] + [c for c in tail_cols if c in gesamt_view.columns]].copy()

    editable_cols=["PB","GES","Bauteil-Name","Pak"]
    cfg = {"N. Y": st.column_config.NumberColumn("N. Y", format="%.3f"),
           "Breite": st.column_config.NumberColumn("Breite", format="%.3f"),
           "Länge": st.column_config.NumberColumn("Länge", format="%.2f")}
    st.data_editor(display_df, use_container_width=True, disabled=[c for c in display_df.columns if c not in editable_cols], column_config=cfg, key="gesamt_editor")

    # ---- Export Gesamtergebnis mit 3 Nachkommastellen (Excel-Format) ----
    ts=datetime.now().strftime("%Y-%m-%d_%H-%M")
    export_cols = ["Etappe","Nesting-Nr","PB","N. Z","N. Y","N. X","N.m2","N.m3","GES",
                   "Bauteil-Name","Pak","PB.","Höhe","Breite","Länge","B.m2","B.m3",
                   "Unvollständig","Quelle","🟣","Rest_Nr"]
    # 'Höhe' sicherstellen (aus N. Z übernehmen, falls nicht vorhanden)
    if 'Höhe' not in gesamt_view.columns:
        gesamt_view['Höhe'] = gesamt_view.get('N. Z')

    # B.m2/B.m3 neu berechnen:
    gesamt_view["B.m2"]=round(pd.to_numeric(gesamt_view["Länge"],errors="coerce")*pd.to_numeric(gesamt_view["Breite"],errors="coerce"),3)
    gesamt_view["B.m3"]=round(gesamt_view["B.m2"]*pd.to_numeric(gesamt_view["N. Z"],errors="coerce")/1000,3)
    gesamt_export = gesamt_view[[c for c in export_cols if c in gesamt_view.columns]].copy()

    for c in ["N. Y","Breite","N.m2","N.m3","B.m2","B.m3"]:
        if c in gesamt_export.columns:
            gesamt_export[c] = pd.to_numeric(gesamt_export[c], errors="coerce").round(3)

    out1=BytesIO()
    with pd.ExcelWriter(out1, engine="openpyxl") as w:
        sheet_name = "Ergebnis"
        gesamt_export.to_excel(w, index=False, sheet_name=sheet_name)
        ws = w.sheets[sheet_name]
        header = {ws.cell(row=1, column=j).value: j for j in range(1, ws.max_column+1)}
        three_dec_cols = [c for c in ["N. Y","Breite","N.m2","N.m3","B.m2","B.m3"] if c in header]
        for col_name in three_dec_cols:
            j = header[col_name]
            for i in range(2, ws.max_row+1):
                ws.cell(row=i, column=j).number_format = "0.000"

    st.download_button("📤 Export Gesamtergebnis",
                       data=out1.getvalue(),
                       file_name=f"nesting_gesamt_{ts}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Export: Reste-Verwertung (wie Gesamtergebnis)
    df_rest_like = gesamt_view[gesamt_view["Rest_Nr"].notna()].copy()
    out2=BytesIO()
    with pd.ExcelWriter(out2, engine="openpyxl") as w:
        rest_cols=[c for c in base_cols if c in df_rest_like.columns] + [c for c in ["Unvollständig","Quelle","🟣","Rest_Nr"] if c in df_rest_like.columns]
        df_rest_like[rest_cols].to_excel(w, index=False, sheet_name="Reste (wie Gesamtergebnis)")
    st.download_button("📤 Export Reste-Verwertung (wie Gesamtergebnis)",
                       data=out2.getvalue(),
                       file_name=f"reste_wie_gesamt_{ts}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ========= Statistik (Gesamt) =========
    st.subheader("🧮 Statistik (Gesamt)")
    total_saw_m2 = total_saw_m3 = total_teile_m2 = total_teile_m3 = 0.0
    nxny_sum = 0.0
    keys_all = ["et1_result","et2_result"] + [d["key"] for d in st.session_state.dyn_etappen]
    for _key in keys_all:
        dfk = st.session_state.get(_key)
        if dfk is not None and not getattr(dfk, "empty", True):
            s = compute_verschnitt_stats_split(dfk, float(max_laenge))
            total_saw_m2 += s["saw_m2"]; total_saw_m3 += s["saw_m3"]
            total_teile_m2 += s["teile_m2"]; total_teile_m3 += s["teile_m3"]
            g = s["gruppen"]
            if g is not None and not g.empty:
                nxny_sum += float((pd.to_numeric(g["nx"], errors="coerce") * pd.to_numeric(g["ny"], errors="coerce")).sum())
    pct_saw_overall = (100.0 * total_saw_m2 / nxny_sum) if nxny_sum > 0 else 0.0
    st.write(f"Sägeverlust: **{round(total_saw_m2,3)} m²** | **{round(total_saw_m3,3)} m³** | **{pct_saw_overall:.2f} %**")
    st.write(f"Bauteile: **{round(total_teile_m2,3)} m²** | **{round(total_teile_m3,3)} m³**")

else:
    st.info("Noch keine Nestings vorhanden.")
