# app_ve10.py
# -*- coding: utf-8 -*-
# ------------------------------------------------------------
# VE.10 — Load Profile Analyzer (PEA)
# Streamlit Web UI for analyzing electronic meter load profiles.
# Author: PEA Change Management Office
# ------------------------------------------------------------

import io
import json
import math
from datetime import datetime, time

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(
    page_title="VE.10 | Load Profile Analyzer (PEA)",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------
# Helpers
# -----------------------------
DEFAULT_COLMAP = {
    "datetime": "CH1",
    "status": "CH2",
    "voltage": "CH3",
    "p_import": "CH4",
    "p_export": "CH5",
}

def load_dataframe(uploaded, csv_sep=",", excel_sheet=None, excel_skiprows=None):
    if uploaded is None:
        return None

    import os
    name = uploaded.name.lower()
    ext = os.path.splitext(name)[-1]

    # Excel
    if ext in [".xls", ".xlsx"]:
        kw = {}
        if excel_sheet not in (None, "", "auto"):
            kw["sheet_name"] = excel_sheet
        if isinstance(excel_skiprows, int) and excel_skiprows > 0:
            kw["skiprows"] = excel_skiprows

        if ext == ".xls":
            # Excel รุ่นเก่า ต้องใช้ xlrd
            df = pd.read_excel(uploaded, engine="xlrd", **kw)
        else:
            # Excel รุ่นใหม่ (xlsx)
            df = pd.read_excel(uploaded, engine="openpyxl", **kw)

    # CSV
    elif ext == ".csv":
        df = pd.read_csv(uploaded, sep=csv_sep)

    else:
        raise ValueError(f"ไม่รองรับไฟล์นามสกุล {ext}")

    return df

def clean_unit_series(s):
    """Remove ' V' / ' kW' suffixes if present; coerce to numeric."""
    if s is None:
        return s
    if pd.api.types.is_numeric_dtype(s):
        return s
    s = (
        s.astype(str)
         .str.replace(" V", "", regex=False)
         .str.replace(" kW", "", regex=False)
         .str.replace(",", "", regex=False)
    )
    return pd.to_numeric(s, errors="coerce")

def parse_datetime_col(s, fmt=None):
    if fmt and fmt.strip():
        return pd.to_datetime(s, format=fmt, errors="coerce")
    return pd.to_datetime(s, errors="coerce")

def summarize_missing(df):
    miss = df.isna().sum().to_frame("Missing")
    miss["Percent"] = (miss["Missing"] / len(df)) * 100 if len(df) else 0
    return miss

def flag_anomalies(df, v_col, p_col, v_sag, v_swell, base_kw_quiet=0.05, z_thresh=3.0):
    out = pd.DataFrame(index=df.index)
    out["sag"] = df[v_col] < v_sag
    out["swell"] = df[v_col] > v_swell
    # constant low import (potential leakage at night 00:00-04:00) for > N points
    # mark each row that meets nighttime & base import
    idx_night = df.index.indexer_between_time(time(0,0), time(4,0))
    night_mask = pd.Series(False, index=df.index)
    night_mask.iloc[idx_night] = True if len(idx_night) else False
    out["night_base"] = (df[p_col] >= base_kw_quiet) & night_mask

    # Z-score outliers on kW import
    x = df[p_col].astype(float)
    if len(x.dropna()) > 5 and x.std(skipna=True) > 0:
        z = (x - x.mean(skipna=True)) / x.std(skipna=True)
        out["kw_outlier"] = z.abs() > z_thresh
    else:
        out["kw_outlier"] = False

    return out

def compute_kwh(df, p_col, interval_minutes):
    return df[p_col].fillna(0) * (interval_minutes / 60.0)

def to_excel_bytes(df_dict):
    """df_dict = {"SheetName": dataframe} -> bytes of xlsx"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, d in df_dict.items():
            d.to_excel(writer, sheet_name=str(name)[:31], index=True)
    return output.getvalue()

# -----------------------------
# Sidebar — Controls
# -----------------------------
st.sidebar.title("VE.10 — Settings")

with st.sidebar.expander("ข้อมูลนำเข้า (Input)", expanded=True):
    uploaded = st.file_uploader("อัปโหลดไฟล์ Excel/CSV", type=["xls","xlsx","csv"])
    filetype_hint = st.caption("รองรับ .xls, .xlsx, .csv")
    csv_sep = st.text_input("CSV Separator", value=",")
    excel_sheet = st.text_input("Excel Sheet name (เว้นว่าง=auto)", value="")
    excel_skiprows = st.number_input("Excel skiprows (เผื่อหัวตาราง)", min_value=0, value=7, step=1)

with st.sidebar.expander("คอลัมน์และรูปแบบเวลา", expanded=True):
    st.caption("ตั้งค่าคอลัมน์ให้ตรงกับไฟล์จริงของ load meter")
    col_datetime = st.text_input("คอลัมน์เวลา", value=DEFAULT_COLMAP["datetime"])
    col_status = st.text_input("คอลัมน์สถานะ", value=DEFAULT_COLMAP["status"])
    col_voltage = st.text_input("คอลัมน์แรงดัน (V)", value=DEFAULT_COLMAP["voltage"])
    col_pimp = st.text_input("คอลัมน์กำลังใช้นำเข้า (kW)", value=DEFAULT_COLMAP["p_import"])
    col_pexp = st.text_input("คอลัมน์กำลังจ่ายย้อน (kW)", value=DEFAULT_COLMAP["p_export"])
    date_fmt = st.text_input("รูปแบบวันที่เวลา (เช่น %d-%m-%Y %H:%M:%S)", value="%d-%m-%Y %H:%M:%S")

with st.sidebar.expander("เกณฑ์/พารามิเตอร์วิเคราะห์", expanded=True):
    interval_min = st.number_input("ช่วงเวลา 1 จุดข้อมูล (นาที)", min_value=1, value=15, step=1)
    v_sag = st.number_input("เกณฑ์ไฟตก (V ต่ำกว่า)", min_value=0, value=210, step=1)
    v_swell = st.number_input("เกณฑ์ไฟเกิน (V สูงกว่า)", min_value=240, step=1)
    base_kw_quiet = st.number_input("เกณฑ์ base load ตอนกลางคืน (kW)", min_value=0.0, value=0.05, step=0.01, format="%.2f")
    z_thresh = st.number_input("Z-score threshold (kW outlier)", min_value=1.0, value=3.0, step=0.5)
    resample_rule = st.selectbox("สรุปช่วงเวลา (Resample)", ["15T","30T","H","D","M"], index=2,
        help="15T=15 นาที, H=ชั่วโมง, D=วัน, M=เดือน")

with st.sidebar.expander("ตัวกรอง", expanded=True):
    filter_positive_power = st.checkbox("แสดงเฉพาะช่วงที่ Import > 0", value=False)

# -----------------------------
# Header
# -----------------------------
st.title("เครื่องมือวิเคราะห์ Load Profile มิเตอร์อิเล็กทรอนิกส์ (VE.10)")
st.markdown(
    """
**เป้าหมาย:** ยกระดับขีดความสามารถทีมตรวจสอบหน่วย/มิเตอร์, ใช้ข้อมูลมิเตอร์ตรวจสอบความผิดปกติ/ตอบร้องเรียน, ลดเวลาวิเคราะห์ และเพิ่มความน่าเชื่อถือ  
**หน่วยงาน:** กองบริหารโครงการและการจัดการการเปลี่ยนแปลง (ผู้สนับสนุน: กมต)
"""
)

# -----------------------------
# Load & Prepare
# -----------------------------
tab_overview, tab_quality, tab_timeseries, tab_anomaly, tab_aggregate, tab_export, tab_about = st.tabs(
    ["📥 ภาพรวม/นำเข้า", "🔎 คุณภาพข้อมูล", "📈 กราฟเวลา", "🚨 ผิดปกติ", "🧮 สรุปตามช่วงเวลา", "📤 ส่งออก", "ℹ️ เกี่ยวกับ"]
)

with tab_overview:
    st.subheader("อัปโหลดและกำหนดค่า")
    if uploaded is None:
        st.info("โปรดอัปโหลดไฟล์ข้อมูลเพื่อเริ่มต้นใช้งาน")
    else:
        try:
            df_raw = load_dataframe(uploaded, csv_sep=csv_sep, excel_sheet=excel_sheet, excel_skiprows=excel_skiprows)
        except Exception as e:
            st.error(f"ไม่สามารถอ่านไฟล์ได้: {e}")
            st.stop()

        st.success(f"โหลดข้อมูลสำเร็จ: {df_raw.shape[0]:,} แถว, {df_raw.shape[1]} คอลัมน์")
        with st.expander("ตัวอย่างข้อมูลดิบ"):
            st.dataframe(df_raw.head(20), use_container_width=True)

        # --- Select only needed columns (if present) ---
        required_cols = [col_datetime, col_status, col_voltage, col_pimp, col_pexp]
        missing = [c for c in required_cols if c not in df_raw.columns]
        if missing:
            st.error(f"ไม่พบคอลัมน์ที่ต้องใช้: {missing}\nโปรดปรับชื่อคอลัมน์ให้ตรงกับไฟล์จริง")
            st.stop()

        df = df_raw[required_cols].copy()

        # Parse datetime
        df[col_datetime] = parse_datetime_col(df[col_datetime], fmt=date_fmt)

        # Clean units -> numeric
        for c in [col_voltage, col_pimp, col_pexp]:
            df[c] = clean_unit_series(df[c])

        # Drop NA datetime & sort
        df = df.dropna(subset=[col_datetime]).sort_values(col_datetime)
        df = df.set_index(col_datetime)

        if filter_positive_power:
            df = df[df[col_pimp] > 0]

        # Store in session for other tabs
        st.session_state["df"] = df
        st.session_state["col_names"] = {
            "status": col_status, "voltage": col_voltage, "p_import": col_pimp, "p_export": col_pexp
        }
        st.session_state["params"] = {
            "interval_min": interval_min, "v_sag": v_sag, "v_swell": v_swell,
            "base_kw_quiet": base_kw_quiet, "z_thresh": z_thresh, "resample_rule": resample_rule
        }

        # Range filter UI
        st.markdown("### เลือกช่วงเวลาใช้งาน")
        min_dt, max_dt = df.index.min(), df.index.max()
        if pd.isna(min_dt) or pd.isna(max_dt):
            st.error("ไม่พบค่าเวลาในข้อมูลหลังแปลง โปรดตรวจรูปแบบวันที่/เวลา")
            st.stop()

        start, end = st.slider(
            "ช่วงเวลา",
            min_value=min_dt.to_pydatetime(),
            max_value=max_dt.to_pydatetime(),
            value=(min_dt.to_pydatetime(), max_dt.to_pydatetime()),
            step=None,
            format="YYYY-MM-DD HH:mm"
        )
        df = df.loc[start:end]
        st.session_state["df"] = df  # update filtered

        st.metric("จำนวนแถวหลังกรองเวลา", len(df))

        with st.expander("ดูตัวอย่างข้อมูลหลัง Clean/Grooming"):
            st.dataframe(df.head(30), use_container_width=True)

with tab_quality:
    st.subheader("คุณภาพข้อมูล (Missing / Types)")
    df = st.session_state.get("df")
    if df is None:
        st.info("กรุณาอัปโหลดข้อมูลในแท็บแรก")
    else:
        miss = summarize_missing(df)
        st.dataframe(miss, use_container_width=True)
        st.caption(f"รวมแถว: {len(df):,}")

        # Quick describe
        st.markdown("#### สถิติเบื้องต้น")
        st.dataframe(df.describe(include="all").T, use_container_width=True)

        # Status "power" occurrences (e.g., power outage)
        col_status = st.session_state["col_names"]["status"]
        if col_status in df.columns:
            with st.expander("ค้นหาคำว่า 'power' ในคอลัมน์สถานะ"):
                power_events = df[df[col_status].astype(str).str.contains("power", case=False, na=False)]
                st.write(f"พบ {len(power_events):,} รายการที่มีคำว่า 'power'")
                st.dataframe(power_events.head(100), use_container_width=True)

with tab_timeseries:
    st.subheader("กราฟเวลา (Voltage / Power)")
    df = st.session_state.get("df")
    if df is None:
        st.info("กรุณาอัปโหลดข้อมูลในแท็บแรก")
    else:
        c1, c2 = st.columns(2)
        v_col = st.session_state["col_names"]["voltage"]
        p_col = st.session_state["col_names"]["p_import"]
        with c1:
            st.plotly_chart(px.line(df.reset_index(), x=df.index, y=v_col, title="แรงดันไฟฟ้า (V)"), use_container_width=True)
        with c2:
            st.plotly_chart(px.line(df.reset_index(), x=df.index, y=p_col, title="กำลังใช้นำเข้า (kW)"), use_container_width=True)

        with st.expander("การกระจายค่า (Distribution)"):
            c3, c4 = st.columns(2)
            with c3:
                st.plotly_chart(px.histogram(df, x=v_col, nbins=40, title="Voltage Histogram"), use_container_width=True)
            with c4:
                st.plotly_chart(px.box(df, y=p_col, points="suspectedoutliers", title="kW Boxplot"), use_container_width=True)

with tab_anomaly:
    st.subheader("การตรวจจับความผิดปกติ (Rules + Z-score)")
    df = st.session_state.get("df")
    if df is None:
        st.info("กรุณาอัปโหลดข้อมูลในแท็บแรก")
    else:
        names = st.session_state["col_names"]
        params = st.session_state["params"]
        v_col, p_col = names["voltage"], names["p_import"]

        flags = flag_anomalies(df, v_col, p_col, params["v_sag"], params["v_swell"],
                               base_kw_quiet=params["base_kw_quiet"], z_thresh=params["z_thresh"])
        st.session_state["flags"] = flags

        counts = flags.sum().to_frame("count")
        st.dataframe(counts.T, use_container_width=True)

        with st.expander("รายละเอียดแถวผิดปกติ (ตัวอย่าง 500 แถวแรก)"):
            show_cols = [v_col, p_col] + list(flags.columns)
            detail = df.join(flags)[show_cols]
            st.dataframe(detail[detail.any(axis=1)].head(500), use_container_width=True)

        # Overlay markers on power chart
        mark_df = df.join(flags)
        mark_df = mark_df.reset_index().rename(columns={mark_df.index.name: "ts"})
        fig = px.line(mark_df, x="ts", y=p_col, title="กำลังใช้นำเข้า (kW) + markers")
        if "sag" in mark_df:
            fig.add_scatter(x=mark_df.loc[mark_df["sag"], "ts"],
                            y=mark_df.loc[mark_df["sag"], p_col],
                            mode="markers", name="sag")
        if "swell" in mark_df:
            fig.add_scatter(x=mark_df.loc[mark_df["swell"], "ts"],
                            y=mark_df.loc[mark_df["swell"], p_col],
                            mode="markers", name="swell")
        if "kw_outlier" in mark_df:
            fig.add_scatter(x=mark_df.loc[mark_df["kw_outlier"], "ts"],
                            y=mark_df.loc[mark_df["kw_outlier"], p_col],
                            mode="markers", name="kW outlier")
        if "night_base" in mark_df:
            fig.add_scatter(x=mark_df.loc[mark_df["night_base"], "ts"],
                            y=mark_df.loc[mark_df["night_base"], p_col],
                            mode="markers", name="night base load")
        st.plotly_chart(fig, use_container_width=True)

with tab_aggregate:
    st.subheader("สรุปผลตามช่วงเวลา (Resample)")
    df = st.session_state.get("df")
    if df is None:
        st.info("กรุณาอัปโหลดข้อมูลในแท็บแรก")
    else:
        names = st.session_state["col_names"]
        params = st.session_state["params"]
        v_col, p_col = names["voltage"], names["p_import"]
        rule = params["resample_rule"]

        # compute kWh
        kWh = compute_kwh(df, p_col, params["interval_min"])
        df_tmp = df.copy()
        df_tmp["kWh"] = kWh

        agg_df = df_tmp.resample(rule).agg(
            kWh=("kWh", "sum"),
            avg_voltage=(v_col, "mean"),
            max_voltage=(v_col, "max"),
            min_voltage=(v_col, "min"),
            avg_kw=(p_col, "mean"),
            max_kw=(p_col, "max")
        )
        st.dataframe(agg_df, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(px.line(agg_df.reset_index(), x=agg_df.index, y="kWh", title=f"kWh ({rule})"), use_container_width=True)
        with c2:
            st.plotly_chart(px.line(agg_df.reset_index(), x=agg_df.index, y="avg_kw", title=f"เฉลี่ย kW ({rule})"), use_container_width=True)

        st.session_state["agg_df"] = agg_df
        st.session_state["df_kwh"] = df_tmp[["kWh"]]

with tab_export:
    st.subheader("ส่งออกผลลัพธ์")
    df = st.session_state.get("df")
    agg_df = st.session_state.get("agg_df")
    flags = st.session_state.get("flags")
    df_kwh = st.session_state.get("df_kwh")

    if df is None:
        st.info("ยังไม่มีข้อมูลให้ส่งออก")
    else:
        # CSV
        csv_clean = df.to_csv(index=True).encode("utf-8")
        st.download_button("⬇️ ดาวน์โหลดข้อมูลหลัง Clean (CSV)", data=csv_clean, file_name="clean_timeseries.csv", mime="text/csv")

        if agg_df is not None:
            csv_agg = agg_df.to_csv(index=True).encode("utf-8")
            st.download_button("⬇️ ดาวน์โหลดสรุปตามช่วงเวลา (CSV)", data=csv_agg, file_name="resampled_summary.csv", mime="text/csv")

        # Excel (multi-sheet)
        xls_bytes = to_excel_bytes({
            "Cleaned": df,
            "Resample": agg_df if agg_df is not None else pd.DataFrame(),
            "Flags": df.join(flags) if flags is not None else pd.DataFrame(),
            "kWh_series": df_kwh if df_kwh is not None else pd.DataFrame()
        })
        st.download_button("⬇️ ดาวน์โหลดไฟล์สรุป (Excel)", data=xls_bytes, file_name="ve10_summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with tab_about:
    st.subheader("เกี่ยวกับเครื่องมือ")
    st.markdown(
        """
- **รหัส/เวอร์ชัน:** VE.10 — Load Profile Analyzer (Streamlit UI)  
- **จุดเด่น:** อัปโหลด/ทำความสะอาด/ตั้งเกณฑ์/ตรวจจับผิดปกติ/สรุปตามช่วงเวลา/ดาวน์โหลดผล  
- **ต่อยอด:** 
  - เพิ่ม rule สำหรับกรณีเฉพาะ (ไฟย้อนจาก PV, มิเตอร์ขัดข้อง, cross-connection) โดยอิงกฎของ กฟภ.
  - ผูกฐานข้อมูล/ระบบสิทธิ์การใช้งาน
  - บันทึกค่า preset ต่อ CA/FO/หน่วยงาน
  - Export เป็นรายงาน PDF พร้อมกราฟ (เช่น รายวัน/รายเดือน)
"""
    )
    st.caption("© PEA – Change Management Office")
