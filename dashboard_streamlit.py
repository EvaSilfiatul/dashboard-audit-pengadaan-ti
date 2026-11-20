"""
BAB 8.6 - DASHBOARD AUDIT MENGGUNAKAN STREAMLIT
Studi Kasus 2: Audit Proses Pengadaan TI (COBIT 5 - BAI03)

Cara menjalankan (di komputer Anda):
    pip install streamlit pandas numpy matplotlib openpyxl reportlab
    streamlit run dashboard_streamlit.py

Taruh file ini satu folder dengan audit_pengadaan_bai03.py dan laporan_pdf.py.
Dashboard memakai data_pengadaan.csv; bila tidak ada, data contoh dibuat otomatis.
Anda juga bisa mengunggah CSV sendiri lewat sidebar.
"""
import io
import os

import pandas as pd
import streamlit as st

import audit_pengadaan_bai03 as audit
import laporan_pdf

FOLDER = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(FOLDER, "data_pengadaan.csv")

st.set_page_config(page_title="Dashboard Audit Pengadaan TI", page_icon="🔎", layout="wide")


# ---------------------------------------------------------------------------
# Memuat data
# ---------------------------------------------------------------------------
@st.cache_data
def muat_data(isi_unggahan):
    if isi_unggahan is not None:
        return pd.read_csv(io.BytesIO(isi_unggahan)), False
    if not os.path.exists(CSV):
        audit.buat_data_contoh().to_csv(CSV, index=False)
    return pd.read_csv(CSV), True


# ---------------------------------------------------------------------------
# Sidebar: data, parameter, filter
# ---------------------------------------------------------------------------
st.sidebar.title("Pengaturan Audit")
unggah = st.sidebar.file_uploader("Unggah data pengadaan (CSV)", type="csv")
data_mentah, data_contoh = muat_data(unggah.getvalue() if unggah else None)

st.sidebar.subheader("Parameter pengujian")
batas = st.sidebar.number_input("Batas pembelian langsung (Rp)", min_value=1_000_000,
                                value=int(audit.BATAS_PEMBELIAN_LANGSUNG), step=5_000_000)
tol = st.sidebar.slider("Toleransi harga di atas estimasi (%)", 0, 50, int(audit.TOLERANSI_HARGA * 100)) / 100
sla = st.sidebar.slider("Batas waktu proses PR ke PO (hari)", 5, 90, int(audit.MAKS_HARI_PROSES))
vend = st.sidebar.slider("Batas porsi satu vendor (%)", 10, 100, int(audit.BATAS_VENDOR_DOMINAN * 100)) / 100

# Parameter dibaca modul audit saat fungsi dipanggil
audit.BATAS_PEMBELIAN_LANGSUNG = batas
audit.TOLERANSI_HARGA = tol
audit.MAKS_HARI_PROSES = sla
audit.BATAS_VENDOR_DOMINAN = vend
parameter = {"batas": batas, "tol": tol, "sla": sla, "vend": vend}

df_all, log = audit.bersihkan(data_mentah.copy())

st.sidebar.subheader("Filter tampilan")
f_kat = st.sidebar.multiselect("Kategori", sorted(df_all.kategori.unique()))
f_vendor = st.sidebar.multiselect("Vendor", sorted(df_all.vendor.unique()))
f_metode = st.sidebar.multiselect("Metode pengadaan", sorted(df_all.metode.unique()))

df = df_all.copy()
if f_kat:
    df = df[df.kategori.isin(f_kat)]
if f_vendor:
    df = df[df.vendor.isin(f_vendor)]
if f_metode:
    df = df[df.metode.isin(f_metode)]

if df.empty:
    st.warning("Tidak ada data yang cocok dengan filter. Kosongkan sebagian filter.")
    st.stop()

# Pengujian dijalankan pada data yang sudah difilter
uji = audit.uji_audit(df)
ringkas = audit.ringkasan_temuan(uji)

# ---------------------------------------------------------------------------
# Tampilan utama
# ---------------------------------------------------------------------------
st.title("🔎 Dashboard Audit Proses Pengadaan TI")
st.caption("Studi Kasus 2 - COBIT 5 domain BAI03 (Manage Solutions Identification and Build)")
if data_contoh:
    st.info("Data yang tampil adalah data simulasi untuk pembelajaran. Unggah CSV Anda di sidebar untuk memakai data lain.")

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Transaksi", len(df))
k2.metric("Total nilai kontrak", f"Rp {df.nilai_kontrak.sum() / 1e9:,.2f} M")
k3.metric("Rata-rata proses PR-PO", f"{df.hari_proses.mean():.1f} hari")
k4.metric("Jenis temuan", int((ringkas["Jumlah Kasus"] > 0).sum()), f"dari {len(ringkas)} uji", delta_color="off")
k5.metric("Kasus risiko tinggi", int(ringkas.loc[ringkas["Tingkat Risiko"] == "Tinggi", "Jumlah Kasus"].sum()))

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["Ringkasan", "Detail temuan", "Vendor & biaya", "Waktu proses", "Laporan & unduhan"])

# --- Tab 1: ringkasan ---
with tab1:
    st.subheader("Temuan dan pemetaan BAI03")
    tampil = ringkas.copy()

    def warna(v):
        return {"Tinggi": "background-color:#f8d7da", "Sedang": "background-color:#fff3cd",
                "Rendah": "background-color:#d4edda"}.get(v, "")

    st.dataframe(tampil.style.map(warna, subset=["Tingkat Risiko"]), use_container_width=True, hide_index=True)
    st.subheader("Jumlah kasus per uji")
    st.bar_chart(ringkas.set_index("Kode")["Jumlah Kasus"])
    st.caption("U1-U3, U5 berisiko tinggi; U4, U6, U8 sedang; U7 rendah. "
               "Kasus 0 berarti tidak ada penyimpangan pada uji tersebut.")

# --- Tab 2: detail temuan ---
with tab2:
    st.subheader("Rincian transaksi per temuan")
    for _, r in ringkas.iterrows():
        k = r["Kode"]
        hasil = uji[k]
        judul = f"{k} - {r['Temuan']}  ({len(hasil)} kasus, risiko {r['Tingkat Risiko']})"
        with st.expander(judul, expanded=len(hasil) > 0 and r["Tingkat Risiko"] == "Tinggi"):
            st.markdown(f"**Praktik BAI03:** {r['Praktik BAI03']}")
            st.markdown(f"**Rekomendasi:** {r['Rekomendasi']}")
            if len(hasil):
                st.dataframe(hasil.reset_index() if k == "U8" else hasil, use_container_width=True, hide_index=True)
            else:
                st.success("Tidak ada penyimpangan.")

# --- Tab 3: vendor & biaya ---
with tab3:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Porsi nilai per vendor (%)")
        porsi = (uji["_porsi_vendor"] * 100).round(1)
        st.bar_chart(porsi)
        dominan = porsi[porsi > vend * 100]
        if len(dominan):
            st.warning("Melebihi batas: " + ", ".join(f"{v} ({p}%)" for v, p in dominan.items()))
    with c2:
        st.subheader("Nilai kontrak per kategori (Rp)")
        st.bar_chart(df.groupby("kategori").nilai_kontrak.sum())
    st.subheader("Selisih harga kontrak terhadap estimasi (%)")
    st.bar_chart(df.set_index("id_pengadaan")["selisih_harga_pct"] * 100)
    st.caption(f"Garis batas toleransi: {tol:.0%}. Batang di atas angka itu masuk temuan U3.")

# --- Tab 4: waktu proses ---
with tab4:
    st.subheader("Sebaran lama proses PR ke PO")
    sebaran = df["hari_proses"].value_counts().sort_index()
    st.bar_chart(sebaran)
    lambat = int((df.hari_proses > sla).sum())
    st.metric(f"Transaksi lebih dari {sla} hari", lambat, f"{lambat / len(df):.1%} dari total", delta_color="off")
    per_bulan = df.assign(bulan=df.tgl_pr.dt.to_period("M").astype(str)).groupby("bulan").hari_proses.mean()
    st.subheader("Rata-rata lama proses per bulan (hari)")
    st.line_chart(per_bulan)

# --- Tab 5: laporan & unduhan ---
with tab5:
    st.subheader("Laporan audit PDF")
    st.write("Laporan berisi latar belakang, hasil pengolahan data, temuan, rekomendasi, dan kesimpulan. "
             "Isi laporan mengikuti parameter dan filter yang sedang aktif.")
    c1, c2, c3 = st.columns(3)
    nama = c1.text_input("Nama penyusun", "[Nama Penyusun]")
    nim = c2.text_input("NIM", "[NIM]")
    instansi = c3.text_input("Program studi / universitas", "[Program Studi / Universitas]")

    if st.button("Buat laporan PDF", type="primary"):
        audit.buat_grafik(df, uji, ringkas)  # grafik mengikuti data terfilter
        buf = io.BytesIO()
        laporan_pdf.buat_laporan(df, uji, ringkas, log, buf, parameter=parameter, data_contoh=data_contoh,
                                 nama=nama, nim=nim, instansi=instansi)
        st.session_state["pdf"] = buf.getvalue()
        st.success("Laporan siap diunduh.")
    if "pdf" in st.session_state:
        st.download_button("Unduh laporan (PDF)", st.session_state["pdf"],
                           file_name="Laporan_Audit_Pengadaan_TI.pdf", mime="application/pdf")

    st.subheader("Unduh hasil analisis (Excel)")
    xbuf = io.BytesIO()
    with pd.ExcelWriter(xbuf, engine="openpyxl") as xw:
        ringkas.to_excel(xw, sheet_name="Ringkasan BAI03", index=False)
        df.to_excel(xw, sheet_name="Data Bersih", index=False)
        for k in audit.PETA:
            h = uji[k]
            (h.reset_index() if k == "U8" else h).to_excel(xw, sheet_name=k, index=False)
    st.download_button("Unduh Excel", xbuf.getvalue(), file_name="hasil_audit_pengadaan.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
