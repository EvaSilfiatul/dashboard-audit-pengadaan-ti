"""
STUDI KASUS 2 - AUDIT PROSES PENGADAAN TI
Pengolahan Data Pengadaan dengan Python + Mapping ke COBIT 5 domain BAI03

Cara pakai:
    pip install pandas numpy matplotlib openpyxl
    python audit_pengadaan_bai03.py

Jika punya data asli, taruh sebagai 'data_pengadaan.csv' dengan kolom yang sama
(lihat fungsi buat_data_contoh) lalu jalankan ulang. Kalau file tidak ada,
skrip membuat data contoh otomatis.
"""
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

FOLDER = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(FOLDER, "data_pengadaan.csv")

BATAS_PEMBELIAN_LANGSUNG = 50_000_000   # batas nilai pembelian langsung (kebijakan contoh)
TOLERANSI_HARGA = 0.10                  # kontrak > estimasi +10% dianggap janggal
MAKS_HARI_PROSES = 30                   # SLA PR -> PO (hari)
BATAS_VENDOR_DOMINAN = 0.30             # satu vendor > 30% total nilai = konsentrasi tinggi


# ---------------------------------------------------------------------------
# 1. DATA CONTOH (hanya dipakai bila data_pengadaan.csv belum ada)
# ---------------------------------------------------------------------------
def buat_data_contoh(n=120, seed=7):
    rng = np.random.default_rng(seed)
    vendor = ["PT Jaya Teknologi", "CV Mitra Digital", "PT Nusa Network",
              "PT Data Prima", "CV Karya Komputer", "PT Sinar Solusi"]
    kategori = {"Laptop": 15e6, "Server": 120e6, "Switch/Router": 40e6,
                "Lisensi Software": 60e6, "CCTV": 30e6, "Printer": 8e6}
    pemohon = ["Andi", "Budi", "Citra", "Dewi", "Eko"]
    penyetuju = ["Kepala IT", "Manajer Umum", "Direktur"]

    baris = []
    for i in range(1, n + 1):
        kat = rng.choice(list(kategori))
        estimasi = kategori[kat] * rng.uniform(0.8, 1.3) * rng.integers(1, 4)
        kontrak = estimasi * rng.uniform(0.9, 1.05)
        tgl_pr = pd.Timestamp("2026-01-05") + pd.Timedelta(days=int(rng.integers(0, 280)))
        tgl_po = tgl_pr + pd.Timedelta(days=int(rng.integers(3, 25)))
        tgl_terima = tgl_po + pd.Timedelta(days=int(rng.integers(7, 30)))
        pm = rng.choice(pemohon)
        baris.append({
            "id_pengadaan": f"PO-{i:03d}",
            "tgl_pr": tgl_pr, "tgl_po": tgl_po, "tgl_terima": tgl_terima,
            "pemohon": pm,
            "penyetuju": rng.choice(penyetuju),
            "vendor": rng.choice(vendor, p=[.40, .15, .15, .12, .10, .08]),
            "kategori": kat,
            "nilai_estimasi": round(estimasi),
            "nilai_kontrak": round(kontrak),
            "metode": rng.choice(["Tender", "Penunjukan Langsung", "Pembelian Langsung"],
                                 p=[.45, .25, .30]),
            "jumlah_penawaran": int(rng.integers(2, 5)),
            "ada_kontrak": "Ya", "ada_bast": "Ya", "status_persetujuan": "Disetujui",
        })
    df = pd.DataFrame(baris)

    # --- sisipkan kejanggalan agar ada temuan ---
    df.loc[3, ["status_persetujuan", "penyetuju"]] = ["Belum Disetujui", np.nan]
    df.loc[17, "status_persetujuan"] = "Belum Disetujui"
    df.loc[25, "penyetuju"] = df.loc[25, "pemohon"]             # pemohon = penyetuju
    df.loc[60, "penyetuju"] = df.loc[60, "pemohon"]
    df.loc[10, "nilai_kontrak"] = round(df.loc[10, "nilai_estimasi"] * 1.45)   # harga terlalu tinggi
    df.loc[44, "nilai_kontrak"] = round(df.loc[44, "nilai_estimasi"] * 1.30)
    for idx in (8, 31, 52, 77):
        df.loc[idx, ["metode", "jumlah_penawaran"]] = ["Penunjukan Langsung", 1]
    df.loc[[5, 90], "ada_kontrak"] = "Tidak"
    df.loc[[12, 66, 101], "ada_bast"] = "Tidak"
    df.loc[70, "tgl_po"] = df.loc[70, "tgl_pr"] + pd.Timedelta(days=55)  # proses lambat
    df.loc[99, "tgl_po"] = df.loc[99, "tgl_pr"] + pd.Timedelta(days=48)

    # pemecahan pembelian (split purchase) tepat di bawah batas 50 juta
    tgl = pd.Timestamp("2026-06-10")
    for k in range(3):
        df.loc[len(df)] = {
            "id_pengadaan": f"PO-{121 + k}", "tgl_pr": tgl, "tgl_po": tgl + pd.Timedelta(days=k + 2),
            "tgl_terima": tgl + pd.Timedelta(days=15), "pemohon": "Budi",
            "penyetuju": "Kepala IT", "vendor": "CV Mitra Digital", "kategori": "Server",
            "nilai_estimasi": 48_000_000, "nilai_kontrak": 48_000_000 - k * 500_000,
            "metode": "Pembelian Langsung", "jumlah_penawaran": 1,
            "ada_kontrak": "Ya", "ada_bast": "Ya", "status_persetujuan": "Disetujui"}
    # data ganda
    df.loc[len(df)] = df.loc[20]
    return df


# ---------------------------------------------------------------------------
# 2. PEMBERSIHAN DATA
# ---------------------------------------------------------------------------
def bersihkan(df):
    log = {}
    log["jumlah_baris_awal"] = len(df)
    for c in ("tgl_pr", "tgl_po", "tgl_terima"):
        df[c] = pd.to_datetime(df[c])
    ganda = df[df.duplicated(subset=["tgl_pr", "vendor", "kategori", "nilai_kontrak", "pemohon"],
                             keep="first")]
    log["data_ganda"] = ganda
    df = df.drop(ganda.index).reset_index(drop=True)
    log["jumlah_baris_bersih"] = len(df)
    log["nilai_kosong"] = df.isna().sum()[lambda s: s > 0]
    df["hari_proses"] = (df["tgl_po"] - df["tgl_pr"]).dt.days
    df["selisih_harga_pct"] = (df["nilai_kontrak"] - df["nilai_estimasi"]) / df["nilai_estimasi"]
    return df, log


# ---------------------------------------------------------------------------
# 3. PENGUJIAN AUDIT  (setiap uji dipetakan ke praktik BAI03)
# ---------------------------------------------------------------------------
def uji_audit(df):
    uji = {}

    # U1 - Persetujuan belum lengkap
    uji["U1"] = df[(df.status_persetujuan != "Disetujui") | df.penyetuju.isna()]

    # U2 - Pemisahan tugas: pemohon = penyetuju
    uji["U2"] = df[df.pemohon == df.penyetuju]

    # U3 - Harga kontrak jauh di atas estimasi
    uji["U3"] = df[df.selisih_harga_pct > TOLERANSI_HARGA]

    # U4 - Kurang persaingan: hanya 1 penawaran pada nilai di atas batas pembelian langsung
    #      (di bawah batas, 1 penawaran masih wajar)
    uji["U4"] = df[(df.jumlah_penawaran < 2) & (df.nilai_kontrak >= BATAS_PEMBELIAN_LANGSUNG)]

    # U5 - Pemecahan pembelian (split): vendor & pemohon sama, <=14 hari, tiap nilai di bawah batas,
    #      namun total melewati batas
    split = []
    kecil = df[(df.nilai_kontrak < BATAS_PEMBELIAN_LANGSUNG) & (df.metode == "Pembelian Langsung")]
    for (v, p), g in kecil.groupby(["vendor", "pemohon"]):
        g = g.sort_values("tgl_pr")
        for i in range(len(g)):
            jendela = g[(g.tgl_pr >= g.iloc[i].tgl_pr) &
                        (g.tgl_pr <= g.iloc[i].tgl_pr + pd.Timedelta(days=14))]
            if len(jendela) >= 2 and jendela.nilai_kontrak.sum() > BATAS_PEMBELIAN_LANGSUNG:
                split.append(jendela)
    uji["U5"] = pd.concat(split).drop_duplicates("id_pengadaan") if split else df.iloc[0:0]

    # U6 - Dokumen pendukung tidak lengkap (kontrak / BAST)
    uji["U6"] = df[(df.ada_kontrak == "Tidak") | (df.ada_bast == "Tidak")]

    # U7 - Proses PR -> PO melebihi SLA
    uji["U7"] = df[df.hari_proses > MAKS_HARI_PROSES]

    # U8 - Konsentrasi vendor
    per_vendor = df.groupby("vendor").nilai_kontrak.sum().sort_values(ascending=False)
    porsi = (per_vendor / per_vendor.sum()).rename("porsi")
    uji["U8"] = porsi[porsi > BATAS_VENDOR_DOMINAN].to_frame()
    uji["_porsi_vendor"] = porsi
    return uji


# Pemetaan ke COBIT 5 BAI03 (Manage Solutions Identification and Build)
PETA = {
    "U1": ("Persetujuan pengadaan belum lengkap", "BAI03.04 Procure solution components",
           "Tinggi", "Pastikan setiap PO disetujui sebelum diproses; blokir PO tanpa persetujuan di sistem."),
    "U2": ("Pemohon sekaligus penyetuju (SoD)", "BAI03.04 Procure solution components",
           "Tinggi", "Pisahkan peran pemohon, penyetuju, dan penerima barang."),
    "U3": ("Harga kontrak melebihi estimasi >10%", "BAI03.04 Procure solution components",
           "Tinggi", "Wajibkan justifikasi tertulis dan persetujuan ulang bila harga melewati toleransi."),
    "U4": ("Kurang persaingan (1 penawaran pada nilai besar)", "BAI03.04 Procure solution components",
           "Sedang", "Tetapkan minimal 3 penawaran untuk nilai tertentu; dokumentasikan alasan pengecualian."),
    "U5": ("Indikasi pemecahan pembelian (split purchase)", "BAI03.04 Procure solution components",
           "Tinggi", "Pantau pembelian berulang ke vendor sama dalam periode pendek; gabungkan sebagai satu paket."),
    "U6": ("Dokumen kontrak/BAST tidak lengkap", "BAI03.06 Perform quality assurance",
           "Sedang", "Pembayaran hanya setelah kontrak dan berita acara serah terima (BAST) lengkap."),
    "U7": ("Proses PR ke PO melebihi batas waktu", "BAI03.04 Procure solution components",
           "Rendah", "Tetapkan SLA proses pengadaan dan pantau berkala."),
    "U8": ("Ketergantungan pada satu vendor", "BAI03.04 Procure solution components",
           "Sedang", "Diversifikasi vendor dan evaluasi kinerja vendor secara berkala."),
}


def ringkasan_temuan(uji):
    baris = []
    for kode, (judul, praktik, risiko, saran) in PETA.items():
        hasil = uji[kode]
        baris.append({"Kode": kode, "Temuan": judul, "Jumlah Kasus": len(hasil),
                      "Praktik BAI03": praktik, "Tingkat Risiko": risiko, "Rekomendasi": saran})
    return pd.DataFrame(baris)


# ---------------------------------------------------------------------------
# 4. GRAFIK
# ---------------------------------------------------------------------------
def buat_grafik(df, uji, ringkas):
    # Grafik 1: jumlah temuan per uji
    plt.figure(figsize=(9, 4.5))
    r = ringkas.sort_values("Jumlah Kasus")
    warna = r["Tingkat Risiko"].map({"Tinggi": "#d9534f", "Sedang": "#f0ad4e", "Rendah": "#5cb85c"})
    plt.barh(r["Kode"] + " " + r["Temuan"].str.slice(0, 38), r["Jumlah Kasus"], color=warna)
    plt.xlabel("Jumlah kasus")
    plt.title("Temuan Audit Pengadaan TI (merah = risiko tinggi)")
    plt.tight_layout()
    plt.savefig(os.path.join(FOLDER, "grafik_temuan.png"), dpi=150)
    plt.close()

    # Grafik 2: porsi nilai per vendor
    plt.figure(figsize=(7, 4))
    (uji["_porsi_vendor"] * 100).plot(kind="bar", color="#4a7fb5")
    plt.axhline(BATAS_VENDOR_DOMINAN * 100, color="red", ls="--", label="Batas 30%")
    plt.ylabel("% dari total nilai kontrak")
    plt.title("Konsentrasi Vendor")
    plt.legend()
    plt.xticks(rotation=30, ha="right")
    plt.tight_layout()
    plt.savefig(os.path.join(FOLDER, "grafik_vendor.png"), dpi=150)
    plt.close()

    # Grafik 3: sebaran lama proses PR->PO
    plt.figure(figsize=(7, 4))
    plt.hist(df["hari_proses"], bins=15, color="#7a9e7e")
    plt.axvline(MAKS_HARI_PROSES, color="red", ls="--", label=f"SLA {MAKS_HARI_PROSES} hari")
    plt.xlabel("Hari dari PR ke PO")
    plt.ylabel("Jumlah pengadaan")
    plt.title("Lama Proses Pengadaan")
    plt.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(FOLDER, "grafik_waktu_proses.png"), dpi=150)
    plt.close()


# ---------------------------------------------------------------------------
# 5. MAIN
# ---------------------------------------------------------------------------
def main():
    if not os.path.exists(CSV):
        buat_data_contoh().to_csv(CSV, index=False)
        print("Data contoh dibuat:", CSV)

    df_raw = pd.read_csv(CSV)
    df, log = bersihkan(df_raw)
    uji = uji_audit(df)
    ringkas = ringkasan_temuan(uji)
    buat_grafik(df, uji, ringkas)

    print("=" * 60)
    print("PEMBERSIHAN DATA")
    print(f"Baris awal : {log['jumlah_baris_awal']}")
    print(f"Data ganda : {len(log['data_ganda'])} (dihapus)")
    print(f"Baris bersih: {log['jumlah_baris_bersih']}")
    print(f"Total nilai kontrak: Rp {df.nilai_kontrak.sum():,.0f}")
    print("=" * 60)
    print("RINGKASAN TEMUAN & PEMETAAN BAI03")
    print(ringkas[["Kode", "Temuan", "Jumlah Kasus", "Praktik BAI03", "Tingkat Risiko"]]
          .to_string(index=False))

    with pd.ExcelWriter(os.path.join(FOLDER, "hasil_audit_pengadaan.xlsx")) as xw:
        ringkas.to_excel(xw, sheet_name="Ringkasan BAI03", index=False)
        df.to_excel(xw, sheet_name="Data Bersih", index=False)
        for k in PETA:
            hasil = uji[k]
            (hasil if k != "U8" else hasil.reset_index()).to_excel(xw, sheet_name=k, index=False)
        log["data_ganda"].to_excel(xw, sheet_name="Data Ganda", index=False)
    print("\nFile hasil: hasil_audit_pengadaan.xlsx, grafik_*.png")


if __name__ == "__main__":
    main()
