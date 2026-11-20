"""
BAB 8.7 - OTOMATISASI PEMBUATAN LAPORAN AUDIT PDF
Studi Kasus 2: Audit Proses Pengadaan TI (COBIT 5 - BAI03)

Cara pakai:
    pip install reportlab pandas matplotlib openpyxl
    python laporan_pdf.py          -> menghasilkan Laporan_Audit_Pengadaan_TI.pdf

Semua angka dan kalimat temuan dihitung otomatis dari data, jadi bila
data_pengadaan.csv diganti dengan data asli, laporan ikut berubah.
"""
import io
import os
from datetime import date

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

import audit_pengadaan_bai03 as audit

FOLDER = os.path.dirname(os.path.abspath(__file__))
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]

# ---------------------------------------------------------------------------
# Teks penjelasan tiap temuan: kriteria (patokan) dan risiko (dampak)
# ---------------------------------------------------------------------------
KRITERIA = {
    "U1": "Setiap pengadaan harus mendapat persetujuan pejabat berwenang sebelum PO diproses "
          "(BAI03.04: pengadaan komponen solusi dilakukan sesuai kebijakan dan kebutuhan yang disetujui).",
    "U2": "Peran pemohon dan penyetuju harus dipisahkan (segregation of duties) agar tidak ada satu orang "
          "yang mengajukan sekaligus menyetujui pengadaan.",
    "U3": "Nilai kontrak harus wajar terhadap estimasi anggaran. Selisih lebih dari {tol:.0%} wajib "
          "disertai justifikasi dan persetujuan ulang.",
    "U4": "Pengadaan bernilai di atas Rp {batas:,.0f} sebaiknya memperoleh lebih dari satu penawaran "
          "agar harga dan kualitas dapat dibandingkan.",
    "U5": "Pembelian tidak boleh dipecah menjadi beberapa transaksi kecil untuk menghindari batas "
          "nilai pembelian langsung (Rp {batas:,.0f}).",
    "U6": "Setiap pengadaan harus didukung kontrak dan berita acara serah terima (BAST) sebagai bukti "
          "barang diterima sesuai spesifikasi (BAI03.06: jaminan kualitas).",
    "U7": "Proses dari permintaan (PR) sampai pesanan (PO) diharapkan selesai paling lama {sla} hari.",
    "U8": "Ketergantungan pada satu vendor sebaiknya dibatasi (di bawah {vend:.0%} dari total nilai kontrak) "
          "untuk menjaga persaingan dan menekan risiko pasokan.",
}
RISIKO = {
    "U1": "Pembelian tanpa otorisasi dapat menimbulkan pemborosan anggaran, pembelian fiktif, atau "
          "kewajiban pembayaran yang tidak sah.",
    "U2": "Pengendalian internal lemah, sehingga peluang kecurangan dan kesalahan tidak terdeteksi meningkat.",
    "U3": "Perusahaan berpotensi membayar lebih mahal dari harga wajar atau terjadi penggelembungan harga (mark-up).",
    "U4": "Harga yang diperoleh mungkin bukan yang terbaik, dan terbuka peluang kedekatan tidak sehat dengan vendor.",
    "U5": "Pengendalian nilai batas dilewati, sehingga pengadaan yang seharusnya lewat tender atau persetujuan "
          "lebih tinggi tidak melalui proses semestinya.",
    "U6": "Tidak ada bukti sah bahwa barang diterima dan sesuai kontrak; sulit menagih garansi atau menuntut vendor.",
    "U7": "Kebutuhan operasional TI terlambat terpenuhi dan dapat mengganggu layanan.",
    "U8": "Bila vendor dominan bermasalah, layanan terganggu dan posisi tawar perusahaan melemah.",
}
WAKTU = {"Tinggi": "Segera (1 bulan)", "Sedang": "Jangka menengah (3 bulan)", "Rendah": "Jangka panjang (6 bulan)"}


def rp(x):
    return f"Rp {x:,.0f}".replace(",", ".")


def tanggal_id(d):
    return f"{d.day} {BULAN[d.month - 1]} {d.year}"


def esc(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ---------------------------------------------------------------------------
# Kalimat "kondisi" per temuan, dibuat dari data
# ---------------------------------------------------------------------------
def kondisi(kode, hasil, df, p):
    n = len(hasil)
    total = len(df)
    if kode == "U1":
        return (f"Ditemukan {n} dari {total} transaksi yang belum mendapat persetujuan atau tidak "
                f"mencantumkan nama penyetuju.")
    if kode == "U2":
        return f"Ditemukan {n} transaksi yang pemohon dan penyetujunya adalah orang yang sama."
    if kode == "U3":
        mx = hasil.selisih_harga_pct.max() * 100
        return (f"Ditemukan {n} transaksi dengan nilai kontrak lebih dari {p['tol']:.0%} di atas estimasi; "
                f"selisih tertinggi mencapai {mx:.0f}% dari estimasi.")
    if kode == "U4":
        return (f"Ditemukan {n} transaksi bernilai di atas batas pembelian langsung yang hanya memiliki "
                f"satu penawaran, dengan total nilai {rp(hasil.nilai_kontrak.sum())}.")
    if kode == "U5":
        return (f"Ditemukan {n} transaksi pembelian langsung ke vendor dan pemohon yang sama dalam waktu "
                f"kurang dari 14 hari. Nilai tiap transaksi berada tepat di bawah batas, namun totalnya "
                f"{rp(hasil.nilai_kontrak.sum())} sehingga melewati batas {rp(p['batas'])}.")
    if kode == "U6":
        tk = int((hasil.ada_kontrak == "Tidak").sum())
        tb = int((hasil.ada_bast == "Tidak").sum())
        return (f"Ditemukan {n} transaksi dengan dokumen tidak lengkap: {tk} tanpa kontrak dan {tb} tanpa BAST.")
    if kode == "U7":
        return (f"Ditemukan {n} transaksi yang membutuhkan lebih dari {p['sla']} hari dari PR ke PO; "
                f"terlama {int(hasil.hari_proses.max())} hari (rata-rata seluruh transaksi "
                f"{df.hari_proses.mean():.1f} hari).")
    if kode == "U8":
        v = hasil.index[0]
        return (f"Vendor {v} menguasai {hasil.iloc[0, 0] * 100:.1f}% dari total nilai kontrak, "
                f"melebihi batas {p['vend']:.0%}.")
    return ""


def daftar_id(kode, hasil):
    if kode == "U8":
        return ", ".join(hasil.index)
    ids = list(hasil["id_pengadaan"])
    return ", ".join(ids[:12]) + (f", dan {len(ids) - 12} lainnya" if len(ids) > 12 else "")


# ---------------------------------------------------------------------------
# Penyusun PDF
# ---------------------------------------------------------------------------
def buat_laporan(df, uji, ringkas, log, tujuan, parameter=None, data_contoh=True,
                 nama="[Nama Penyusun]", nim="[NIM]", instansi="[Program Studi / Universitas]"):
    """tujuan: path file atau objek BytesIO."""
    p = parameter or {"batas": audit.BATAS_PEMBELIAN_LANGSUNG, "tol": audit.TOLERANSI_HARGA,
                      "sla": audit.MAKS_HARI_PROSES, "vend": audit.BATAS_VENDOR_DOMINAN}

    ss = getSampleStyleSheet()
    body = ParagraphStyle("Body", parent=ss["BodyText"], fontName="Helvetica", fontSize=10.5,
                          leading=15, alignment=TA_JUSTIFY, spaceAfter=6)
    kecil = ParagraphStyle("Kecil", parent=body, fontSize=8.5, leading=11, alignment=0, spaceAfter=0)
    kecil_b = ParagraphStyle("KecilB", parent=kecil, fontName="Helvetica-Bold")
    h1 = ParagraphStyle("H1", parent=ss["Heading1"], fontName="Helvetica-Bold", fontSize=15,
                        spaceBefore=6, spaceAfter=8, textColor=colors.HexColor("#1f3a5f"))
    h2 = ParagraphStyle("H2", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=12,
                        spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#1f3a5f"))
    judul = ParagraphStyle("Judul", parent=h1, fontSize=22, leading=28, alignment=TA_CENTER, spaceAfter=10)
    sub = ParagraphStyle("Sub", parent=body, fontSize=12, alignment=TA_CENTER, leading=17)
    biru = colors.HexColor("#1f3a5f")
    abu = colors.HexColor("#f1f4f8")
    warna_risiko = {"Tinggi": colors.HexColor("#f8d7da"), "Sedang": colors.HexColor("#fff3cd"),
                    "Rendah": colors.HexColor("#d4edda")}

    def P(t, s=body):
        return Paragraph(t, s)

    def tabel(data, lebar, header=True, zebra=True):
        t = Table(data, colWidths=lebar, repeatRows=1 if header else 0)
        st = [("VALIGN", (0, 0), (-1, -1), "TOP"),
              ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#b8c2cf")),
              ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
              ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
        if header:
            st += [("BACKGROUND", (0, 0), (-1, 0), biru), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white)]
        if zebra:
            for i in range(1 if header else 0, len(data)):
                if i % 2 == 0:
                    st.append(("BACKGROUND", (0, i), (-1, i), abu))
        t.setStyle(TableStyle(st))
        return t

    hdr = ParagraphStyle("Hdr", parent=kecil_b, textColor=colors.white)

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawString(2 * cm, 1.2 * cm, "Laporan Audit Proses Pengadaan TI - COBIT 5 BAI03")
        canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Halaman {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(tujuan, pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm,
                            topMargin=2 * cm, bottomMargin=2 * cm,
                            title="Laporan Audit Proses Pengadaan TI", author=nama)
    S = []
    lebar = A4[0] - 4 * cm

    # ------------------------------ Sampul ------------------------------
    S += [Spacer(1, 5 * cm), P("LAPORAN AUDIT<br/>PROSES PENGADAAN TEKNOLOGI INFORMASI", judul),
          P("Studi Kasus 2 - Audit Sistem Informasi Berbantu Python", sub),
          P("Kerangka Acuan: COBIT 5 - Domain BAI03 (Manage Solutions Identification and Build)", sub),
          Spacer(1, 4 * cm), P(f"<b>{esc(nama)}</b><br/>NIM: {esc(nim)}<br/>{esc(instansi)}", sub),
          Spacer(1, 1 * cm), P(tanggal_id(date.today()), sub), PageBreak()]

    # ------------------------------ Ringkasan eksekutif ------------------------------
    jml_temuan = int((ringkas["Jumlah Kasus"] > 0).sum())
    tinggi = ringkas[(ringkas["Tingkat Risiko"] == "Tinggi") & (ringkas["Jumlah Kasus"] > 0)]
    S.append(P("RINGKASAN EKSEKUTIF", h1))
    S.append(P(
        f"Audit ini menguji {len(df)} transaksi pengadaan TI dengan total nilai kontrak {rp(df.nilai_kontrak.sum())}. "
        f"Dari delapan pengujian yang dilakukan, ditemukan {jml_temuan} jenis temuan, dengan "
        f"{len(tinggi)} di antaranya berisiko tinggi: "
        f"{esc('; '.join(tinggi['Temuan'].str.lower())) if len(tinggi) else 'tidak ada'}. "
        f"Temuan dipetakan ke praktik BAI03.04 (pengadaan komponen solusi) dan BAI03.06 (jaminan kualitas). "
        f"Rekomendasi utama adalah memperkuat otorisasi, pemisahan tugas, kewajaran harga, dan kelengkapan "
        f"dokumen pengadaan."))
    if data_contoh:
        S.append(P("<b>Catatan:</b> data yang dipakai dalam laporan ini adalah data simulasi yang dibuat untuk "
                   "keperluan pembelajaran, sehingga temuan bersifat ilustratif dan bukan kondisi organisasi "
                   "yang sebenarnya.", body))
    S.append(PageBreak())

    # ------------------------------ BAB 1 ------------------------------
    S.append(P("BAB 1. LATAR BELAKANG", h1))
    S.append(P("1.1 Latar Belakang", h2))
    S.append(P(
        "Pengadaan teknologi informasi (TI) mencakup pembelian perangkat keras, perangkat lunak, jaringan, "
        "dan layanan pendukung. Nilainya biasanya besar dan menjadi dasar berjalannya layanan TI organisasi. "
        "Karena itu proses pengadaan rawan terhadap penyimpangan, misalnya pembelian tanpa persetujuan, harga "
        "tidak wajar, pemilihan vendor yang tidak kompetitif, dan dokumen yang tidak lengkap."))
    S.append(P(
        "Audit sistem informasi berfungsi menilai apakah proses tersebut dikendalikan dengan baik. Audit manual "
        "atas ratusan transaksi memakan waktu dan mudah terlewat. Dengan bantuan Python, seluruh transaksi dapat "
        "diperiksa sekaligus (bukan hanya sampel), hasilnya konsisten, dan pengujian dapat diulang kapan saja "
        "ketika data baru tersedia."))
    S.append(P(
        "COBIT 5 dipilih sebagai acuan karena menyediakan kerangka tata kelola dan manajemen TI yang diakui "
        "secara luas. Proses pengadaan dibahas pada domain BAI (Build, Acquire and Implement), khususnya proses "
        "BAI03 yang mengelola identifikasi dan pembangunan solusi, termasuk pengadaan komponen solusinya."))
    S.append(P("1.2 Tujuan Audit", h2))
    for t in ["Menilai apakah proses pengadaan TI telah dijalankan sesuai kebijakan dan prinsip pengendalian internal.",
              "Mengidentifikasi transaksi yang menyimpang atau berisiko menggunakan analisis data dengan Python.",
              "Memetakan hasil temuan ke praktik BAI03 COBIT 5 dan menyusun rekomendasi perbaikan."]:
        S.append(P("- " + t))
    S.append(P("1.3 Ruang Lingkup dan Data", h2))
    periode = f"{tanggal_id(df.tgl_pr.min())} sampai {tanggal_id(df.tgl_pr.max())}"
    S.append(P(
        f"Ruang lingkup audit adalah transaksi pengadaan TI pada periode {periode}. Data berisi nomor "
        f"pengadaan, tanggal permintaan (PR), pesanan (PO), dan penerimaan, pemohon, penyetuju, vendor, kategori, "
        f"nilai estimasi, nilai kontrak, metode pengadaan, jumlah penawaran, serta status kontrak dan BAST. "
        f"Dari {log['jumlah_baris_awal']} baris awal, {len(log['data_ganda'])} baris ganda dihapus sehingga "
        f"{log['jumlah_baris_bersih']} transaksi dianalisis."))
    S.append(P("1.4 Metodologi", h2))
    langkah = [["Tahap", "Kegiatan"],
               ["1. Pengumpulan data", "Membaca data pengadaan (CSV) dengan pandas."],
               ["2. Pembersihan data", "Menyeragamkan format tanggal, menghapus data ganda, memeriksa nilai kosong."],
               ["3. Pengujian", "Menjalankan 8 uji audit: persetujuan, pemisahan tugas, kewajaran harga, persaingan, "
                                "pemecahan pembelian, kelengkapan dokumen, waktu proses, konsentrasi vendor."],
               ["4. Pemetaan", "Mengaitkan tiap temuan dengan praktik BAI03 dan menilai tingkat risikonya."],
               ["5. Pelaporan", "Menyajikan hasil pada dashboard Streamlit dan laporan PDF ini."]]
    S.append(tabel([[P(c, hdr if i == 0 else kecil) for c in r] for i, r in enumerate(langkah)],
                   [3.5 * cm, lebar - 3.5 * cm]))
    S.append(P("1.5 Parameter Pengujian", h2))
    S.append(P(
        f"Batas pembelian langsung {rp(p['batas'])}; toleransi selisih harga {p['tol']:.0%}; batas waktu proses "
        f"PR ke PO {p['sla']} hari; batas porsi satu vendor {p['vend']:.0%}. Parameter ini merupakan asumsi audit "
        f"dan dapat disesuaikan dengan kebijakan organisasi yang diaudit."))
    S.append(PageBreak())

    # ------------------------------ BAB 2 ------------------------------
    S.append(P("BAB 2. HASIL PENGOLAHAN DATA", h1))
    metode = df.metode.value_counts()
    ringkas_tbl = [["Indikator", "Nilai"],
                   ["Jumlah transaksi dianalisis", f"{len(df)}"],
                   ["Total nilai kontrak", rp(df.nilai_kontrak.sum())],
                   ["Rata-rata nilai kontrak", rp(df.nilai_kontrak.mean())],
                   ["Rata-rata lama proses PR ke PO", f"{df.hari_proses.mean():.1f} hari"],
                   ["Jumlah vendor", f"{df.vendor.nunique()}"],
                   ["Komposisi metode", "; ".join(f"{k}: {v}" for k, v in metode.items())]]
    S.append(tabel([[P(esc(c), hdr if i == 0 else kecil) for c in r] for i, r in enumerate(ringkas_tbl)],
                   [6 * cm, lebar - 6 * cm]))
    S.append(Spacer(1, 8))
    S.append(P("Gambar 1. Jumlah kasus per jenis temuan", kecil_b))
    S.append(Image(os.path.join(FOLDER, "grafik_temuan.png"), width=lebar, height=lebar * 0.5))
    S.append(Spacer(1, 6))
    dua = Table([[Image(os.path.join(FOLDER, "grafik_vendor.png"), width=lebar / 2 - 4, height=(lebar / 2 - 4) * 4 / 7),
                  Image(os.path.join(FOLDER, "grafik_waktu_proses.png"), width=lebar / 2 - 4, height=(lebar / 2 - 4) * 4 / 7)],
                 [P("Gambar 2. Konsentrasi vendor", kecil_b), P("Gambar 3. Lama proses PR ke PO", kecil_b)]])
    S.append(dua)
    S.append(PageBreak())

    # ------------------------------ BAB 3 ------------------------------
    S.append(P("BAB 3. TEMUAN AUDIT", h1))
    S.append(P("Setiap temuan disajikan dengan kondisi (apa yang ditemukan), kriteria (apa yang seharusnya), "
               "risiko (dampak bila dibiarkan), dan acuan BAI03. Rekomendasi dibahas pada Bab 4."))
    no = 0
    for _, r in ringkas.iterrows():
        k = r["Kode"]
        hasil = uji[k]
        if len(hasil) == 0:
            continue
        no += 1
        fmt = {"tol": p["tol"], "batas": p["batas"], "sla": p["sla"], "vend": p["vend"]}
        isi = [["Kondisi", kondisi(k, hasil, df, p)],
               ["Kriteria", KRITERIA[k].format(**fmt)],
               ["Risiko", RISIKO[k]],
               ["Acuan BAI03", r["Praktik BAI03"]],
               ["Tingkat risiko", r["Tingkat Risiko"]],
               ["Referensi data", daftar_id(k, hasil)]]
        t = tabel([[P(f"<b>{a}</b>", kecil), P(esc(b), kecil)] for a, b in isi],
                  [3 * cm, lebar - 3 * cm], header=False, zebra=False)
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, -1), abu),
                               ("BACKGROUND", (1, 4), (1, 4), warna_risiko[r["Tingkat Risiko"]])]))
        S.append(KeepTogether([P(f"Temuan {no}: {esc(r['Temuan'])}", h2), t, Spacer(1, 6)]))
    if no == 0:
        S.append(P("Tidak ditemukan penyimpangan pada pengujian yang dilakukan."))
    S.append(PageBreak())

    # ------------------------------ BAB 4 ------------------------------
    S.append(P("BAB 4. REKOMENDASI", h1))
    S.append(P("Rekomendasi diurutkan berdasarkan tingkat risiko. Temuan berisiko tinggi sebaiknya ditangani lebih dulu "
               "karena berkaitan langsung dengan otorisasi, pemisahan tugas, dan pengendalian nilai."))
    urut = ringkas[ringkas["Jumlah Kasus"] > 0].copy()
    urut["_o"] = urut["Tingkat Risiko"].map({"Tinggi": 0, "Sedang": 1, "Rendah": 2})
    urut = urut.sort_values(["_o", "Kode"])
    rek = [[P(c, hdr) for c in ["No", "Rekomendasi", "Terkait temuan", "Risiko", "Target waktu"]]]
    for i, (_, r) in enumerate(urut.iterrows(), 1):
        rek.append([P(str(i), kecil), P(esc(r["Rekomendasi"]), kecil), P(f"{r['Kode']} - {esc(r['Temuan'])}", kecil),
                    P(r["Tingkat Risiko"], kecil), P(WAKTU[r["Tingkat Risiko"]], kecil)])
    t = tabel(rek, [0.9 * cm, 6.2 * cm, 4.6 * cm, 1.6 * cm, lebar - 13.3 * cm], zebra=False)
    for i, (_, r) in enumerate(urut.iterrows(), 1):
        t.setStyle(TableStyle([("BACKGROUND", (3, i), (3, i), warna_risiko[r["Tingkat Risiko"]])]))
    S.append(t)
    S.append(Spacer(1, 8))
    S.append(P("<b>Rekomendasi umum</b>", body))
    for tx in ["Menerapkan alur persetujuan elektronik yang memblokir PO bila persetujuan belum lengkap.",
               "Menjalankan analisis data pengadaan ini secara berkala (misalnya bulanan) sebagai pemantauan berkelanjutan.",
               "Menetapkan kebijakan tertulis tentang batas nilai, jumlah penawaran minimum, dan dokumen wajib pengadaan.",
               "Menunjuk pihak independen untuk menindaklanjuti temuan dan melaporkan statusnya kepada manajemen."]:
        S.append(P("- " + tx))
    S.append(PageBreak())

    # ------------------------------ BAB 5 ------------------------------
    S.append(P("BAB 5. KESIMPULAN", h1))
    S.append(P(
        f"Berdasarkan pengujian atas {len(df)} transaksi, ditemukan {jml_temuan} jenis temuan dengan total "
        f"{int(ringkas['Jumlah Kasus'].sum())} kasus. Temuan berisiko tinggi berkaitan dengan "
        f"{esc(', '.join(tinggi['Temuan'].str.lower())) if len(tinggi) else 'tidak ada'}. "
        f"Hal ini menunjukkan bahwa pengendalian pada praktik BAI03.04 dan BAI03.06 belum sepenuhnya berjalan "
        f"efektif pada transaksi yang diuji."))
    S.append(P(
        "Penggunaan Python terbukti membantu auditor memeriksa seluruh populasi data dengan cepat, konsisten, "
        "dan dapat diulang. Hasil pengujian dapat langsung dipetakan ke COBIT dan disajikan melalui dashboard "
        "serta laporan otomatis. Temuan pada laporan ini adalah indikasi yang perlu dikonfirmasi lewat "
        "pemeriksaan dokumen dan wawancara sebelum dijadikan kesimpulan akhir."))

    # ------------------------------ Lampiran ------------------------------
    S.append(PageBreak())
    S.append(P("LAMPIRAN: RINGKASAN TEMUAN DAN PEMETAAN BAI03", h1))
    lamp = [[P(c, hdr) for c in ["Kode", "Temuan", "Kasus", "Praktik BAI03", "Risiko"]]]
    for _, r in ringkas.iterrows():
        lamp.append([P(r["Kode"], kecil), P(esc(r["Temuan"]), kecil), P(str(int(r["Jumlah Kasus"])), kecil),
                     P(esc(r["Praktik BAI03"]), kecil), P(r["Tingkat Risiko"], kecil)])
    S.append(tabel(lamp, [1.2 * cm, 6.3 * cm, 1.4 * cm, 5.6 * cm, lebar - 14.5 * cm]))

    doc.build(S, onFirstPage=lambda c, d: None, onLaterPages=footer)


def main():
    csv = os.path.join(FOLDER, "data_pengadaan.csv")
    if not os.path.exists(csv):
        audit.buat_data_contoh().to_csv(csv, index=False)
    df, log = audit.bersihkan(pd.read_csv(csv))
    uji = audit.uji_audit(df)
    ringkas = audit.ringkasan_temuan(uji)
    audit.buat_grafik(df, uji, ringkas)
    keluaran = os.path.join(FOLDER, "Laporan_Audit_Pengadaan_TI.pdf")
    buat_laporan(df, uji, ringkas, log, keluaran)
    print("Laporan dibuat:", keluaran)


if __name__ == "__main__":
    main()
