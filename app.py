"""
Website Powder Shade Alignment (Streamlit)

Cara menjalankan:
    pip install -r requirements.txt
    streamlit run app.py
"""
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

import analisis as A

st.set_page_config(page_title="Powder Shade Alignment", page_icon="🎨", layout="wide")


# =====================================================================
# FUNGSI BANTU TAMPILAN
# =====================================================================
def tampil_fig(fig):
    st.pyplot(fig, clear_figure=True)
    plt.close(fig)


def tabel(df, desimal=4, **kwargs):
    st.dataframe(df.round(desimal), width="stretch", **kwargs)


def tombol_csv(df, nama_file, label="⬇️ Download CSV"):
    st.download_button(label, df.to_csv(index=True).encode("utf-8"), file_name=nama_file,
                       mime="text/csv", key=nama_file)


def kotak_warna(lab, judul):
    hexa = A.rgb_hex(lab)
    st.markdown(
        f"<div style='display:flex;align-items:center;gap:10px'>"
        f"<div style='width:46px;height:46px;border-radius:8px;background:{hexa};border:1px solid #999'></div>"
        f"<div><b>{judul}</b><br><span style='font-size:0.85em'>L={lab[0]:.2f}, a={lab[1]:.2f}, "
        f"b={lab[2]:.2f}</span></div></div>", unsafe_allow_html=True)


# =====================================================================
# CACHE (supaya tidak menghitung ulang setiap klik)
# =====================================================================
@st.cache_data(show_spinner=False)
def muat_upload(file_bytes, nama_file, sheet):
    return A.baca_data(file_bytes, nama_file, sheet)


@st.cache_data(show_spinner=False)
def muat_contoh():
    return A.contoh_data()


@st.cache_data(show_spinner="Menghitung analisis stage...")
def jalankan_stage(df, key, versi):
    try:
        return A.analisis_stage(df, key, versi), None
    except Exception as e:
        return None, str(e)


@st.cache_data(show_spinner="Menghitung analisis Stage 3...")
def jalankan_stage3(df, versi, coef_versi_s1):
    try:
        return A.analisis_stage3(df, versi, coef_versi_s1), None
    except Exception as e:
        return None, str(e)


@st.cache_data(show_spinner="Mencari batas undertone...")
def jalankan_undertone(df, sumber, batas_manual, langkah):
    return A.analisis_undertone(df, sumber, batas_manual, langkah)


@st.cache_data(show_spinner=False)
def jalankan_level(df, sumber, undertone):
    return A.analisis_level(df, sumber, undertone)


# =====================================================================
# SIDEBAR
# =====================================================================
st.sidebar.title("🎨 Shade Alignment")
st.sidebar.header("1. Data")
file = st.sidebar.file_uploader("Upload data (.xlsx / .csv)", type=["xlsx", "xls", "csv"])
st.sidebar.download_button("⬇️ Download template Excel", A.buat_template(),
                           file_name="Template_Data_Shade_Alignment.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

df = None
if file is not None:
    file_bytes = file.getvalue()
    sheet = None
    if not file.name.lower().endswith(".csv"):
        sheets = A.daftar_sheet(file_bytes)
        sheet = st.sidebar.selectbox("Sheet", sheets, index=sheets.index("New_data") if "New_data" in sheets else 0)
    try:
        df = muat_upload(file_bytes, file.name, sheet)
    except Exception as e:
        st.sidebar.error(f"Gagal membaca file: {e}")
elif st.sidebar.checkbox("Coba pakai data contoh (dummy)"):
    df = muat_contoh()
    st.sidebar.warning("Sedang memakai data DUMMY (acak), bukan data asli.")

st.sidebar.header("2. Pengaturan")
VERSI = st.sidebar.selectbox("Versi regresi warna", list(A.VERSI_REGRESI),
                             format_func=lambda v: f"{v}  →  {A.VERSI_REGRESI[v]}")
URUT = st.sidebar.radio("Urutkan rekomendasi berdasarkan", ["P_match", "dE"],
                        format_func=lambda v: {"P_match": "P_match (paling yakin Match)",
                                               "dE": "dE (paling dekat ke bare face)"}[v])
TOP_K = st.sidebar.slider("Jumlah rekomendasi ditampilkan", 3, 20, 5)
N_GRID = st.sidebar.slider("Kerapatan grid Stage 1 & 2 (titik per kanal)", 5, 25, 15)
N_GRID_GAB = st.sidebar.slider("Kerapatan grid Stage 3 (titik per kanal)", 4, 10, 8,
                               help="Jumlah pasangan = (n³ liquid) × (n³ powder), jadi jangan terlalu besar.")

st.sidebar.header("3. Bare face yang dicari produknya")
c1, c2, c3 = st.sidebar.columns(3)
BARE = {"L": c1.number_input("L", value=62.5, step=0.1, format="%.2f"),
        "a": c2.number_input("a", value=11.8, step=0.1, format="%.2f"),
        "b": c3.number_input("b", value=16.4, step=0.1, format="%.2f")}
with st.sidebar:
    kotak_warna(np.array([BARE[k] for k in A.KANAL]), "Warna bare face")


# =====================================================================
# HALAMAN AWAL (BELUM ADA DATA)
# =====================================================================
st.title("Powder Shade Alignment")
if df is None:
    st.info("👈 Upload file data di sidebar untuk mulai, atau centang **Coba pakai data contoh**.")
    st.markdown("""
**Alur analisis di website ini**

| Bagian | Isi |
|---|---|
| Stage 1 — Liquid | Regresi Bare Face + Liquid → After Liquid, klasifikasi Match/Unmatch, rekomendasi liquid |
| Stage 2 — Powder | Regresi Bare Face + Powder → After Powder, klasifikasi Match/Unmatch, rekomendasi powder |
| Stage 3 — Gabungan | Bare Face + Liquid → After Liquid, lalu + Powder → hasil akhir, rekomendasi pasangan |
| Undertone C/N/W | Batas Hue manual, Hue otomatis, dan garis a–b terbaik (lengkap dengan persamaannya) |
| Level L | Batas nilai L untuk tiap level, semua data dan per undertone |
""")
    st.subheader("Format data")
    st.markdown("Satu baris = satu pengujian. Nama kolom harus **persis** seperti di bawah "
                "(download template di sidebar). Kolom yang tidak dipakai boleh dikosongkan; "
                "bagian analisis yang kolomnya belum lengkap akan dilewati.")
    tabel(pd.DataFrame({"Kolom": A.SEMUA_KOLOM, "Keterangan": [A.DESKRIPSI_KOLOM[c] for c in A.SEMUA_KOLOM]}),
          hide_index=True)
    st.stop()


# =====================================================================
# HITUNG SEMUA ANALISIS STAGE
# =====================================================================
kolom_kurang = A.cek_kolom(df)
hasil, error = {}, {}
for key in ["s1", "s2"]:
    if kolom_kurang[key]:
        hasil[key], error[key] = None, "Kolom belum lengkap: " + ", ".join(kolom_kurang[key])
    else:
        hasil[key], error[key] = jalankan_stage(df, key, VERSI)
if kolom_kurang["s3"]:
    hasil["s3"], error["s3"] = None, "Kolom belum lengkap: " + ", ".join(kolom_kurang["s3"])
elif hasil["s1"] is None:
    hasil["s3"], error["s3"] = None, "Stage 3 butuh hasil Stage 1. " + (error["s1"] or "")
else:
    hasil["s3"], error["s3"] = jalankan_stage3(df, VERSI, hasil["s1"]["coef_versi"])

katalog_liquid = A.buat_katalog(df, "{}_liquid") if not [c for c in A.kolom_lab("{}_liquid") if c not in df] else None
katalog_powder = (A.buat_katalog(df, "{}_produk", kolom_info=["Undertone", "L"])
                  if not [c for c in A.kolom_lab("{}_produk") if c not in df] else None)

tabs = st.tabs(["📋 Data", "1️⃣ Stage 1 — Liquid", "2️⃣ Stage 2 — Powder", "3️⃣ Stage 3 — Gabungan",
                "🎨 Undertone C/N/W", "💡 Level L", "📊 Ringkasan Model"])


# =====================================================================
# TAB DATA
# =====================================================================
with tabs[0]:
    st.subheader("Data yang dipakai")
    st.write(f"**{df.shape[0]} baris**, **{df.shape[1]} kolom**")
    nama_bagian = {"s1": "Stage 1 — Liquid", "s2": "Stage 2 — Powder", "s3": "Stage 3 — Gabungan",
                   "undertone": "Undertone C/N/W", "level": "Level L"}
    status = []
    for k, nama in nama_bagian.items():
        n = len(df[A.KOLOM_STAGE[k]].dropna()) if k in A.KOLOM_STAGE and not kolom_kurang[k] else None
        status.append({"Bagian": nama, "Status": "✅ Siap" if not kolom_kurang[k] else "⚠️ Kolom kurang",
                       "Baris lengkap": str(n) if n is not None else "-",
                       "Kolom yang belum ada": ", ".join(kolom_kurang[k]) or "-"})
    st.dataframe(pd.DataFrame(status), hide_index=True, width="stretch")
    st.dataframe(df, width="stretch", height=400)


# =====================================================================
# TAB STAGE (dipakai Stage 1, 2, 3)
# =====================================================================
def tampil_bagian_model(key, h):
    info = A.STAGE_INFO[key]
    y = h["y"]
    m1, m2, m3 = st.columns(3)
    m1.metric("Jumlah data", len(y)); m2.metric("Match", int(y.sum())); m3.metric("Unmatch", int((y == 0).sum()))

    st.subheader("1. Perbandingan versi regresi")
    st.caption("AUC & akurasi = hasil classifier Match/Unmatch (fitur dL, da, db, dE) kalau prediksi warnanya "
               "memakai versi tersebut. Versi yang dipilih di sidebar ditandai ⭐.")
    banding = h["banding"].copy()
    banding.insert(0, "", np.where(banding["Versi"] == VERSI, "⭐", ""))
    tabel(banding, hide_index=True)
    with st.expander("Lihat persamaan semua versi"):
        for v, pers in h["persamaan_versi"].items():
            st.markdown(f"**{v}** — `{A.VERSI_REGRESI[v]}`")
            tabel(pers, hide_index=True)

    st.subheader(f"2. Regresi versi terpilih: `{VERSI}`")
    tabel(h["persamaan_versi"][VERSI], hide_index=True)
    tampil_fig(A.fig_aktual_vs_prediksi(h["after"], h["pred"], info["nama_after"]))

    st.subheader("3. Uji tiap fitur delta (prediksi − bare face)")
    st.caption("Diurutkan dari AUC tertinggi. Kolom 'Aturan Match' = batas nilai fisik tiap fitur.")
    tabel(h["uji"], hide_index=True)
    with st.expander("Boxplot fitur delta (Match vs Unmatch)"):
        tampil_fig(A.fig_boxplot_fitur(h["delta"], y))

    st.subheader("4. Klasifikasi Match/Unmatch (fitur: dL, da, db, dE)")
    k = h["klas"]
    m1, m2, m3 = st.columns(3)
    m1.metric("AUC", f"{k['auc']:.4f}"); m2.metric("Akurasi", f"{k['acc']*100:.2f}%")
    m3.metric("Threshold P(Match)", f"{k['threshold']:.4f}")
    col1, col2 = st.columns([1, 2])
    with col1:
        tabel(k["laporan"], 3)
    with col2:
        tampil_fig(A.fig_klasifikasi(y, k["pred"]))


def tampil_rekomendasi(rek, nama_file, gabungan=False):
    for p in rek["peringatan"]:
        st.warning(p)
    kotak_warna(rek["bare"], "Bare face yang dicari")
    t1, t2 = st.tabs(["A. Dari katalog", "B. Dari range katalog"])
    with t1:
        st.write(f"**{rek['n_match_katalog']}** dari **{rek['n_katalog']}** "
                 f"{'pasangan' if gabungan else 'produk'} katalog diprediksi Match.")
        top = rek["katalog"].head(TOP_K)
        tabel(top, 3)
        if gabungan:
            kolom = {"Bare Face": np.tile(rek["bare"], (len(top), 1)),
                     "Liquid": top[["L_liq", "a_liq", "b_liq"]].values,
                     "Powder": top[["L_pow", "a_pow", "b_pow"]].values,
                     "Prediksi Hasil": top[["pred_L", "pred_a", "pred_b"]].values}
            label = [f"Liq {l} + Pow {q} (P={p:.2f})" for l, q, p in zip(top.baris_liquid, top.baris_powder, top.P_match)]
        else:
            kolom = {"Bare Face": np.tile(rek["bare"], (len(top), 1)), "Produk": top[A.KANAL].values,
                     "Prediksi Hasil": top[["pred_L", "pred_a", "pred_b"]].values}
            label = [f"Baris {i} (P={p:.2f})" for i, p in zip(top.index, top.P_match)]
        if len(top):
            tampil_fig(A.fig_swatch(kolom, label, f"Top-{len(top)} dari katalog"))
        tombol_csv(rek["katalog"], f"rekomendasi_katalog_{nama_file}.csv", "⬇️ Download semua hasil katalog (CSV)")
    with t2:
        st.write(f"Grid **{rek['n_grid']:,}** kombinasi → **{rek['n_match_grid']:,}** Match "
                 f"({100 * rek['n_match_grid'] / max(rek['n_grid'], 1):.1f}%).")
        st.caption("Range yang Match = semua kombinasi grid yang Match. Range terbaik = 10% kombinasi Match "
                   "dengan peringkat teratas (lebih sempit & lebih aman).")
        tabel(rek["range"], hide_index=True)
        st.markdown("**Contoh kombinasi terbaik dari grid** (+ produk katalog terdekat)")
        tabel(rek["grid_top"], 3)
        tombol_csv(rek["range"], f"rekomendasi_range_{nama_file}.csv", "⬇️ Download range (CSV)")


def tampil_swatch_data(kolom, h, judul):
    with st.expander("Swatch warna semua data"):
        tampil_fig(A.fig_swatch(kolom, [f"Baris {i}" for i in h["data"].index], judul))


# ---------------- Stage 1 ----------------
with tabs[1]:
    st.header(A.STAGE_INFO["s1"]["judul"])
    st.caption("Prediksi warna After Liquid dari Bare Face + Liquid, lalu dinilai Match/Unmatch terhadap bare face.")
    h = hasil["s1"]
    if h is None:
        st.error(error["s1"])
    else:
        tampil_bagian_model("s1", h)
        tampil_swatch_data({"Liquid": h["produk"], "Bare Face": h["bare"], "After Liquid": h["after"],
                            "Prediksi": h["pred"]}, h, "Stage 1: Liquid")
        st.subheader("5. Rekomendasi LIQUID")
        rek = A.rekomendasi_produk(BARE, katalog_liquid, h["coef"], h["klas"], "Liquid", h["data"],
                                   TOP_K, N_GRID, URUT)
        tampil_rekomendasi(rek, "liquid")

# ---------------- Stage 2 ----------------
with tabs[2]:
    st.header(A.STAGE_INFO["s2"]["judul"])
    st.caption("Prediksi warna After Powder dari Bare Face + Powder (kolom *_produk), lalu dinilai Match/Unmatch.")
    h = hasil["s2"]
    if h is None:
        st.error(error["s2"])
    else:
        tampil_bagian_model("s2", h)
        tampil_swatch_data({"Powder": h["produk"], "Bare Face": h["bare"], "After Powder": h["after"],
                            "Prediksi": h["pred"]}, h, "Stage 2: Powder")
        st.subheader("5. Rekomendasi POWDER")
        rek = A.rekomendasi_produk(BARE, katalog_powder, h["coef"], h["klas"], "Powder", h["data"],
                                   TOP_K, N_GRID, URUT)
        tampil_rekomendasi(rek, "powder")

# ---------------- Stage 3 ----------------
with tabs[3]:
    st.header(A.STAGE_INFO["s3"]["judul"])
    st.caption("Langkah 1: Bare Face + Liquid → After Liquid (regresi Stage 1, versi yang sama). "
               "Langkah 2: Prediksi After Liquid + Powder → Hasil Akhir.")
    h = hasil["s3"]
    if h is None:
        st.error(error["s3"])
    else:
        tampil_bagian_model("s3", h)
        tampil_swatch_data({"Bare Face": h["bare"], "Liquid": h["liquid"], "Powder": h["powder"],
                            "Pred After Liquid": h["after_liq"], "Aktual Liq+Pow": h["after"],
                            "Prediksi Liq+Pow": h["pred"]}, h, "Stage 3: Bare Face → Hasil Akhir")
        st.subheader("5. Rekomendasi pasangan LIQUID + POWDER")
        rek = A.rekomendasi_gabungan(BARE, katalog_liquid, katalog_powder, hasil["s1"]["coef"], h["coef"],
                                     h["klas"], h["data"], TOP_K, N_GRID_GAB, URUT)
        tampil_rekomendasi(rek, "gabungan", gabungan=True)


# =====================================================================
# TAB UNDERTONE
# =====================================================================
with tabs[4]:
    st.header("Klasifikasi Undertone (Cool / Neutral / Warm)")
    st.markdown("""
| Metode | Cara kerja | Batas di bidang a–b |
|---|---|---|
| **[1] Hue manual** | `Hue = arctan(b/a)`, batas diisi sendiri | `b = tan(Hue)·a` (garis lewat 0,0) |
| **[2] Hue otomatis** | Batas Hue dicari otomatis (decision tree) | `b = tan(Hue)·a` (garis lewat 0,0) |
| **[3] Garis a–b terbaik** | `skor = cos(θ)·a + sin(θ)·b`, θ dicoba 0°–180°, pilih akurasi tertinggi | `cos(θ)·a + sin(θ)·b = batas` ⇔ `b = m·a + c` |

Semua sudut dalam **derajat**. Pada metode [3], `cos(θ)` = bobot a dan `sin(θ)` = bobot b.
""")
    if kolom_kurang["undertone"]:
        st.error("Kolom 'Undertone' belum ada.")
    else:
        sumber_ada = [s for s, f in A.SUMBER_WARNA.items() if f.format("a") in df and f.format("b") in df]
        c1, c2, c3, c4 = st.columns([2, 1, 1, 1])
        sumber = c1.radio("Sumber warna", sumber_ada, horizontal=True, key="ut_sumber")
        b1 = c2.number_input("Batas manual Cool/Neutral (°)", value=42.0, step=0.5)
        b2 = c3.number_input("Batas manual Neutral/Warm (°)", value=52.0, step=0.5)
        langkah = c4.selectbox("Langkah sudut θ (°)", [1.0, 0.5, 0.25], index=1)

        hu = jalankan_undertone(df, sumber, (b1, b2), langkah)
        if hu is None:
            st.warning("Data undertone C/N/W untuk sumber ini belum cukup.")
        else:
            g = hu["garis"]
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Jumlah data", hu["n"])
            m2.metric("[1] Hue manual", f"{hu['acc_manual']*100:.1f}%")
            m3.metric("[2] Hue otomatis", f"{hu['acc_hue']*100:.1f}%")
            m4.metric("[3] Garis a–b", f"{g['acc']*100:.1f}%")

            with st.container(border=True):
                st.markdown("**🔍 Cek undertone dari nilai a, b**")
                k1, k2, k3 = st.columns([1, 1, 3])
                a_baru = k1.number_input("a", value=float(np.median(hu["a"])), step=0.1, key="ut_a")
                b_baru = k2.number_input("b", value=float(np.median(hu["b"])), step=0.1, key="ut_b")
                k3.dataframe(pd.DataFrame([A.prediksi_undertone(hu, a_baru, b_baru)]), hide_index=True,
                             width="stretch")

            tampil_fig(A.fig_undertone(hu, (a_baru, b_baru)))

            k1, k2, k3 = st.columns(3)
            with k1:
                st.markdown("**[1] Hue manual**")
                st.dataframe(pd.DataFrame({"Kondisi": [f"Hue < {b1}°", f"{b1}° ≤ Hue ≤ {b2}°", f"Hue > {b2}°"],
                                           "Hasil": ["Cool", "Neutral", "Warm"]}), hide_index=True)
                st.code("\n".join(A.garis_hue(s) for s in (b1, b2)), language=None)
            with k2:
                st.markdown("**[2] Hue otomatis**")
                st.dataframe(A.aturan_wilayah(hu["tree_hue"], hu["batas_hue"], "Hue", "°"), hide_index=True)
                st.code("\n".join(A.garis_hue(s) for s in hu["batas_hue"]), language=None)
            with k3:
                st.markdown("**[3] Garis a–b terbaik**")
                st.code(f"θ = {g['theta']}°\n"
                        f"cos(θ) = {g['w_a']:.4f}  (bobot a)\n"
                        f"sin(θ) = {g['w_b']:.4f}  (bobot b)\n"
                        f"skor = {A.teks_skor(g['w_a'], g['w_b'])}", language=None)
                st.dataframe(A.aturan_wilayah(g["tree"], g["batas"], "skor"), hide_index=True)
                st.code("\n".join(A.garis_skor(g["w_a"], g["w_b"], t) for t in g["batas"]), language=None)

        st.subheader("Ringkasan semua sumber warna")
        baris = []
        for s in sumber_ada:
            r = jalankan_undertone(df, s, (b1, b2), langkah)
            if r is None:
                continue
            g = r["garis"]
            baris.append({"Sumber": s, "N": r["n"], "Akurasi Hue manual": r["acc_manual"],
                          "Akurasi Hue otomatis": r["acc_hue"],
                          "Batas Hue otomatis (°)": ", ".join(f"{x:.2f}" for x in r["batas_hue"]),
                          "Akurasi Garis a-b": g["acc"], "θ (°)": g["theta"],
                          "Persamaan skor": A.teks_skor(g["w_a"], g["w_b"]),
                          "Batas skor": ", ".join(f"{x:.2f}" for x in g["batas"])})
        tabel(pd.DataFrame(baris), hide_index=True)


# =====================================================================
# TAB LEVEL L
# =====================================================================
with tabs[5]:
    st.header("Klasifikasi Level L")
    st.caption("Level L (kolom 'L') diprediksi dari nilai L warna. Batas antar level dicari otomatis "
               "(decision tree 1 fitur, jumlah daun = jumlah level). Area berwarna = level hasil prediksi.")
    if kolom_kurang["level"]:
        st.error("Kolom 'L' (level) belum ada.")
    else:
        sumber_ada = [s for s, f in A.SUMBER_WARNA.items() if f.format("L") in df]
        semua_level = sorted(df["L"].dropna().unique())
        warna = A.warna_level(semua_level)
        c1, c2 = st.columns([2, 2])
        sumber = c1.radio("Sumber warna", sumber_ada, horizontal=True, key="lv_sumber")
        mode = c2.radio("Kelompok data", ["Semua data", "Per undertone (C/N/W)"], horizontal=True)
        if mode == "Semua data":
            kelompok = [None]
        elif "Undertone" in df:
            kelompok = A.UNDERTONES
        else:
            st.error("Kolom 'Undertone' belum ada.")
            kelompok = []

        hasil_lv = {ut: jalankan_level(df, sumber, ut) for ut in kelompok}

        with st.container(border=True):
            st.markdown("**🔍 Cek level dari nilai L**")
            k1, k2, k3 = st.columns([1, 1, 3])
            kol_L = A.SUMBER_WARNA[sumber].format("L")
            L_baru = k1.number_input(f"Nilai {kol_L}", value=float(df[kol_L].median()), step=0.1)
            pilihan = [ut for ut, r in hasil_lv.items() if r is not None]
            if pilihan:
                ut_cek = k2.selectbox("Kelompok", pilihan, format_func=lambda u: u or "Semua data")
                r = hasil_lv[ut_cek]
                k3.metric("Prediksi level L", f"{float(r['tree'].predict([[L_baru]])[0]):g}")

        valid = {ut: r for ut, r in hasil_lv.items() if r is not None}
        if valid:
            fig, axes = plt.subplots(1, len(valid), figsize=(6.5 * len(valid), 5.5), sharey=True, squeeze=False)
            for ax, (ut, r) in zip(axes[0], valid.items()):
                A.gambar_level(ax, r, warna, semua_level, f"{sumber} — {ut or 'Semua data'}", L_baru)
            fig.tight_layout()
            tampil_fig(fig)
            cols = st.columns(len(valid))
            for col, (ut, r) in zip(cols, valid.items()):
                with col:
                    st.markdown(f"**Aturan {ut or 'Semua data'}**")
                    st.dataframe(r["aturan"], hide_index=True, width="stretch")
        for ut, r in hasil_lv.items():
            if r is None:
                st.warning(f"Data untuk {ut or 'Semua data'} belum cukup.")

        st.subheader("Ringkasan akurasi semua sumber")
        baris = []
        for s in sumber_ada:
            for ut in [None] + (A.UNDERTONES if "Undertone" in df else []):
                r = jalankan_level(df, s, ut)
                if r is not None:
                    baris.append({"Sumber": s, "Kelompok": ut or "Semua", "N": r["n"], "Akurasi": r["acc"],
                                  "Batas L": ", ".join(f"{x:.2f}" for x in r["batas"])})
        if baris:
            ringkas = pd.DataFrame(baris)
            urutan = [c for c in ["Semua"] + A.UNDERTONES if c in ringkas["Kelompok"].values]
            tabel(ringkas.pivot(index="Sumber", columns="Kelompok", values="Akurasi")[urutan])
            with st.expander("Detail batas L"):
                tabel(ringkas, hide_index=True)


# =====================================================================
# TAB RINGKASAN
# =====================================================================
with tabs[6]:
    st.header("Ringkasan Model per Stage & per Versi Regresi")
    nama_stage = {"s1": "1. Bare Face + Liquid", "s2": "2. Bare Face + Powder", "s3": "3. Gabungan Liquid + Powder"}
    daftar = [hasil[k]["banding"].assign(Stage=nama_stage[k]) for k in nama_stage if hasil[k] is not None]
    if not daftar:
        st.warning("Belum ada stage yang bisa dihitung.")
    else:
        semua = pd.concat(daftar, ignore_index=True)
        semua = semua[["Stage"] + [c for c in semua.columns if c != "Stage"]]
        for metrik in ["R² rata-rata", "AUC Match/Unmatch", "Akurasi Match/Unmatch"]:
            st.markdown(f"**{metrik}**")
            tabel(semua.pivot(index="Stage", columns="Versi", values=metrik)[list(A.VERSI_REGRESI)])
        with st.expander("Tabel lengkap"):
            tabel(semua, hide_index=True)
        tombol_csv(semua, "ringkasan_model.csv")
    st.caption("Catatan: semua akurasi dihitung di data yang sama dengan data latih (tanpa cross-validation) "
               "dan sampelnya kecil, jadi lebih tepat dibaca sebagai arah/tren.")
