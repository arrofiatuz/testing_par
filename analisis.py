"""
analisis.py — semua perhitungan Powder Shade Alignment (tanpa Streamlit).

Isi file ini sama dengan notebook gabungan, hanya saja fungsi-fungsinya MENGEMBALIKAN hasil
(tabel / angka / gambar) supaya bisa ditampilkan di website (app.py).
"""
import io

import numpy as np
import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from scipy.stats import mannwhitneyu
from skimage import color
from sklearn.linear_model import LinearRegression, LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import (r2_score, root_mean_squared_error, roc_auc_score, roc_curve,
                             accuracy_score, confusion_matrix, classification_report)

# =====================================================================
# KONSTANTA & FORMAT DATA
# =====================================================================
KANAL = ["L", "a", "b"]

VERSI_REGRESI = {
    "berganda":            "hasil = c1*input1 + c2*input2 + c0",
    "sum":                 "hasil = c*(input1 + input2) + c0",
    "sum_tanpa_intercept": "hasil = c*(input1 + input2)",
}

UNDERTONES = ["Cool", "Neutral", "Warm"]
WARNA_UNDERTONE = {"Cool": "#3d5a80", "Neutral": "#8d8d8d", "Warm": "#e07a5f"}
SUMBER_WARNA = {"Powder": "{}_produk", "Liquid": "{}_liquid", "Bare Face": "{}_bare face"}


def kolom_lab(fmt):
    return [fmt.format(k) for k in KANAL]


STAGE_INFO = {
    "s1": {"judul": "Stage 1 — Bare Face + Liquid", "produk": "{}_liquid", "after": "{}_after liquid",
           "match": "(Match/Unmatch)_liquid", "nama_produk": "Liquid", "nama_after": "After Liquid",
           "nama_input": ("bare_face", "liquid"), "nama_target": "after_liquid"},
    "s2": {"judul": "Stage 2 — Bare Face + Powder", "produk": "{}_produk", "after": "{}_after powder",
           "match": "(Match/Unmatch)_powder", "nama_produk": "Powder", "nama_after": "After Powder",
           "nama_input": ("bare_face", "powder"), "nama_target": "after_powder"},
    "s3": {"judul": "Stage 3 — Gabungan Liquid + Powder", "after": "{}_after liquid_powder",
           "match": "(Match/Unmatch)_liquid_powder", "nama_after": "After Liquid + Powder",
           "nama_input": ("pred_after_liquid", "powder"), "nama_target": "final"},
}

KOLOM_STAGE = {
    "s1": kolom_lab("{}_bare face") + kolom_lab("{}_liquid") + kolom_lab("{}_after liquid")
          + ["(Match/Unmatch)_liquid"],
    "s2": kolom_lab("{}_bare face") + kolom_lab("{}_produk") + kolom_lab("{}_after powder")
          + ["(Match/Unmatch)_powder"],
    "s3": kolom_lab("{}_bare face") + kolom_lab("{}_liquid") + kolom_lab("{}_produk")
          + kolom_lab("{}_after liquid_powder") + ["(Match/Unmatch)_liquid_powder"],
}

# Deskripsi tiap kolom (dipakai di halaman format data & template Excel)
DESKRIPSI_KOLOM = {
    "Undertone": "Undertone shade: Cool / Neutral / Warm (boleh nilai lain, tapi hanya C/N/W yang dianalisis)",
    "L": "Level L shade (angka, misal 1.0, 1.5, 2.0, ...)",
    "Olive": "Opsional (True/False)",
    "L_produk": "L powder (produk)", "a_produk": "a powder (produk)", "b_produk": "b powder (produk)",
    "L_liquid": "L liquid", "a_liquid": "a liquid", "b_liquid": "b liquid",
    "L_bare face": "L kulit asli", "a_bare face": "a kulit asli", "b_bare face": "b kulit asli",
    "L_after liquid": "L setelah liquid", "a_after liquid": "a setelah liquid", "b_after liquid": "b setelah liquid",
    "(Match/Unmatch)_liquid": "Hasil liquid: Match/Unmatch (atau True/False, 1/0)",
    "L_after liquid_powder": "L setelah liquid + powder", "a_after liquid_powder": "a setelah liquid + powder",
    "b_after liquid_powder": "b setelah liquid + powder",
    "(Match/Unmatch)_liquid_powder": "Hasil liquid + powder: Match/Unmatch (atau True/False, 1/0)",
    "L_after powder": "L setelah powder", "a_after powder": "a setelah powder", "b_after powder": "b setelah powder",
    "(Match/Unmatch)_powder": "Hasil powder: Match/Unmatch (atau True/False, 1/0)",
}
SEMUA_KOLOM = list(DESKRIPSI_KOLOM)


# =====================================================================
# LOAD DATA
# =====================================================================
def ke_biner(seri):
    peta = {"match": 1, "unmatch": 0, "true": 1, "false": 0, "1": 1, "0": 0, "1.0": 1, "0.0": 0}
    return seri.map(lambda v: np.nan if pd.isna(v) else peta.get(str(v).strip().lower(), np.nan))


def bersihkan_data(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for c in df.columns:
        if c.split("_")[0] in KANAL and "_" in c:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        if c.startswith("(Match/Unmatch)"):
            df[c] = ke_biner(df[c])
    if "L" in df.columns:
        df["L"] = pd.to_numeric(df["L"], errors="coerce")
    if "Undertone" in df.columns:
        df["Undertone"] = df["Undertone"].astype(str).str.strip().str.capitalize().replace("Nan", np.nan)
    return df


def daftar_sheet(file_bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


def baca_data(file_bytes, nama_file, sheet=None):
    if nama_file.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet or 0)
    return bersihkan_data(df)


def cek_kolom(df):
    """Kolom yang belum ada untuk tiap bagian analisis."""
    hasil = {k: [c for c in kol if c not in df.columns] for k, kol in KOLOM_STAGE.items()}
    hasil["undertone"] = [c for c in ["Undertone"] if c not in df.columns]
    hasil["level"] = [c for c in ["L"] if c not in df.columns]
    return hasil


def buat_template():
    """File Excel template: sheet New_data (header kosong) + sheet Petunjuk."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(columns=SEMUA_KOLOM).to_excel(w, sheet_name="New_data", index=False)
        pd.DataFrame({"Kolom": SEMUA_KOLOM, "Keterangan": [DESKRIPSI_KOLOM[c] for c in SEMUA_KOLOM]}) \
          .to_excel(w, sheet_name="Petunjuk", index=False)
    return buf.getvalue()


def contoh_data(seed=0, n=79):
    """Data DUMMY (acak) untuk mencoba website. Bukan data asli."""
    rng = np.random.default_rng(seed)
    ut = rng.choice(UNDERTONES, n, p=[.3, .3, .4])
    lv = rng.choice([1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0], n)
    df = pd.DataFrame({"Undertone": ut, "L": lv, "Olive": False})
    df["L_produk"] = 82 - 3 * lv + rng.normal(0, 1, n)
    hue = np.where(ut == "Cool", 55, np.where(ut == "Neutral", 61, 67)) + rng.normal(0, 2, n)
    C = rng.uniform(18, 28, n)
    df["a_produk"], df["b_produk"] = C * np.cos(np.radians(hue)), C * np.sin(np.radians(hue))
    liq = np.ones(n, bool); liq[:19] = False; liq[40:51] = False
    df["L_liquid"] = np.where(liq, rng.uniform(55, 75, n), np.nan)
    df["a_liquid"] = np.where(liq, rng.uniform(10, 19, n), np.nan)
    df["b_liquid"] = np.where(liq, rng.uniform(19, 28, n), np.nan)
    df["L_bare face"], df["a_bare face"], df["b_bare face"] = (rng.uniform(54, 66, n), rng.uniform(9, 14.5, n),
                                                               rng.uniform(11, 20, n))
    for k in KANAL:
        df[f"{k}_after liquid"] = np.where(liq, 0.8 * df[f"{k}_bare face"] + 0.15 * df[f"{k}_liquid"]
                                           + rng.normal(0, 1, n), np.nan)
    df["(Match/Unmatch)_liquid"] = np.where(liq, rng.random(n) > .5, None)
    for k in KANAL:
        df[f"{k}_after liquid_powder"] = np.where(liq, 0.8 * df[f"{k}_after liquid"] + 0.2 * df[f"{k}_produk"]
                                                  + rng.normal(0, 1, n), np.nan)
    df["(Match/Unmatch)_liquid_powder"] = np.where(liq, rng.random(n) > .5, None)
    for k in KANAL:
        df[f"{k}_after powder"] = 0.7 * df[f"{k}_bare face"] + 0.25 * df[f"{k}_produk"] + rng.normal(0, 1, n)
    df["(Match/Unmatch)_powder"] = rng.random(n) > .5
    return bersihkan_data(df)


# =====================================================================
# REGRESI WARNA
# =====================================================================
def ambil_lab(data, fmt):
    return data[kolom_lab(fmt)].values.astype(float)


def prediksi_lab(coef, input1, input2):
    return coef[:, 0] * input1 + coef[:, 1] * input2 + coef[:, 2]


def tanda(x):
    return f"{'+' if x >= 0 else '-'} {abs(x):.4f}"


def fit_regresi(X1, X2, Y, nama_input=("input1", "input2"), nama_target="hasil", versi="berganda"):
    """Regresi per kanal L, a, b -> coef (3x3: [c1, c2, c0]), prediksi (N,3), tabel persamaan."""
    n1, n2 = nama_input
    coef = np.zeros((3, 3))
    persamaan = []
    for j, k in enumerate(KANAL):
        x1, x2, y = X1[:, j], X2[:, j], Y[:, j]
        if versi == "berganda":
            reg = LinearRegression().fit(np.column_stack([x1, x2]), y)
            c1, c2, c0 = reg.coef_[0], reg.coef_[1], reg.intercept_
            rumus = f"{c1:.4f}*{n1} {tanda(c2)}*{n2} {tanda(c0)}"
        else:
            pakai_c0 = versi == "sum"
            reg = LinearRegression(fit_intercept=pakai_c0).fit((x1 + x2).reshape(-1, 1), y)
            c1 = c2 = reg.coef_[0]
            c0 = reg.intercept_ if pakai_c0 else 0.0
            rumus = f"{c1:.4f}*({n1} + {n2})" + (f" {tanda(c0)}" if pakai_c0 else "")
        coef[j] = [c1, c2, c0]
        pred_k = c1 * x1 + c2 * x2 + c0
        persamaan.append({"Kanal": k, "Persamaan": f"{k}_{nama_target} = {rumus}",
                          "R²": r2_score(y, pred_k), "RMSE": root_mean_squared_error(y, pred_k)})
    return coef, prediksi_lab(coef, X1, X2), pd.DataFrame(persamaan)


# =====================================================================
# FITUR DELTA & KLASIFIKASI MATCH/UNMATCH
# =====================================================================
def fitur_match(pred, bare):
    d = pred - bare
    return np.column_stack([d, np.sqrt((d ** 2).sum(axis=1))])


def tabel_fitur_delta(pred, bare):
    def turunan(L, a, b):
        return {"L": L, "a": a, "b": b, "a-b": a - b, "2a-b": 2 * a - b, "2b-a": 2 * b - a,
                "b/a": b / a, "Chroma": np.hypot(a, b), "Hue": np.degrees(np.arctan2(b, a)) % 360}
    p, s = turunan(*pred.T), turunan(*bare.T)
    hasil = pd.DataFrame({f"d{n}": p[n] - s[n] for n in p})
    hasil["dE"] = np.sqrt(hasil["dL"] ** 2 + hasil["da"] ** 2 + hasil["db"] ** 2)
    return hasil


def uji_fitur(tabel, y):
    y = np.asarray(y).astype(int)
    baris = []
    for col in tabel.columns:
        x = tabel[col].values.astype(float)
        auc = roc_auc_score(y, x)
        arah = ">="
        if auc < 0.5:
            x, auc, arah = -x, 1 - auc, "<="
        fpr, tpr, thr = roc_curve(y, x)
        ok = np.isfinite(thr)
        batas = thr[ok][np.argmax(tpr[ok] - fpr[ok])]
        acc = accuracy_score(y, (x >= batas).astype(int))
        cutoff = batas if arah == ">=" else -batas
        p_val = mannwhitneyu(x[y == 1], x[y == 0]).pvalue
        baris.append({"Fitur": col, "AUC": auc, "Akurasi": acc, "Aturan Match": f"{col} {arah} {cutoff:.3f}",
                      "p-value": p_val, "Signifikan (p<0.05)": "Ya" if p_val < 0.05 else "-"})
    return pd.DataFrame(baris).sort_values("AUC", ascending=False).reset_index(drop=True)


def fit_klasifikasi(X, y):
    y = np.asarray(y).astype(int)
    model = LogisticRegression(class_weight="balanced", random_state=42).fit(X, y)
    probs = model.predict_proba(X)[:, 1]
    auc = roc_auc_score(y, probs)
    fpr, tpr, thr = roc_curve(y, probs)
    ok = np.isfinite(thr) & (thr <= 1)
    threshold = thr[ok][np.argmax(tpr[ok] - fpr[ok])]
    pred = (probs >= threshold).astype(int)
    laporan = pd.DataFrame(classification_report(y, pred, target_names=["Unmatch", "Match"],
                                                 zero_division=0, output_dict=True)).T
    return {"model": model, "threshold": threshold, "auc": auc, "acc": accuracy_score(y, pred),
            "pred": pred, "probs": probs, "laporan": laporan}


def peluang_match(klas, pred, bare):
    X = fitur_match(pred, bare)
    p = klas["model"].predict_proba(X)[:, 1]
    return p, X[:, 3], p >= klas["threshold"]


def bandingkan_versi(X1, X2, Y, bare, y, nama_input, nama_target):
    baris, semua_coef, semua_persamaan = [], {}, {}
    for versi, rumus in VERSI_REGRESI.items():
        x1 = X1[versi] if isinstance(X1, dict) else X1
        coef, pred, pers = fit_regresi(x1, X2, Y, nama_input, nama_target, versi)
        klas = fit_klasifikasi(fitur_match(pred, bare), y)
        semua_coef[versi], semua_persamaan[versi] = coef, pers
        r2 = pers["R²"].values
        baris.append({"Versi": versi, "Rumus": rumus, "R² L": r2[0], "R² a": r2[1], "R² b": r2[2],
                      "R² rata-rata": r2.mean(), "RMSE rata-rata": pers["RMSE"].mean(),
                      "AUC Match/Unmatch": klas["auc"], "Akurasi Match/Unmatch": klas["acc"]})
    return pd.DataFrame(baris), semua_coef, semua_persamaan


# =====================================================================
# ANALISIS PER STAGE
# =====================================================================
def _cek_data_stage(d, y):
    if len(d) < 5:
        raise ValueError(f"Data terlalu sedikit ({len(d)} baris lengkap).")
    if len(np.unique(y)) < 2:
        raise ValueError("Kolom Match/Unmatch hanya berisi satu kelas (butuh Match dan Unmatch).")


def analisis_stage(df, key, versi):
    """Stage 1 (s1) atau Stage 2 (s2)."""
    info = STAGE_INFO[key]
    d = df[KOLOM_STAGE[key]].dropna().copy()
    y = d[info["match"]].astype(int).values
    _cek_data_stage(d, y)
    bare, produk, after = ambil_lab(d, "{}_bare face"), ambil_lab(d, info["produk"]), ambil_lab(d, info["after"])
    banding, coef_versi, pers_versi = bandingkan_versi(bare, produk, after, bare, y,
                                                       info["nama_input"], info["nama_target"])
    coef = coef_versi[versi]
    pred = prediksi_lab(coef, bare, produk)
    delta = tabel_fitur_delta(pred, bare)
    return {"data": d, "y": y, "bare": bare, "produk": produk, "after": after,
            "banding": banding, "coef_versi": coef_versi, "persamaan_versi": pers_versi,
            "coef": coef, "pred": pred, "delta": delta, "uji": uji_fitur(delta, y),
            "klas": fit_klasifikasi(fitur_match(pred, bare), y)}


def analisis_stage3(df, versi, coef_versi_s1):
    info = STAGE_INFO["s3"]
    d = df[KOLOM_STAGE["s3"]].dropna().copy()
    y = d[info["match"]].astype(int).values
    _cek_data_stage(d, y)
    bare, liquid = ambil_lab(d, "{}_bare face"), ambil_lab(d, "{}_liquid")
    powder, final = ambil_lab(d, "{}_produk"), ambil_lab(d, info["after"])
    after_per_versi = {v: prediksi_lab(coef_versi_s1[v], bare, liquid) for v in VERSI_REGRESI}
    banding, coef_versi, pers_versi = bandingkan_versi(after_per_versi, powder, final, bare, y,
                                                       info["nama_input"], info["nama_target"])
    coef = coef_versi[versi]
    after_liq = after_per_versi[versi]
    pred = prediksi_lab(coef, after_liq, powder)
    delta = tabel_fitur_delta(pred, bare)
    return {"data": d, "y": y, "bare": bare, "liquid": liquid, "powder": powder, "after": final,
            "after_liq": after_liq, "banding": banding, "coef_versi": coef_versi,
            "persamaan_versi": pers_versi, "coef": coef, "pred": pred, "delta": delta,
            "uji": uji_fitur(delta, y), "klas": fit_klasifikasi(fitur_match(pred, bare), y)}


# =====================================================================
# REKOMENDASI PRODUK
# =====================================================================
def buat_katalog(data, fmt, kolom_info=None):
    kolom = kolom_lab(fmt)
    info = [c for c in (kolom_info or []) if c in data.columns]
    katalog = data[kolom + info].dropna(subset=kolom).drop_duplicates(subset=kolom)
    katalog = katalog.rename(columns={**dict(zip(kolom, KANAL)), **{c: f"label_{c}" for c in info}})
    katalog.index.name = "baris_data"
    return katalog


def buat_grid(katalog, n_grid):
    sumbu = [np.linspace(katalog[k].min(), katalog[k].max(), n_grid) for k in KANAL]
    mesh = np.meshgrid(*sumbu, indexing="ij")
    return np.column_stack([m.ravel() for m in mesh])


def urutkan(tabel, urut):
    if urut == "dE":
        return tabel.sort_values(["Match", "dE"], ascending=[False, True])
    return tabel.sort_values(["P_match", "dE"], ascending=[False, True])


def produk_terdekat(titik, katalog):
    jarak = np.linalg.norm(titik[:, None, :] - katalog[KANAL].values[None, :, :], axis=2)
    return katalog.index.values[jarak.argmin(axis=1)], jarak.min(axis=1)


def peringatan_rentang(bare, data):
    pesan = []
    for j, k in enumerate(KANAL):
        lo, hi = data[f"{k}_bare face"].min(), data[f"{k}_bare face"].max()
        if not lo <= bare[j] <= hi:
            pesan.append(f"{k} bare face = {bare[j]:.2f} di luar rentang data latih ({lo:.1f} – {hi:.1f}), "
                         "hasilnya ekstrapolasi.")
    return pesan


def teks_range(nilai):
    return f"{nilai.min():.2f} – {nilai.max():.2f}" if len(nilai) else "tidak ada"


def tabel_range(titik_match, titik_terbaik, katalog, nama):
    return pd.DataFrame([{"Produk": nama, "Kanal": k,
                          "Range katalog": teks_range(katalog[k].values),
                          "Range yang Match": teks_range(titik_match[:, j]),
                          "Range terbaik (top 10%)": teks_range(titik_terbaik[:, j])}
                         for j, k in enumerate(KANAL)])


def ambil_top_persen(tabel, persen=10):
    return tabel.head(max(1, int(len(tabel) * persen / 100)))


def rekomendasi_produk(bare_face, katalog, coef, klas, nama_produk, data_latih,
                       top_k=5, n_grid=15, urut="P_match"):
    bare = np.array([bare_face[k] for k in KANAL], dtype=float)

    # A. Dari katalog
    pred = prediksi_lab(coef, bare, katalog[KANAL].values)
    p, dE, match = peluang_match(klas, pred, bare)
    hasil_katalog = katalog.copy()
    hasil_katalog[["pred_L", "pred_a", "pred_b"]] = pred
    hasil_katalog["dE"], hasil_katalog["P_match"], hasil_katalog["Match"] = dE, p, match
    hasil_katalog = urutkan(hasil_katalog, urut)

    # B. Dari range katalog (grid)
    G = buat_grid(katalog, n_grid)
    pred_g = prediksi_lab(coef, bare, G)
    p_g, dE_g, match_g = peluang_match(klas, pred_g, bare)
    grid = pd.DataFrame(G, columns=KANAL)
    grid[["pred_L", "pred_a", "pred_b"]] = pred_g
    grid["dE"], grid["P_match"], grid["Match"] = dE_g, p_g, match_g
    grid = urutkan(grid[grid.Match], urut)
    hasil_range = tabel_range(G[match_g], ambil_top_persen(grid)[KANAL].values, katalog, nama_produk)
    grid_top = grid.head(top_k).reset_index(drop=True)
    if len(grid_top):
        grid_top["katalog_terdekat"], grid_top["jarak_ke_katalog"] = produk_terdekat(grid_top[KANAL].values, katalog)

    return {"bare": bare, "peringatan": peringatan_rentang(bare, data_latih),
            "katalog": hasil_katalog, "n_katalog": len(katalog), "n_match_katalog": int(match.sum()),
            "range": hasil_range, "grid_top": grid_top, "n_grid": len(G), "n_match_grid": int(match_g.sum())}


def rekomendasi_gabungan(bare_face, katalog_liq, katalog_pow, coef_s1, coef_s3, klas, data_latih,
                         top_k=5, n_grid=8, urut="P_match"):
    bare = np.array([bare_face[k] for k in KANAL], dtype=float)

    def hitung(LQ, PW):
        i_liq = np.repeat(np.arange(len(LQ)), len(PW))
        i_pow = np.tile(np.arange(len(PW)), len(LQ))
        final = prediksi_lab(coef_s3, prediksi_lab(coef_s1, bare, LQ[i_liq]), PW[i_pow])
        p, dE, match = peluang_match(klas, final, bare)
        return i_liq, i_pow, final, p, dE, match

    def susun(LQ, PW, i_liq, i_pow, final, p, dE, match):
        return pd.DataFrame({
            "L_liq": LQ[i_liq, 0], "a_liq": LQ[i_liq, 1], "b_liq": LQ[i_liq, 2],
            "L_pow": PW[i_pow, 0], "a_pow": PW[i_pow, 1], "b_pow": PW[i_pow, 2],
            "pred_L": final[:, 0], "pred_a": final[:, 1], "pred_b": final[:, 2],
            "dE": dE, "P_match": p, "Match": match})

    # A. Dari katalog
    LQ, PW = katalog_liq[KANAL].values, katalog_pow[KANAL].values
    i_liq, i_pow, final, p, dE, match = hitung(LQ, PW)
    hasil_katalog = susun(LQ, PW, i_liq, i_pow, final, p, dE, match)
    hasil_katalog.insert(0, "baris_powder", katalog_pow.index.values[i_pow])
    hasil_katalog.insert(0, "baris_liquid", katalog_liq.index.values[i_liq])
    for c in katalog_pow.columns:
        if c.startswith("label_"):
            hasil_katalog[f"{c}_powder"] = katalog_pow[c].values[i_pow]
    hasil_katalog = urutkan(hasil_katalog, urut).reset_index(drop=True)

    # B. Dari range katalog
    GL, GP = buat_grid(katalog_liq, n_grid), buat_grid(katalog_pow, n_grid)
    i_lg, i_pg, final_g, p_g, dE_g, match_g = hitung(GL, GP)
    grid = urutkan(susun(GL, GP, i_lg, i_pg, final_g, p_g, dE_g, match_g).loc[match_g], urut)
    terbaik = ambil_top_persen(grid)
    hasil_range = pd.concat([
        tabel_range(GL[i_lg[match_g]], terbaik[["L_liq", "a_liq", "b_liq"]].values, katalog_liq, "Liquid"),
        tabel_range(GP[i_pg[match_g]], terbaik[["L_pow", "a_pow", "b_pow"]].values, katalog_pow, "Powder")],
        ignore_index=True)
    grid_top = grid.head(top_k).reset_index(drop=True)
    if len(grid_top):
        grid_top["liquid_terdekat"], grid_top["jarak_liquid"] = produk_terdekat(
            grid_top[["L_liq", "a_liq", "b_liq"]].values, katalog_liq)
        grid_top["powder_terdekat"], grid_top["jarak_powder"] = produk_terdekat(
            grid_top[["L_pow", "a_pow", "b_pow"]].values, katalog_pow)

    return {"bare": bare, "peringatan": peringatan_rentang(bare, data_latih),
            "katalog": hasil_katalog, "n_katalog": len(match), "n_match_katalog": int(match.sum()),
            "range": hasil_range, "grid_top": grid_top, "n_grid": len(match_g), "n_match_grid": int(match_g.sum())}


# =====================================================================
# KLASIFIKASI UNDERTONE (C / N / W)
# =====================================================================
def hitung_hue(a, b):
    return np.degrees(np.arctan2(b, a)) % 360


def cari_batas(x, y):
    tree = DecisionTreeClassifier(max_leaf_nodes=len(np.unique(y)), random_state=42)
    tree.fit(np.asarray(x).reshape(-1, 1), y)
    batas = sorted(t for t in tree.tree_.threshold if t != -2)
    return tree, batas, accuracy_score(y, tree.predict(np.asarray(x).reshape(-1, 1)))


def undertone_hue_manual(hue, batas):
    return np.where(hue < batas[0], "Cool", np.where(hue <= batas[1], "Neutral", "Warm"))


def cari_garis_terbaik(a, b, y, langkah=0.5):
    terbaik = {"acc": -1}
    for theta in np.arange(0, 180, langkah):
        w_a, w_b = np.cos(np.radians(theta)), np.sin(np.radians(theta))
        tree, batas, acc = cari_batas(w_a * a + w_b * b, y)
        if acc > terbaik["acc"]:
            terbaik = {"theta": float(theta), "w_a": w_a, "w_b": w_b, "tree": tree, "batas": batas, "acc": acc}
    return terbaik


def aturan_wilayah(tree, batas, nama, satuan=""):
    tepi = [-np.inf] + list(batas) + [np.inf]
    aturan = []
    for lo, hi in zip(tepi[:-1], tepi[1:]):
        titik = hi - 1 if lo == -np.inf else (lo + 1 if hi == np.inf else (lo + hi) / 2)
        kelas = tree.predict([[titik]])[0]
        if lo == -np.inf:
            teks = f"{nama} ≤ {hi:.2f}{satuan}"
        elif hi == np.inf:
            teks = f"{nama} > {lo:.2f}{satuan}"
        else:
            teks = f"{lo:.2f}{satuan} < {nama} ≤ {hi:.2f}{satuan}"
        aturan.append({"Kondisi": teks, "Hasil": kelas})
    return pd.DataFrame(aturan)


def garis_hue(sudut):
    return f"Hue = {sudut:.2f}°  ⇔  b = {np.tan(np.radians(sudut)):.4f}·a"


def teks_skor(w_a, w_b):
    return f"{w_a:.4f}·a {'+' if w_b >= 0 else '−'} {abs(w_b):.4f}·b"


def garis_skor(w_a, w_b, t):
    kiri = f"{teks_skor(w_a, w_b)} = {t:.2f}"
    if abs(w_b) < 1e-9:
        return f"{kiri}  ⇔  a = {t / w_a:.4f}"
    m, c = -w_a / w_b, t / w_b
    return f"{kiri}  ⇔  b = {m:.4f}·a {'+' if c >= 0 else '−'} {abs(c):.4f}"


def analisis_undertone(df, sumber, batas_manual=(42, 52), langkah=0.5):
    fmt = SUMBER_WARNA[sumber]
    ka, kb = fmt.format("a"), fmt.format("b")
    if ka not in df.columns or kb not in df.columns or "Undertone" not in df.columns:
        return None
    d = df[df["Undertone"].isin(UNDERTONES)].dropna(subset=[ka, kb])
    if len(d) < 3 or d["Undertone"].nunique() < 2:
        return None
    a, b, y = d[ka].values, d[kb].values, d["Undertone"].values
    hue = hitung_hue(a, b)
    tree_hue, batas_hue, acc_hue = cari_batas(hue, y)
    return {"sumber": sumber, "a": a, "b": b, "y": y, "n": len(d), "batas_manual": list(batas_manual),
            "acc_manual": accuracy_score(y, undertone_hue_manual(hue, batas_manual)),
            "tree_hue": tree_hue, "batas_hue": batas_hue, "acc_hue": acc_hue,
            "garis": cari_garis_terbaik(a, b, y, langkah)}


def prediksi_undertone(hasil, a, b):
    hue = hitung_hue(a, b)
    g = hasil["garis"]
    return {"Hue (°)": round(float(hue), 2),
            "[1] Hue manual": str(undertone_hue_manual(np.array([hue]), hasil["batas_manual"])[0]),
            "[2] Hue otomatis": str(hasil["tree_hue"].predict([[hue]])[0]),
            "[3] Garis a-b": str(g["tree"].predict([[g["w_a"] * a + g["w_b"] * b]])[0])}


# =====================================================================
# KLASIFIKASI LEVEL L
# =====================================================================
def analisis_level(df, sumber, undertone=None):
    kol = SUMBER_WARNA[sumber].format("L")
    if kol not in df.columns or "L" not in df.columns:
        return None
    d = df.dropna(subset=[kol, "L"])
    if undertone:
        d = d[d["Undertone"] == undertone]
    if len(d) < 2 or d["L"].nunique() < 2:
        return None
    x, level = d[kol].values, d["L"].values
    tree, batas, acc = cari_batas(x, level.astype(str))
    return {"kolom": kol, "x": x, "level": level, "n": len(d), "tree": tree, "batas": batas, "acc": acc,
            "aturan": aturan_wilayah(tree, batas, kol)}


# =====================================================================
# GAMBAR
# =====================================================================
def lab_to_rgb(lab):
    lab = np.asarray(lab, dtype=float).reshape(-1, 1, 3)
    return np.clip(color.lab2rgb(lab).reshape(-1, 3), 0, 1)


def rgb_hex(lab):
    r, g, b = (lab_to_rgb(lab)[0] * 255).round().astype(int)
    return f"#{r:02x}{g:02x}{b:02x}"


def fig_aktual_vs_prediksi(aktual, pred, nama):
    fig, axes = plt.subplots(1, 3, figsize=(18, 4))
    idx = np.arange(len(aktual))
    for j, k in enumerate(KANAL):
        axes[j].scatter(idx, aktual[:, j], label="Aktual", color="royalblue", s=40, alpha=0.8)
        axes[j].scatter(idx, pred[:, j], label="Prediksi", color="darkorange", marker="x", s=40)
        axes[j].set_title(f"{k} {nama}", fontweight="bold"); axes[j].set_xlabel("Index data")
        axes[j].grid(True, linestyle="--", alpha=0.5); axes[j].legend()
    fig.tight_layout()
    return fig


def fig_boxplot_fitur(tabel, y):
    data = tabel.copy()
    data["Status"] = np.where(np.asarray(y).astype(int) == 1, "Match", "Unmatch")
    n_cols = 4
    n_rows = int(np.ceil(len(tabel.columns) / n_cols))
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(16, 3.6 * n_rows))
    axes = axes.flatten()
    for i, col in enumerate(tabel.columns):
        sns.boxplot(data=data, x="Status", y=col, hue="Status", order=["Unmatch", "Match"],
                    palette={"Unmatch": "#e41a1c", "Match": "#377eb8"}, width=0.5, legend=False, ax=axes[i])
        p_val = mannwhitneyu(data.loc[data.Status == "Match", col], data.loc[data.Status == "Unmatch", col]).pvalue
        axes[i].set_title(f"{col}\np = {p_val:.4f}", fontweight="bold", fontsize=10,
                          color="darkred" if p_val < 0.05 else "black")
        axes[i].set_xlabel(""); axes[i].grid(axis="y", linestyle="--", alpha=0.5)
    for j in range(len(tabel.columns), len(axes)):
        fig.delaxes(axes[j])
    fig.tight_layout()
    return fig


def fig_klasifikasi(y, pred):
    fig, axes = plt.subplots(1, 2, figsize=(13, 4.2))
    sns.heatmap(confusion_matrix(y, pred, labels=[0, 1]), annot=True, fmt="d", cmap="Blues", cbar=False,
                xticklabels=["Unmatch", "Match"], yticklabels=["Unmatch", "Match"], ax=axes[0])
    axes[0].set_xlabel("Prediksi"); axes[0].set_ylabel("Aktual"); axes[0].set_title("Confusion Matrix")
    idx = np.arange(len(y))
    axes[1].scatter(idx, y, color="royalblue", label="Aktual", s=50)
    axes[1].scatter(idx, pred, color="red", marker="x", label="Prediksi", s=50)
    axes[1].set_yticks([0, 1], ["Unmatch", "Match"]); axes[1].set_xlabel("Index Sampel")
    axes[1].set_title("Aktual vs Prediksi"); axes[1].legend()
    fig.tight_layout()
    return fig


def fig_swatch(kolom, label_baris=None, judul=""):
    rgb = np.stack([lab_to_rgb(v) for v in kolom.values()], axis=1)
    n = rgb.shape[0]
    fig, ax = plt.subplots(figsize=(2 * len(kolom) + 2, max(2.2, 0.35 * n)))
    ax.imshow(rgb, aspect="auto")
    for x in np.arange(0.5, len(kolom) - 1, 1):
        ax.axvline(x, color="white", linewidth=3)
    for yy in np.arange(0.5, n - 1, 1):
        ax.axhline(yy, color="white", linewidth=0.8)
    ax.set_xticks(range(len(kolom)), list(kolom.keys()), fontsize=10, fontweight="bold")
    ax.set_yticks(range(n), label_baris if label_baris is not None else [f"Sampel-{i+1}" for i in range(n)])
    ax.tick_params(length=0)
    ax.set_title(judul, fontweight="bold", pad=12)
    fig.tight_layout()
    return fig


def fig_undertone(hasil, titik_baru=None):
    a, b, y, g = hasil["a"], hasil["b"], hasil["y"], hasil["garis"]
    fig, axes = plt.subplots(1, 3, figsize=(17, 6.2))
    a_max, b_max = a.max() * 1.1, b.max() * 1.1
    if titik_baru is not None:
        a_max, b_max = max(a_max, titik_baru[0] * 1.1), max(b_max, titik_baru[1] * 1.1)
    r = np.hypot(a_max, b_max) * 1.2
    judul = [f"[1] Hue manual {hasil['batas_manual']}° | Akurasi {hasil['acc_manual']*100:.1f}%",
             f"[2] Hue otomatis {np.round(hasil['batas_hue'], 1).tolist()}° | Akurasi {hasil['acc_hue']*100:.1f}%",
             f"[3] skor = {teks_skor(g['w_a'], g['w_b'])} (θ={g['theta']}°)\nAkurasi {g['acc']*100:.1f}%"]
    for i, ax in enumerate(axes):
        for ut in UNDERTONES:
            m = y == ut
            ax.scatter(a[m], b[m], c=WARNA_UNDERTONE[ut], label=ut, edgecolors="white", s=70, alpha=0.85)
        if i < 2:
            for s in (hasil["batas_manual"] if i == 0 else hasil["batas_hue"]):
                ax.plot([0, r * np.cos(np.radians(s))], [0, r * np.sin(np.radians(s))], "k--", lw=1.8,
                        label=f"b = {np.tan(np.radians(s)):.3f}·a ({s:.1f}°)")
        else:
            for t in g["batas"]:
                if abs(g["w_b"]) > 1e-9:
                    av = np.linspace(0, a_max, 100)
                    c = t / g["w_b"]
                    ax.plot(av, (t - g["w_a"] * av) / g["w_b"], "k--", lw=1.8,
                            label=f"b = {-g['w_a'] / g['w_b']:.3f}·a {'+' if c >= 0 else '−'} {abs(c):.2f}")
                else:
                    ax.axvline(t / g["w_a"], color="k", ls="--", lw=1.8, label=f"a = {t / g['w_a']:.2f}")
        if titik_baru is not None:
            ax.scatter(*titik_baru, marker="*", s=350, c="gold", edgecolors="black", zorder=6, label="Titik dicek")
        ax.set_xlim(0, a_max); ax.set_ylim(0, b_max); ax.set_aspect("equal", adjustable="box")
        ax.set_xlabel(f"a {hasil['sumber']}"); ax.set_ylabel(f"b {hasil['sumber']}")
        ax.set_title(judul[i], fontsize=10); ax.grid(True, linestyle=":", alpha=0.6)
        ax.legend(fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=2)
    fig.tight_layout()
    return fig


def warna_level(semua_level):
    return {lv: plt.colormaps["tab10"](i % 10) for i, lv in enumerate(sorted(semua_level))}


def gambar_level(ax, hasil, warna, semua_level, judul, titik_baru=None):
    x, level, tree, batas = hasil["x"], hasil["level"], hasil["tree"], hasil["batas"]
    ax.scatter(x, level, c=[warna[lv] for lv in level], edgecolors="white", s=70, zorder=4)
    lo, hi = x.min() - 1, x.max() + 1
    if titik_baru is not None:
        lo, hi = min(lo, titik_baru - 1), max(hi, titik_baru + 1)
    tepi = [lo] + list(batas) + [hi]
    area = []
    for k in range(len(tepi) - 1):
        lv = float(tree.predict([[(tepi[k] + tepi[k + 1]) / 2]])[0])
        area.append(lv)
        ax.axvspan(tepi[k], tepi[k + 1], color=warna[lv], alpha=0.15, zorder=1)
    for t in batas:
        ax.axvline(t, color="black", ls="--", lw=1.3, zorder=5)
    if titik_baru is not None:
        ax.axvline(titik_baru, color="gold", lw=3, zorder=6)
    ax.set_xlim(lo, hi)
    ax.set_yticks(sorted(semua_level))
    ax.set_xlabel(hasil["kolom"]); ax.set_ylabel("Level L")
    ax.set_title(f"{judul}\nAkurasi {hasil['acc']*100:.1f}% (N={hasil['n']})", fontweight="bold", fontsize=11)
    ax.grid(True, linestyle=":", alpha=0.6, zorder=2)
    ax.legend(handles=[mpatches.Patch(color=warna[lv], alpha=0.35, label=f"Area Lvl {lv:g}")
                       for lv in sorted(set(area))],
              loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=3, fontsize=8)
