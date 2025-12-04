#remember to pip install numpy pandas matplotlib scipy gradio xlsxwriter openpyxl
# Basic settings
import sys
import os, time, numpy as np, pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pathlib import Path

PREFERRED_BASENAME = "student_grades_2027-2028.xlsx"
FALLBACK_PATH = "/mnt/data/dfc93bb8-7b03-47f1-9bb4-24268f1c089e.xlsx"
OUTPUT_ROOT = "outputs"
SHEETS = ["Data", "Finance", "BM"]  # If not found, read all sheets

# For statistical tests
try:
    from scipy import stats
except Exception as e:
    raise ImportError("This script requires scipy. Please install: pip install scipy") from e


#  Utility functions
def ensure_dir(p: str) -> str:
    Path(p).mkdir(parents=True, exist_ok=True)
    return p


def ts_dir(root: str) -> str:
    ts = time.strftime("%Y%m%d_%H%M%S")
    outdir = os.path.join(root, ts)
    ensure_dir(outdir)
    ensure_dir(os.path.join(outdir, "tables"))
    ensure_dir(os.path.join(outdir, "figures"))
    return outdir


def yn_standardize(val, true_set=None, false_set=None, out_true="Y", out_false="N"):
    if pd.isna(val):
        return np.nan
    s = str(val).strip().lower()
    true_set = set(true_set or ["y", "yes", "true", "1", "t"])
    false_set = set(false_set or ["n", "no", "false", "0", "f"])
    if s in true_set:
        return out_true
    if s in false_set:
        return out_false
    return np.nan


# CI for mean, Hedges' g & Holm correction
def ci_mean(series: pd.Series, level: float = 0.95):
    s = pd.to_numeric(series, errors="coerce").dropna()
    n = len(s)
    if n <= 1:
        return np.nan, np.nan
    m = s.mean()
    sd = s.std(ddof=1)
    se = sd / np.sqrt(n)
    t_crit = stats.t.ppf(0.5 + level / 2.0, df=n - 1)
    return m - t_crit * se, m + t_crit * se


def hedges_g(x, y):
    x = pd.to_numeric(pd.Series(x), errors="coerce").dropna().to_numpy(dtype=float)
    y = pd.to_numeric(pd.Series(y), errors="coerce").dropna().to_numpy(dtype=float)
    nx, ny = len(x), len(y)
    if nx < 2 or ny < 2:
        return np.nan
    sx2, sy2 = x.var(ddof=1), y.var(ddof=1)
    sp = np.sqrt(((nx - 1) * sx2 + (ny - 1) * sy2) / (nx + ny - 2))
    if sp <= 0:
        return np.nan
    d = (x.mean() - y.mean()) / sp
    J = 1 - 3 / (4 * (nx + ny) - 9)
    return J * d


def holm_adjust(pvals):
    """Simple Holm multiple-testing correction, without statsmodels."""
    pvals = np.asarray(pvals, dtype=float)
    n = len(pvals)
    order = np.argsort(pvals)
    adj = np.empty(n, dtype=float)
    for rank, idx in enumerate(order, start=1):
        adj[idx] = min((n - rank + 1) * pvals[idx], 1.0)
    # Enforce monotonicity
    for i in range(1, n):
        adj[order[i]] = max(adj[order[i]], adj[order[i - 1]])
    return adj



#   Main analysis pipeline: run_analysis

def run_analysis(excel_path=None):
    """
    Run the full pipeline: cleaning + analysis + plotting.

    If excel_path is None, use the default search logic;
    otherwise, use the given Excel path.
    """

    #  Locate input file
    if excel_path is not None:
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Specified Excel file does not exist: {excel_path}")
        EXCEL_PATH = excel_path
    else:
        SEARCH_DIRS = [".", "/mnt/data"]
        found = None
        for d in SEARCH_DIRS:
            p = os.path.join(d, PREFERRED_BASENAME)
            if os.path.exists(p):
                found = p
                break
        if not found and os.path.exists(FALLBACK_PATH):
            found = FALLBACK_PATH
        assert found is not None, f"Could not find {PREFERRED_BASENAME}, and fallback path does not exist."
        EXCEL_PATH = found

    print(f"[info] Using Excel file: {EXCEL_PATH}")

    #  Read and merge sheets
    xls = pd.ExcelFile(EXCEL_PATH)
    sheets = [s for s in SHEETS if s in xls.sheet_names] or xls.sheet_names

    frames = []
    for s in sheets:
        df0 = pd.read_excel(EXCEL_PATH, sheet_name=s)
        df0["program"] = s
        frames.append(df0)
    df = pd.concat(frames, ignore_index=True)
    if "StudentID" in df.columns:
        df["StudentID"] = df["StudentID"].astype(str)
        df["StudentID"] = df["StudentID"].str.strip()
        df["StudentID"] = df["StudentID"].str.extract(r"(\d+)", expand=False)
        df["StudentID"] = pd.to_numeric(df["StudentID"], errors="coerce")
        df["StudentID"] = df["StudentID"].dropna().astype(int).reindex(df.index)
        df["StudentID"] = df["StudentID"].replace({None: np.nan})

# --- clean Class ---
    if "Class" in df.columns:
        df["Class"] = df["Class"].astype(str)
        df["Class"] = df["Class"].str.strip()
        df["Class"] = df["Class"].str.replace(r"\s+", "", regex=True)
        df["Class"] = df["Class"].replace("", np.nan)
    #clear class columns
    if "Class" in df.columns:
        df["Class"] = df["Class"].astype(str)
        df["Class"] = df["Class"].str.strip()
        df["Class"] = df["Class"].str.replace(r"\s+", "", regex=True)
        df["Class"] = df["Class"].replace("", np.nan)

    # Common columns
    SUBJECTS = [c for c in ["Math", "English", "Science", "History"] if c in df.columns]
    ATT_COL = next((c for c in ["Attendance (%)", "Attendance(%)", "Attendance"] if c in df.columns), None)
    PROJ_COL = next((c for c in ["ProjectScore", "Project", "Project Score"] if c in df.columns), None)
    COHORT_COL = "Cohort" if "Cohort" in df.columns else None
    TERM_COL = "Term" if "Term" in df.columns else None
    PASSED_COL = "Passed (Y/N)" if "Passed (Y/N)" in df.columns else None
    INCOME_COL = "IncomeStudent" if "IncomeStudent" in df.columns else None

    numeric_cols = SUBJECTS + ([ATT_COL] if ATT_COL else []) + ([PROJ_COL] if PROJ_COL else [])
    numeric_cols = [c for c in numeric_cols if c]  # Defensive

    bounds = {c: (0, 100) for c in SUBJECTS}
    if ATT_COL:
        bounds[ATT_COL] = (0, 100)
    if PROJ_COL:
        bounds[PROJ_COL] = (0, 100)

    #  Cleaning logic

    # 1) Cast to numeric
    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # 2) Standardize Y/N-type columns
    if PASSED_COL:
        df[PASSED_COL] = df[PASSED_COL].apply(lambda x: yn_standardize(x, out_true="Y", out_false="N"))
    if INCOME_COL:
        original_income = df[INCOME_COL].copy()
        tmp = df[INCOME_COL].apply(
            lambda x: yn_standardize(
                x,
                true_set=["y", "yes", "true", "1", "t", "income", "lowincome", "eligible"],
                false_set=["n", "no", "false", "0", "f", "non-income", "nonincome", "notincome"],
                out_true=True,
                out_false=False,
            )
        )

        mask = ~pd.isna(tmp)
        df.loc[mask, INCOME_COL] = tmp[mask]

    # 3) Term validation (only 1/2 allowed; others -> NaN so that fallback works)
    if TERM_COL:
    # Convert to string first
        df[TERM_COL] = df[TERM_COL].astype(str)

    # Remove spaces
        df[TERM_COL] = df[TERM_COL].str.strip()

    # Keep only digits (remove accidental text)
        df[TERM_COL] = df[TERM_COL].str.extract(r"(\d+)", expand=False)

    # Convert to numeric
        df[TERM_COL] = pd.to_numeric(df[TERM_COL], errors="coerce")

    # Term should be only 1 or 2
        df.loc[~df[TERM_COL].isin([1, 2]), TERM_COL] = np.nan

    # 4) Out-of-range → NaN (do not drop rows yet)
    for c, (lo, hi) in bounds.items():
        if c in df.columns:
            mask = (df[c] < lo) | (df[c] > hi)
            df.loc[mask, c] = np.nan

    # 5) Three-level median imputation:
    #    1) (Cohort, Term), 2) Cohort only, 3) global
    def median_impute_with_fallback(frame, column, cohort_col=COHORT_COL, term_col=TERM_COL):
        s = pd.to_numeric(frame[column], errors="coerce")
        # 1) (Cohort, Term)
        if cohort_col and term_col and cohort_col in frame.columns and term_col in frame.columns:
            med_ct = frame.groupby([cohort_col, term_col])[column].transform("median")
        else:
            med_ct = pd.Series(np.nan, index=frame.index)
        # 2) Cohort only
        if cohort_col and cohort_col in frame.columns:
            med_c = frame.groupby(cohort_col)[column].transform("median")
        else:
            med_c = pd.Series(np.nan, index=frame.index)
        # 3) global
        med_g = s.median()
        # Compose
        out = s.copy()
        out = out.fillna(med_ct)
        out = out.fillna(med_c)
        out = out.fillna(med_g)
        return out

    for c in numeric_cols:
        if c in df.columns:
            df[c] = median_impute_with_fallback(df, c)

    # Extra: binary pass column
    if PASSED_COL:
        df["Passed01"] = df[PASSED_COL].map({"Y": 1, "N": 0})
    else:
        df["Passed01"] = np.nan

    # Explicitly skip rows that are severely incomplete/corrupted
    key_numeric_cols = [c for c in numeric_cols if c in df.columns]
    if key_numeric_cols:
        bad_mask = df[key_numeric_cols].isna().all(axis=1)
        dropped_rows = int(bad_mask.sum())
        if dropped_rows > 0:
            df = df.loc[~bad_mask].reset_index(drop=True)
            print(f"[info] Skipped {dropped_rows} corrupted/incomplete rows (all key numeric columns NaN).")

    OUTDIR = ts_dir(OUTPUT_ROOT)
    FIGDIR = os.path.join(OUTDIR, "figures")
    TABLEDIR = os.path.join(OUTDIR, "tables")

    print(f"[info] Output directory: {OUTDIR}")

    
    # Part 1: Advanced analysis by Track / Program


    # 1.1 For each subject: descriptive stats by program (incl. skew/kurtosis & CI)
    desc_by_subject_program = None
    if SUBJECTS:
        long_rows = []
        for subj in SUBJECTS:
            sub_df = df[["program", subj]].copy()
            sub_df = sub_df.rename(columns={subj: "Score"})
            sub_df["Subject"] = subj
            long_rows.append(sub_df)
        long_df = pd.concat(long_rows, ignore_index=True)

        def _q(s, q):
            return s.quantile(q)

        g = long_df.groupby(["Subject", "program"])["Score"]
        desc_by_subject_program = g.agg(
            n="count",
            mean="mean",
            std=lambda s: s.std(ddof=1),
            p25=lambda s: _q(s, 0.25),
            median="median",
            p75=lambda s: _q(s, 0.75),
            skew=lambda s: stats.skew(s.dropna(), bias=False) if s.dropna().shape[0] >= 3 else np.nan,
            kurtosis=lambda s: stats.kurtosis(s.dropna(), bias=False, fisher=True)
            if s.dropna().shape[0] >= 4
            else np.nan,
        ).reset_index()

        ci_low, ci_high = [], []
        for _, row in desc_by_subject_program.iterrows():
            s = long_df[
                (long_df["Subject"] == row["Subject"]) & (long_df["program"] == row["program"])
            ]["Score"]
            lo, hi = ci_mean(s, 0.95)
            ci_low.append(lo)
            ci_high.append(hi)
        desc_by_subject_program["ci_low"] = ci_low
        desc_by_subject_program["ci_high"] = ci_high

    # 1.2 Math scores: ANOVA across programs + effect sizes + pairwise comparisons
    anova_math_table = None
    posthoc_math_table = None
    if "Math" in df.columns and df["program"].nunique() >= 2:
        math = pd.to_numeric(df["Math"], errors="coerce")
        groups = []
        labels = []
        for p, d_sub in df.groupby("program"):
            vals = pd.to_numeric(d_sub["Math"], errors="coerce").dropna().values
            if len(vals) >= 2:
                groups.append(vals)
                labels.append(str(p))
        if len(groups) >= 2:
            # Levene test
            lev_F, lev_p = stats.levene(*groups, center="median")
            # One-way ANOVA
            F, p = stats.f_oneway(*groups)
            k = len(groups)
            n_total = sum(len(g) for g in groups)
            df_between = k - 1
            df_within = n_total - k
            # Effect size eta^2
            grand_mean = math.dropna().mean()
            ss_between = sum(len(g) * (g.mean() - grand_mean) ** 2 for g in groups)
            ss_total = ((math.dropna() - grand_mean) ** 2).sum()
            eta2 = ss_between / ss_total if ss_total > 0 else np.nan

            anova_math_table = pd.DataFrame(
                [
                    {
                        "method": "One-way ANOVA",
                        "F": F,
                        "p": p,
                        "df_between": df_between,
                        "df_within": df_within,
                        "levene_F": lev_F,
                        "levene_p": lev_p,
                        "eta2": eta2,
                    }
                ]
            )

            # Pairwise Welch t-tests + Holm + Hedges' g
            rows = []
            for i in range(len(labels)):
                for j in range(i + 1, len(labels)):
                    a = groups[i]
                    b = groups[j]
                    t_stat, p_raw = stats.ttest_ind(a, b, equal_var=False, nan_policy="omit")
                    rows.append(
                        {
                            "group1": labels[i],
                            "group2": labels[j],
                            "t_stat": t_stat,
                            "p_raw": p_raw,
                            "g_hedges": hedges_g(a, b),
                        }
                    )
            if rows:
                posthoc_math_table = pd.DataFrame(rows)
                posthoc_math_table["p_adj"] = holm_adjust(posthoc_math_table["p_raw"].values)
                posthoc_math_table["significant"] = posthoc_math_table["p_adj"] < 0.05

    # 1.3 Attendance vs Project score: correlations overall + per program
    corr_attend_project = None
    if ATT_COL and PROJ_COL:
        rows = []
        # Overall
        x = pd.to_numeric(df[ATT_COL], errors="coerce")
        y = pd.to_numeric(df[PROJ_COL], errors="coerce")
        m = x.notna() & y.notna()
        if m.sum() >= 2:
            r, p_r = stats.pearsonr(x[m], y[m])
            rho, p_rho = stats.spearmanr(x[m], y[m])
        else:
            r = p_r = rho = p_rho = np.nan
        rows.append(
            {
                "group": "Overall",
                "n": int(m.sum()),
                "pearson_r": r,
                "pearson_p": p_r,
                "spearman_rho": rho,
                "spearman_p": p_rho,
            }
        )
        # By program
        for prog, d_sub in df.groupby("program"):
            x = pd.to_numeric(d_sub[ATT_COL], errors="coerce")
            y = pd.to_numeric(d_sub[PROJ_COL], errors="coerce")
            m = x.notna() & y.notna()
            if m.sum() >= 2:
                r, p_r = stats.pearsonr(x[m], y[m])
                rho, p_rho = stats.spearmanr(x[m], y[m])
            else:
                r = p_r = rho = p_rho = np.nan
            rows.append(
                {
                    "group": str(prog),
                    "n": int(m.sum()),
                    "pearson_r": r,
                    "pearson_p": p_r,
                    "spearman_rho": rho,
                    "spearman_p": p_rho,
                }
            )
        corr_attend_project = pd.DataFrame(rows)

    
    # Part 2: Cohort and Income analysis
    

    cohort_stats = None
    cohort_percentiles = None
    if COHORT_COL:
        num_for_cohort = [c for c in numeric_cols if c in df.columns]
        group = df.groupby(COHORT_COL)
        # Means + Count + PassRate
        agg_dict = {c: "mean" for c in num_for_cohort}
        agg_dict["StudentID"] = "count" if "StudentID" in df.columns else "size"
        cs = group.agg(agg_dict)
        cs.rename(columns={"StudentID": "Count"}, inplace=True)
        if "Passed01" in df.columns:
            pass_rate = group["Passed01"].mean()
            cs["PassRate"] = pass_rate
        cohort_stats = cs.reset_index()

        # 25/50/75 percentiles
        pct = group[num_for_cohort].quantile([0.25, 0.5, 0.75]).unstack(level=-1)

        label_map = {0.25: "p25", 0.5: "p50", 0.75: "p75"}
        pct.columns = [f"{col}_{label_map.get(q, str(q))}" for (col, q) in pct.columns]

        cohort_percentiles = pct.reset_index()
        income_stats = None
    if INCOME_COL:
        num_for_income = [c for c in numeric_cols if c in df.columns]


        group_i = df.groupby(INCOME_COL)
        agg_dict_i = {c: "mean" for c in num_for_income}
        agg_dict_i["StudentID"] = "count" if "StudentID" in df.columns else "size"
        is_df = group_i.agg(agg_dict_i)
        is_df.rename(columns={"StudentID": "Count"}, inplace=True)

        if "Passed01" in df.columns:
            is_df["PassRate"] = group_i["Passed01"].mean()

        is_df = is_df.reset_index()


        if len(is_df) == 2:
            is_df[INCOME_COL] = ["incomestudent", "nonincomestudent"]
        else:

            is_df[INCOME_COL] = is_df[INCOME_COL].astype(str)

        income_stats = is_df
    else:
        income_stats = None

    
    # Deliverable 1: Cleaned CSV
    
    clean_csv = os.path.join(TABLEDIR, "cleaned_merged.csv")
    df.to_csv(clean_csv, index=False, encoding="utf-8-sig")

    
    # Deliverable 2: Summary stats (CSV + XLSX)
    
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    desc_overall = df[num_cols].describe().T.reset_index().rename(columns={"index": "column"})

    if num_cols:
        by_program = df.groupby("program")[num_cols].agg(["count", "mean", "median", "std"])
        by_program.columns = [f"{col}_{stat}" for col, stat in by_program.columns]
        by_program = by_program.reset_index()
    else:
        by_program = pd.DataFrame({"program": df["program"].unique()})

    summary_csv = os.path.join(TABLEDIR, "summary_stats.csv")
    desc_overall.to_csv(summary_csv, index=False, encoding="utf-8-sig")

    summary_xlsx = os.path.join(TABLEDIR, "summary_stats.xlsx")
    with pd.ExcelWriter(summary_xlsx, engine="xlsxwriter") as writer:
        desc_overall.to_excel(writer, sheet_name="describe_overall", index=False)
        by_program.to_excel(writer, sheet_name="describe_by_program", index=False)
        if desc_by_subject_program is not None:
            desc_by_subject_program.to_excel(writer, sheet_name="subject_by_program", index=False)
        if anova_math_table is not None:
            anova_math_table.to_excel(writer, sheet_name="anova_math", index=False)
        if posthoc_math_table is not None:
            posthoc_math_table.to_excel(writer, sheet_name="posthoc_math", index=False)
        if corr_attend_project is not None:
            corr_attend_project.to_excel(writer, sheet_name="attend_project_corr", index=False)
        if cohort_stats is not None:
            cohort_stats.to_excel(writer, sheet_name="by_cohort", index=False)
        if cohort_percentiles is not None:
            cohort_percentiles.to_excel(writer, sheet_name="cohort_percentiles", index=False)
        if income_stats is not None:
            income_stats.to_excel(writer, sheet_name="income_comparison", index=False)

   
    # Deliverable 3: PNG figures
   
    fig_paths = []

    # Histogram of the first numeric column
    if num_cols:
        target = num_cols[0]
        plt.figure()
        df[target].dropna().plot(kind="hist", bins=20)
        plt.title(f"Histogram of {target}")
        plt.xlabel(target)
        plt.ylabel("Frequency")
        p_hist = os.path.join(FIGDIR, f"hist_{target}.png")
        plt.savefig(p_hist, dpi=160, bbox_inches="tight")
        plt.close()
        fig_paths.append(p_hist)

    # Overall: attendance vs project scatter + regression line
    ATT = next((c for c in ["Attendance (%)", "Attendance(%)", "Attendance"] if c in df.columns), None)
    PROJ = next((c for c in ["ProjectScore", "Project", "Project Score"] if c in df.columns), None)
    if ATT and PROJ and pd.api.types.is_numeric_dtype(df[ATT]) and pd.api.types.is_numeric_dtype(df[PROJ]):
        m = df[ATT].notna() & df[PROJ].notna()
        if m.sum() >= 2:
            x = df.loc[m, ATT].values
            y = df.loc[m, PROJ].values
            plt.figure()
            plt.scatter(x, y, s=12)
            A = np.vstack([x, np.ones_like(x)]).T
            beta, *_ = np.linalg.lstsq(A, y, rcond=None)
            xs = np.linspace(x.min(), x.max(), 100)
            ys = beta[0] * xs + beta[1]
            plt.plot(xs, ys)
            plt.xlabel(ATT)
            plt.ylabel(PROJ)
            plt.title(f"{ATT} vs {PROJ} (overall)")
            p_scatter = os.path.join(FIGDIR, f"scatter_{ATT}_vs_{PROJ}.png")
            plt.savefig(p_scatter, dpi=160, bbox_inches="tight")
            plt.close()
            fig_paths.append(p_scatter)

    # Per subject: overall distributions + boxplots by program
    if SUBJECTS:
        for subj in SUBJECTS:
            col = subj
            # Overall histogram
            if col in df.columns:
                s = pd.to_numeric(df[col], errors="coerce").dropna()
                if len(s) > 0:
                    plt.figure()
                    plt.hist(s, bins=20, density=False, alpha=0.7)
                    plt.xlabel(col)
                    plt.ylabel("Count")
                    plt.title(f"Distribution of {col} (overall)")
                    fp = os.path.join(FIGDIR, f"dist_{col}.png")
                    plt.savefig(fp, dpi=160, bbox_inches="tight")
                    plt.close()
                    fig_paths.append(fp)
            # Boxplot by program
            data = []
            labels = []
            for prog, d_sub in df.groupby("program"):
                vals = pd.to_numeric(d_sub[col], errors="coerce").dropna()
                if len(vals) > 0:
                    data.append(vals.values)
                    labels.append(str(prog))
            if len(data) >= 1:
                plt.figure()
                plt.boxplot(data, labels=labels, showmeans=True)
                plt.xlabel("program")
                plt.ylabel(col)
                plt.title(f"{col} by program (boxplot)")
                fp = os.path.join(FIGDIR, f"box_{col}_by_program.png")
                plt.savefig(fp, dpi=160, bbox_inches="tight")
                plt.close()
                fig_paths.append(fp)

    # Attendance vs project by program
    if ATT and PROJ:
        for prog, d_sub in df.groupby("program"):
            x = pd.to_numeric(d_sub[ATT], errors="coerce")
            y = pd.to_numeric(d_sub[PROJ], errors="coerce")
            m = x.notna() & y.notna()
            if m.sum() < 2:
                continue
            xv = x[m].values
            yv = y[m].values
            plt.figure()
            plt.scatter(xv, yv, s=12)
            A = np.vstack([xv, np.ones_like(xv)]).T
            beta, *_ = np.linalg.lstsq(A, yv, rcond=None)
            xs = np.linspace(xv.min(), xv.max(), 100)
            ys = beta[0] * xs + beta[1]
            plt.plot(xs, ys)
            r, p_r = stats.pearsonr(xv, yv)
            plt.text(0.02, 0.96, f"r={r:.3f}, p={p_r:.3g}, n={len(xv)}", transform=plt.gca().transAxes, va="top")
            plt.xlabel(ATT)
            plt.ylabel(PROJ)
            plt.title(f"{ATT} vs {PROJ} ({prog})")
            fp = os.path.join(FIGDIR, f"scatter_{ATT}_vs_{PROJ}_{prog}.png")
            plt.savefig(fp, dpi=160, bbox_inches="tight")
            plt.close()
            fig_paths.append(fp)

    # Cohort plots
    if cohort_stats is not None:
        # Attendance by Cohort
        if ATT_COL in cohort_stats.columns:
            plt.figure()
            plt.bar(cohort_stats[COHORT_COL].astype(str), cohort_stats[ATT_COL])
            plt.xlabel("Cohort")
            plt.ylabel(ATT_COL)
            plt.title("Average attendance per Cohort")
            fp = os.path.join(FIGDIR, "cohort_attendance.png")
            plt.savefig(fp, dpi=160, bbox_inches="tight")
            plt.close()
            fig_paths.append(fp)
        # Pass rate by Cohort
        if "PassRate" in cohort_stats.columns:
            plt.figure()
            plt.bar(cohort_stats[COHORT_COL].astype(str), cohort_stats["PassRate"])
            plt.xlabel("Cohort")
            plt.ylabel("PassRate")
            plt.title("Pass rate per Cohort")
            fp = os.path.join(FIGDIR, "cohort_pass_rate.png")
            plt.savefig(fp, dpi=160, bbox_inches="tight")
            plt.close()
            fig_paths.append(fp)
        # Subject means by Cohort
        for subj in SUBJECTS:
            if subj in cohort_stats.columns:
                plt.figure()
                plt.bar(cohort_stats[COHORT_COL].astype(str), cohort_stats[subj])
                plt.xlabel("Cohort")
                plt.ylabel(subj)
                plt.title(f"Average {subj} per Cohort")
                fp = os.path.join(FIGDIR, f"cohort_avg_{subj}.png")
                plt.savefig(fp, dpi=160, bbox_inches="tight")
                plt.close()
                fig_paths.append(fp)

    # Income plots
    if income_stats is not None:
        # Attendance by income group
        if ATT_COL in income_stats.columns:
            plt.figure()
            plt.bar(income_stats[INCOME_COL].astype(str), income_stats[ATT_COL])
            plt.xlabel("Income group")
            plt.ylabel(ATT_COL)
            plt.title("Attendance: Income vs Non-Income")
            fp = os.path.join(FIGDIR, "income_attendance.png")
            plt.savefig(fp, dpi=160, bbox_inches="tight")
            plt.close()
            fig_paths.append(fp)
        # Pass rate by income group
        if "PassRate" in income_stats.columns:
            plt.figure()
            plt.bar(income_stats[INCOME_COL].astype(str), income_stats["PassRate"])
            plt.xlabel("Income group")
            plt.ylabel("PassRate")
            plt.title("Pass rate: Income vs Non-Income")
            fp = os.path.join(FIGDIR, "income_pass_rate.png")
            plt.savefig(fp, dpi=160, bbox_inches="tight")
            plt.close()
            fig_paths.append(fp)
        # Subject distributions by income group
        for subj in SUBJECTS:
            if subj in df.columns:
                plt.figure()
                data = []
                labels = []
                for val, d_sub in df.groupby(INCOME_COL):
                    vals = pd.to_numeric(d_sub[subj], errors="coerce").dropna()
                    if len(vals) > 0:
                        data.append(vals.values)
                        labels.append(
                            "Income" if val is True else "Non-Income" if val is False else "Unknown"
                        )
                if data:
                    plt.boxplot(data, labels=labels, showmeans=True)
                    plt.xlabel("Income group")
                    plt.ylabel(subj)
                    plt.title(f"{subj} distribution: Income vs Non-Income")
                    fp = os.path.join(FIGDIR, f"income_box_{subj}.png")
                    plt.savefig(fp, dpi=160, bbox_inches="tight")
                    plt.close()
                    fig_paths.append(fp)

    
    # Deliverable 4: Excel with embedded figures
    
    embed_xlsx = os.path.join(TABLEDIR, "report_with_embeds.xlsx")
    with pd.ExcelWriter(embed_xlsx, engine="xlsxwriter") as writer:
        df.head(50).to_excel(writer, sheet_name="sample_data", index=False)
        desc_overall.to_excel(writer, sheet_name="describe_overall", index=False)
        by_program.to_excel(writer, sheet_name="describe_by_program", index=False)
        ws = writer.book.add_worksheet("figures")
        row = 1
        for fp in fig_paths:
            if os.path.exists(fp):
                ws.insert_image(row, 1, fp, {"x_scale": 0.7, "y_scale": 0.7})
                row += 20

    
    # methods.md (statistical methods)
    
    methods_md = os.path.join(OUTDIR, "methods.md")
    methods_text = r"""
# Methods and Formulas

**Mean**  
\(\bar{x} = \frac{1}{n}\sum_{i=1}^{n} x_i\).

**Standard Deviation**  
\(s = \sqrt{\frac{1}{n-1}\sum_{i=1}^{n}(x_i-\bar{x})^2}\).

**Skewness**  
\(\gamma_1 = \frac{\frac{1}{n}\sum(x_i-\bar{x})^3}{s^3}\).  
\(\gamma_1 > 0\): right-skewed; \(\gamma_1 < 0\): left-skewed.

**Excess Kurtosis**  
\(\gamma_2 = \frac{\frac{1}{n}\sum(x_i-\bar{x})^4}{s^4} - 3\).  
\(\gamma_2 > 0\): sharper peak and heavier tails.

**t-based Confidence Interval for the Mean**  
\([\bar{x} \pm t_{\alpha/2, n-1} \cdot s/\sqrt{n}]\).

**Pearson Correlation (Pearson r)**  
\(r = \frac{\sum (x_i-\bar{x})(y_i-\bar{y})}{\sqrt{\sum (x_i-\bar{x})^2} \sqrt{\sum (y_i-\bar{y})^2}}\).

**Spearman Correlation (Spearman ρ)**  
Pearson correlation computed on ranks; useful for non-normal data or outliers.

**Simple Linear Regression**  
\(y = \beta_0 + \beta_1 x + \varepsilon\),  
where \(x\) = attendance rate, \(y\) = project score.

**One-way ANOVA**  
\(F = \frac{\text{MS}_\mathrm{between}}{\text{MS}_\mathrm{within}}\),  
tests whether group means differ significantly.

**Levene Test for Homogeneity of Variances**  
Tests whether group variances are equal; if \(p < 0.05\), variances are considered unequal.

**Effect Size \(\eta^2\)**  
\(\eta^2 = \frac{\mathrm{SS}_\mathrm{between}}{\mathrm{SS}_\mathrm{total}}\),  
proportion of total variance explained by group differences.

**Hedges' g (Two-group Effect Size)**  
A small-sample bias-corrected version of Cohen's d,  
measuring standardized mean differences between two groups.

**Holm Multiple-testing Correction**  
Sequentially adjusts \(p\)-values to control the family-wise error rate across multiple tests.
"""
    with open(methods_md, "w", encoding="utf-8") as f:
        f.write(methods_text.strip() + "\n")

    # Return summary paths
    result = {
        "input_used": EXCEL_PATH,
        "outdir": OUTDIR,
        "cleaned_csv": clean_csv,
        "summary_csv": summary_csv,
        "summary_xlsx": summary_xlsx,
        "embedded_report_xlsx": embed_xlsx,
        "figures": fig_paths,
        "methods_md": methods_md,
    }

    return result



#   Web UI with Gradio: upload Excel, run analysis, preview results

import gradio as gr


def gradio_analyze(file):
    """
    Gradio callback: takes an uploaded Excel file, runs the analysis,
    and returns a text summary + list of figure paths + output directory.
    """
    if file is None:
        return "Please upload an Excel file (.xlsx) first.", [], ""

    # In Gradio, `file` is a TemporaryFile-like object with a `.name` path
    excel_path = file.name

    try:
        result = run_analysis(excel_path)
    except Exception as e:
        return f"[error] An error occurred during analysis: {e}", [], ""

    msg_lines = [
        f"Input file used: {result['input_used']}",
        f"Output directory: {result['outdir']}",
        "",
        "Key outputs:",
        f"- Cleaned CSV: {result['cleaned_csv']}",
        f"- Summary CSV: {result['summary_csv']}",
        f"- Summary XLSX: {result['summary_xlsx']}",
        f"- Embedded report (Excel with figures): {result['embedded_report_xlsx']}",
        f"- Methods description: {result['methods_md']}",
        "",
        "All PNG figures are shown below in the gallery."
    ]
    msg = "\n".join(msg_lines)

    fig_list = result["figures"]
    return msg, fig_list, result["outdir"]


demo = gr.Interface(
    fn=gradio_analyze,
    inputs=gr.File(label="Upload Excel file (.xlsx)", file_types=[".xlsx"]),
    outputs=[
        gr.Textbox(label="Run summary", lines=12),
        gr.Gallery(label="Figures (PNG)", columns=3, height="auto"),
        gr.Textbox(label="Output folder on disk"),
    ],
    title="Student Grades Report Generator",
    description=(
        "Upload a gradebook Excel file (.xlsx). The script will clean and merge the data, "
        "compute summary statistics, run statistical tests, and generate PNG figures "
        "and Excel reports under an 'outputs/...' folder on disk."
    ),
)


if __name__ == "__main__":
    # Launch the web UI. By default, Gradio will open a browser tab.
    demo.launch()




#   Command-line entry point

def main():
    """
    CLI entry: run the full pipeline once.

    Usage:
        python report_generator.py              # use default search logic
        python report_generator.py input.xlsx   # use a specific Excel file
    """
    import argparse

    parser = argparse.ArgumentParser(
        description="Run the student grades report pipeline."
    )
    parser.add_argument(
        "excel_path",
        nargs="?",
        default=None,
        help="Path to the input Excel file (optional). "
             "If omitted, the script will search for the preferred basename "
             "or fallback path.",
    )
    args = parser.parse_args()

    result = run_analysis(args.excel_path)
    print("\nDone.")
    print("Output directory:", result["outdir"])
    print("Cleaned CSV:", result["cleaned_csv"])
    print("Summary XLSX:", result["summary_xlsx"])
    print("Embedded report:", result["embedded_report_xlsx"])


if __name__ == "__main__":
    
    if "ipykernel" in sys.modules:
        
        demo.launch()
    else:
        
        main()


