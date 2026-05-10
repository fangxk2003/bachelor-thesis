#!/usr/bin/env python3

from __future__ import annotations

import argparse
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence
import xml.etree.ElementTree as ET

import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
from scipy import stats


@dataclass(frozen=True)
class TTestResult:
    column: str
    n_acl: int
    n_mlki: int
    mean_acl: float
    mean_mlki: float
    t_stat: float
    p_value: float


@dataclass(frozen=True)
class RobustGroupResult:
    column: str
    n_acl: int
    n_mlki: int
    mean_acl: float
    mean_mlki: float
    median_acl: float
    median_mlki: float
    q1_acl: float
    q3_acl: float
    q1_mlki: float
    q3_mlki: float
    zero_n_acl: int
    zero_n_mlki: int
    zero_pct_acl: float
    zero_pct_mlki: float
    welch_t_stat: float
    welch_p_value: float
    mannwhitney_u: float
    mannwhitney_p_value: float
    permutation_mean_diff: float
    permutation_p_value: float


@dataclass(frozen=True)
class AdjustedRegressionResult:
    outcome: str
    n: int
    n_acl: int
    n_mlki: int
    beta_mlki_vs_acl: float
    se_mlki_vs_acl: float
    t_stat_mlki_vs_acl: float
    p_value_mlki_vs_acl: float
    ci95_low_mlki_vs_acl: float
    ci95_high_mlki_vs_acl: float
    beta_age: float
    beta_male: float
    r_squared: float
    adjusted_significant_p05: bool


def _load_csv(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    # Normalize common missing/inf values
    df = df.replace([np.inf, -np.inf], np.nan)
    return df


def _normalize_case_id(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def _col_to_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def _xlsx_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    strings: list[str] = []
    for item in root.findall("x:si", ns):
        strings.append("".join(t.text or "" for t in item.findall(".//x:t", ns)))
    return strings


def _xlsx_cell_value(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    cell_type = cell.get("t")

    if cell_type == "inlineStr":
        return "".join(t.text or "" for t in cell.findall(".//x:t", ns))

    value = cell.find("x:v", ns)
    if value is None or value.text is None:
        return ""

    text = value.text
    if cell_type == "s":
        return shared_strings[int(text)]
    return text


def _read_xlsx_tables(path: Path) -> dict[str, pd.DataFrame]:
    ns = {
        "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    }

    tables: dict[str, pd.DataFrame] = {}
    with zipfile.ZipFile(path) as archive:
        shared_strings = _xlsx_shared_strings(archive)
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.get("Id"): rel.get("Target")
            for rel in rels.findall("rel:Relationship", ns)
            if rel.get("Id") and rel.get("Target")
        }

        for sheet in workbook.findall(".//x:sheet", ns):
            sheet_name = sheet.get("name", "sheet")
            rel_id = sheet.get(f"{{{ns['r']}}}id")
            target = rel_map.get(rel_id)
            if not target:
                continue

            sheet_path = f"xl/{target}" if not target.startswith("/") else target.lstrip("/")
            root = ET.fromstring(archive.read(sheet_path))
            rows: list[list[str]] = []
            for row in root.findall(".//x:sheetData/x:row", ns):
                values: list[str] = []
                for cell in row.findall("x:c", ns):
                    ref = cell.get("r", "")
                    if ref:
                        idx = _col_to_index(ref)
                        while len(values) <= idx:
                            values.append("")
                        values[idx] = _xlsx_cell_value(cell, shared_strings)
                    else:
                        values.append(_xlsx_cell_value(cell, shared_strings))

                while values and values[-1] == "":
                    values.pop()
                if values:
                    rows.append(values)

            if not rows:
                continue

            header = [str(c).strip() for c in rows[0]]
            width = len(header)
            data_rows = [(row + [""] * width)[:width] for row in rows[1:]]
            tables[sheet_name] = pd.DataFrame(data_rows, columns=header)

    return tables


def _first_present(columns: Sequence[str], candidates: Sequence[str]) -> str | None:
    normalized = {str(c).strip().lower(): c for c in columns}
    for candidate in candidates:
        found = normalized.get(candidate.lower())
        if found is not None:
            return str(found)
    return None


def _parse_age(value: object) -> float:
    if pd.isna(value):
        return float("nan")
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)

    text = str(value).strip()
    match = re.search(r"\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else float("nan")


def _encode_male(value: object) -> float:
    if pd.isna(value):
        return float("nan")

    text = str(value).strip().lower()
    if text in {"男", "男性", "m", "male", "1"}:
        return 1.0
    if text in {"女", "女性", "f", "female", "0"}:
        return 0.0
    return float("nan")


def load_demographics(path: Path) -> pd.DataFrame:
    case_candidates = ("case_id", "住院号", "patient_id", "subject_id", "id", "病历号")
    age_candidates = ("age", "年龄")
    sex_candidates = ("sex", "gender", "性别")

    frames: list[pd.DataFrame] = []
    for table in _read_xlsx_tables(path).values():
        case_col = _first_present(table.columns, case_candidates)
        age_col = _first_present(table.columns, age_candidates)
        sex_col = _first_present(table.columns, sex_candidates)
        if not (case_col and age_col and sex_col):
            continue

        frame = table[[case_col, age_col, sex_col]].copy()
        frame.columns = ["case_id", "age", "sex"]
        frames.append(frame)

    if not frames:
        raise ValueError(f"No demographic sheet with case_id/age/sex columns found in {path}")

    demographics = pd.concat(frames, ignore_index=True)
    demographics["case_id"] = demographics["case_id"].map(_normalize_case_id)
    demographics["age"] = demographics["age"].map(_parse_age)
    demographics["male"] = demographics["sex"].map(_encode_male)
    demographics = demographics[demographics["case_id"] != ""]
    demographics = demographics.drop_duplicates(subset=["case_id"], keep="first")
    return demographics[["case_id", "age", "sex", "male"]].reset_index(drop=True)


def _is_id_like(column: str) -> bool:
    c = column.strip().lower()
    return c == "case_id" or c.endswith("_id") or c in {"id", "subject_id", "patient_id"}


def _numeric_feature_columns(df: pd.DataFrame) -> list[str]:
    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
    return [c for c in numeric_cols if not _is_id_like(c)]


def _welch_ttest(x: pd.Series, y: pd.Series) -> tuple[float, float, int, int, float, float]:
    x_clean = pd.to_numeric(x, errors="coerce").dropna().astype(float)
    y_clean = pd.to_numeric(y, errors="coerce").dropna().astype(float)

    n1 = int(x_clean.shape[0])
    n2 = int(y_clean.shape[0])
    mean1 = float(x_clean.mean()) if n1 else float("nan")
    mean2 = float(y_clean.mean()) if n2 else float("nan")

    if n1 < 2 or n2 < 2:
        return float("nan"), float("nan"), n1, n2, mean1, mean2

    # Welch's t-test
    res = stats.ttest_ind(x_clean.to_numpy(), y_clean.to_numpy(), equal_var=False)
    return float(res.statistic), float(res.pvalue), n1, n2, mean1, mean2


def _clean_numeric(series: pd.Series) -> np.ndarray:
    return (
        pd.to_numeric(series, errors="coerce")
        .replace([np.inf, -np.inf], np.nan)
        .dropna()
        .astype(float)
        .to_numpy()
    )


def _quantile(values: np.ndarray, q: float) -> float:
    if values.size == 0:
        return float("nan")
    return float(np.quantile(values, q))


def _permutation_mean_test(
    x: np.ndarray,
    y: np.ndarray,
    *,
    n_resamples: int = 20000,
    seed: int = 20260510,
) -> tuple[float, float]:
    if x.size < 2 or y.size < 2:
        return float("nan"), float("nan")

    observed = float(y.mean() - x.mean())
    pooled = np.concatenate([x, y])
    n_x = x.size
    rng = np.random.default_rng(seed)
    count = 0

    for _ in range(n_resamples):
        permuted = rng.permutation(pooled)
        diff = float(permuted[n_x:].mean() - permuted[:n_x].mean())
        if abs(diff) >= abs(observed):
            count += 1

    # Add-one smoothing keeps p nonzero and is standard for random permutation tests.
    p_value = (count + 1) / (n_resamples + 1)
    return observed, float(p_value)


def run_robust_group_tests(acl: pd.DataFrame, mlki: pd.DataFrame) -> pd.DataFrame:
    acl_cols = set(_numeric_feature_columns(acl))
    mlki_cols = set(_numeric_feature_columns(mlki))
    cols = sorted(acl_cols.intersection(mlki_cols))

    results: list[RobustGroupResult] = []
    for col in cols:
        x = _clean_numeric(acl[col])
        y = _clean_numeric(mlki[col])
        t_stat, p_value, n1, n2, mean1, mean2 = _welch_ttest(acl[col], mlki[col])

        if x.size < 1 or y.size < 1:
            u_stat = float("nan")
            u_p = float("nan")
        else:
            u_res = stats.mannwhitneyu(x, y, alternative="two-sided", method="asymptotic")
            u_stat = float(u_res.statistic)
            u_p = float(u_res.pvalue)

        perm_diff, perm_p = _permutation_mean_test(x, y, seed=20260510 + len(results))

        results.append(
            RobustGroupResult(
                column=col,
                n_acl=n1,
                n_mlki=n2,
                mean_acl=mean1,
                mean_mlki=mean2,
                median_acl=_quantile(x, 0.50),
                median_mlki=_quantile(y, 0.50),
                q1_acl=_quantile(x, 0.25),
                q3_acl=_quantile(x, 0.75),
                q1_mlki=_quantile(y, 0.25),
                q3_mlki=_quantile(y, 0.75),
                zero_n_acl=int(np.sum(np.isclose(x, 0.0))),
                zero_n_mlki=int(np.sum(np.isclose(y, 0.0))),
                zero_pct_acl=float(np.mean(np.isclose(x, 0.0)) * 100.0) if x.size else float("nan"),
                zero_pct_mlki=float(np.mean(np.isclose(y, 0.0)) * 100.0) if y.size else float("nan"),
                welch_t_stat=t_stat,
                welch_p_value=p_value,
                mannwhitney_u=u_stat,
                mannwhitney_p_value=u_p,
                permutation_mean_diff=perm_diff,
                permutation_p_value=perm_p,
            )
        )

    out = pd.DataFrame([r.__dict__ for r in results])
    if not out.empty:
        out = out.sort_values(["mannwhitney_p_value", "column"], ascending=[True, True]).reset_index(drop=True)
    return out


def run_ttests(acl: pd.DataFrame, mlki: pd.DataFrame) -> pd.DataFrame:
    acl_cols = set(_numeric_feature_columns(acl))
    mlki_cols = set(_numeric_feature_columns(mlki))
    cols = sorted(acl_cols.intersection(mlki_cols))

    results: list[TTestResult] = []
    for col in cols:
        t_stat, p_value, n1, n2, mean1, mean2 = _welch_ttest(acl[col], mlki[col])
        results.append(
            TTestResult(
                column=col,
                n_acl=n1,
                n_mlki=n2,
                mean_acl=mean1,
                mean_mlki=mean2,
                t_stat=t_stat,
                p_value=p_value,
            )
        )

    out = pd.DataFrame([r.__dict__ for r in results])
    if not out.empty:
        out = out.sort_values(["p_value", "column"], ascending=[True, True]).reset_index(drop=True)
    return out


def _fit_ols(y: np.ndarray, x: np.ndarray) -> tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray, float]:
    beta, _, rank, _ = np.linalg.lstsq(x, y, rcond=None)
    n, p = x.shape
    fitted = np.sum(x * beta, axis=1)
    residuals = y - fitted
    df = n - p
    if rank < p or df <= 0:
        nan_arr = np.full(p, np.nan)
        return beta, nan_arr, nan_arr, nan_arr, float("nan")

    rss = float(residuals.T @ residuals)
    sigma2 = rss / df
    cov = sigma2 * np.linalg.pinv(x.T @ x)
    se = np.sqrt(np.diag(cov))
    t_stats = beta / se
    p_values = 2 * stats.t.sf(np.abs(t_stats), df=df)

    tss = float(((y - y.mean()).T @ (y - y.mean())))
    r_squared = 1.0 - rss / tss if tss > 0 else float("nan")
    return beta, se, t_stats, p_values, r_squared


def run_adjusted_regressions(acl: pd.DataFrame, mlki: pd.DataFrame, demographics: pd.DataFrame) -> pd.DataFrame:
    acl_cols = set(_numeric_feature_columns(acl))
    mlki_cols = set(_numeric_feature_columns(mlki))
    cols = sorted(acl_cols.intersection(mlki_cols))

    combined = pd.concat(
        [
            acl.assign(injury_type="ACL", injury_mlki=0.0),
            mlki.assign(injury_type="MLKI", injury_mlki=1.0),
        ],
        ignore_index=True,
    )
    combined["case_id"] = combined["case_id"].map(_normalize_case_id)
    combined = combined.merge(demographics, on="case_id", how="left", validate="many_to_one")

    results: list[AdjustedRegressionResult] = []
    for col in cols:
        model_df = combined[["injury_type", "injury_mlki", "age", "male", col]].copy()
        model_df[col] = pd.to_numeric(model_df[col], errors="coerce")
        model_df = model_df.replace([np.inf, -np.inf], np.nan).dropna()

        n = int(model_df.shape[0])
        n_acl = int((model_df["injury_type"] == "ACL").sum())
        n_mlki = int((model_df["injury_type"] == "MLKI").sum())
        if n < 5 or n_acl < 2 or n_mlki < 2:
            results.append(
                AdjustedRegressionResult(
                    outcome=col,
                    n=n,
                    n_acl=n_acl,
                    n_mlki=n_mlki,
                    beta_mlki_vs_acl=float("nan"),
                    se_mlki_vs_acl=float("nan"),
                    t_stat_mlki_vs_acl=float("nan"),
                    p_value_mlki_vs_acl=float("nan"),
                    ci95_low_mlki_vs_acl=float("nan"),
                    ci95_high_mlki_vs_acl=float("nan"),
                    beta_age=float("nan"),
                    beta_male=float("nan"),
                    r_squared=float("nan"),
                    adjusted_significant_p05=False,
                )
            )
            continue

        y = model_df[col].to_numpy(dtype=float)
        x = np.column_stack(
            [
                np.ones(n),
                model_df["injury_mlki"].to_numpy(dtype=float),
                model_df["age"].to_numpy(dtype=float),
                model_df["male"].to_numpy(dtype=float),
            ]
        )
        beta, se, t_stats, p_values, r_squared = _fit_ols(y, x)
        df_resid = n - x.shape[1]
        ci_delta = float(stats.t.ppf(0.975, df=df_resid) * se[1]) if df_resid > 0 and np.isfinite(se[1]) else float("nan")
        p_injury = float(p_values[1]) if np.isfinite(p_values[1]) else float("nan")

        results.append(
            AdjustedRegressionResult(
                outcome=col,
                n=n,
                n_acl=n_acl,
                n_mlki=n_mlki,
                beta_mlki_vs_acl=float(beta[1]),
                se_mlki_vs_acl=float(se[1]),
                t_stat_mlki_vs_acl=float(t_stats[1]),
                p_value_mlki_vs_acl=p_injury,
                ci95_low_mlki_vs_acl=float(beta[1] - ci_delta),
                ci95_high_mlki_vs_acl=float(beta[1] + ci_delta),
                beta_age=float(beta[2]),
                beta_male=float(beta[3]),
                r_squared=float(r_squared),
                adjusted_significant_p05=bool(p_injury < 0.05) if np.isfinite(p_injury) else False,
            )
        )

    out = pd.DataFrame([r.__dict__ for r in results])
    if not out.empty:
        out = out.sort_values(["p_value_mlki_vs_acl", "outcome"], ascending=[True, True]).reset_index(drop=True)
    return out


def _pick_corr_columns(df: pd.DataFrame, *, explicit: Sequence[str] | None, n: int = 3) -> list[str]:
    if explicit is not None:
        cols = list(explicit)
        if len(cols) != n:
            raise ValueError(f"Expected exactly {n} correlation columns, got {len(cols)}")
        missing = [c for c in cols if c not in df.columns]
        if missing:
            raise ValueError(f"Missing columns in CSV: {missing}")
        non_numeric = [c for c in cols if not pd.api.types.is_numeric_dtype(df[c])]
        if non_numeric:
            raise ValueError(f"Correlation columns must be numeric: {non_numeric}")
        return cols

    candidates = _numeric_feature_columns(df)
    if len(candidates) < n:
        raise ValueError(
            f"Need at least {n} numeric (non-ID) columns for correlation, found {len(candidates)}: {candidates}"
        )
    return candidates[:n]


def _save_corr_heatmap(corr: pd.DataFrame, *, title: str, out_path: Path) -> None:
    labels = list(corr.columns)
    data = corr.to_numpy(dtype=float)

    fig, ax = plt.subplots(figsize=(6, 5), dpi=160)
    im = ax.imshow(data, vmin=-1.0, vmax=1.0, cmap="coolwarm")

    ax.set_xticks(range(len(labels)), labels=labels, rotation=45, ha="right")
    ax.set_yticks(range(len(labels)), labels=labels)
    ax.set_title(title)

    for i in range(data.shape[0]):
        for j in range(data.shape[1]):
            val = data[i, j]
            if np.isfinite(val):
                ax.text(j, i, f"{val:.2f}", ha="center", va="center", fontsize=9)

    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    fig.tight_layout()
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)


def _format_p(p_value: float) -> str:
    if not np.isfinite(p_value):
        return "NA"
    if p_value < 1e-4:
        return f"{p_value:.2e}"
    return f"{p_value:.4f}"


def _save_robust_summary_plot(robust_df: pd.DataFrame, out_path: Path) -> None:
    if robust_df.empty:
        return

    preferred = ["avg_bml_mm3_per_slice", "avg_thickness_mm", "contact_area_mm2"]
    available = robust_df["column"].astype(str).tolist()
    cols = [c for c in preferred if c in available] + [c for c in available if c not in preferred]
    plot_df = robust_df.set_index("column").loc[cols].reset_index()

    x = np.arange(len(plot_df))
    width = 0.25
    p_welch = np.clip(plot_df["welch_p_value"].to_numpy(dtype=float), 1e-300, 1.0)
    p_mwu = np.clip(plot_df["mannwhitney_p_value"].to_numpy(dtype=float), 1e-300, 1.0)
    p_perm = np.clip(plot_df["permutation_p_value"].to_numpy(dtype=float), 1e-300, 1.0)

    fig, ax = plt.subplots(figsize=(10, 5), dpi=180)
    ax.bar(x - width, -np.log10(p_welch), width, label="Welch t-test")
    ax.bar(x, -np.log10(p_mwu), width, label="Mann-Whitney U")
    ax.bar(x + width, -np.log10(p_perm), width, label="Permutation")
    ax.axhline(-np.log10(0.05), color="#444444", linestyle="--", linewidth=1, label="p=0.05")
    ax.set_xticks(x, plot_df["column"].astype(str).tolist(), rotation=18, ha="right")
    ax.set_ylabel(r"$-\log_{10}(p)$")
    ax.set_title("ACL vs MLKI: Robust group-difference tests")
    ax.legend(frameon=False, ncols=2)
    ax.grid(axis="y", alpha=0.25)

    fig.tight_layout()
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)


def _save_boxplots(
    acl: pd.DataFrame,
    mlki: pd.DataFrame,
    columns: Sequence[str],
    *,
    out_path: Path,
    title: str,
) -> None:
    cols = list(columns)
    if len(cols) != 3:
        raise ValueError(f"Expected exactly 3 boxplot columns, got {len(cols)}: {cols}")

    fig, axes = plt.subplots(1, 3, figsize=(14, 5), dpi=180)
    fig.suptitle(title)

    for ax, col in zip(axes, cols, strict=True):
        acl_vals = _clean_numeric(acl[col])
        mlki_vals = _clean_numeric(mlki[col])

        ax.boxplot(
            [acl_vals, mlki_vals],
            tick_labels=["ACL", "MLKI"],
            showmeans=True,
            meanline=True,
            patch_artist=True,
            boxprops={"facecolor": "#f6f6f6", "edgecolor": "#222222"},
            medianprops={"color": "#d55e00", "linewidth": 1.8},
            meanprops={"color": "#009e73", "linewidth": 1.4, "linestyle": "--"},
        )
        ax.set_title(col)
        ax.set_ylabel("value")
        ax.grid(axis="y", alpha=0.25)

        y_min, y_max = ax.get_ylim()
        y_span = y_max - y_min
        for pos, vals in enumerate([acl_vals, mlki_vals], start=1):
            med = _quantile(vals, 0.50)
            q1 = _quantile(vals, 0.25)
            q3 = _quantile(vals, 0.75)
            zero_pct = float(np.mean(np.isclose(vals, 0.0)) * 100.0) if vals.size else float("nan")
            label = f"median [IQR]\n{med:.2f} [{q1:.2f}, {q3:.2f}]\nzero {zero_pct:.1f}%"
            ax.text(
                pos,
                y_min - 0.18 * y_span,
                label,
                ha="center",
                va="top",
                fontsize=7.5,
            )

    fig.text(0.012, 0.015, "Orange: median; green dashed: mean.", fontsize=8)
    fig.tight_layout(rect=[0, 0.16, 1, 0.92])
    fig.savefig(out_path, bbox_inches="tight")
    plt.close(fig)


def _write_df(df: pd.DataFrame, path: Path) -> None:
    df.to_csv(path, index=False)


def _parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="T-tests + correlations for ACL and MLKI CSVs")
    p.add_argument("--acl", type=Path, default=Path("acl.csv"), help="Path to acl.csv")
    p.add_argument("--mlki", type=Path, default=Path("mlki.csv"), help="Path to mlki.csv")
    p.add_argument("--demographics", type=Path, default=Path("data.xlsx"), help="Path to demographics xlsx")
    p.add_argument("--out", type=Path, default=Path("outputs"), help="Output directory")
    p.add_argument(
        "--acl-corr-cols",
        nargs="+",
        default=None,
        help="Exactly 3 numeric column names for ACL correlation (e.g. col1 col2 col3)",
    )
    p.add_argument(
        "--mlki-corr-cols",
        nargs="+",
        default=None,
        help="Exactly 3 numeric column names for MLKI correlation (e.g. col1 col2 col3)",
    )
    return p.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv)

    out_dir: Path = args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    acl = _load_csv(args.acl)
    mlki = _load_csv(args.mlki)
    demographics = load_demographics(args.demographics)

    # --- t-tests ---
    ttest_df = run_ttests(acl, mlki)
    _write_df(ttest_df, out_dir / "t_test_results.csv")

    robust_df = run_robust_group_tests(acl, mlki)
    _write_df(robust_df, out_dir / "robust_group_results.csv")
    _save_robust_summary_plot(robust_df, out_dir / "t_test_summary.png")

    # --- multivariable linear regression adjusted for age and sex ---
    adjusted_df = run_adjusted_regressions(acl, mlki, demographics)
    _write_df(adjusted_df, out_dir / "adjusted_regression_results.csv")

    # --- correlations (3 columns per CSV) ---
    acl_corr_cols = _pick_corr_columns(acl, explicit=args.acl_corr_cols, n=3)
    mlki_corr_cols = _pick_corr_columns(mlki, explicit=args.mlki_corr_cols, n=3)

    acl_corr = acl[acl_corr_cols].corr(method="pearson")
    mlki_corr = mlki[mlki_corr_cols].corr(method="pearson")

    acl_corr.to_csv(out_dir / "acl_corr_matrix.csv")
    mlki_corr.to_csv(out_dir / "mlki_corr_matrix.csv")

    _save_corr_heatmap(acl_corr, title=f"ACL correlation: {', '.join(acl_corr_cols)}", out_path=out_dir / "acl_corr.png")
    _save_corr_heatmap(
        mlki_corr, title=f"MLKI correlation: {', '.join(mlki_corr_cols)}", out_path=out_dir / "mlki_corr.png"
    )

    # --- boxplots (ACL vs MLKI) for 3 indicators ---
    # Prefer ACL-picked columns, but ensure they exist in MLKI too.
    box_cols = [c for c in acl_corr_cols if c in mlki.columns]
    if len(box_cols) != 3:
        # Fallback to shared numeric columns (non-ID) if user provided mismatched lists.
        shared = [c for c in _numeric_feature_columns(acl) if c in _numeric_feature_columns(mlki)]
        box_cols = shared[:3]
    _save_boxplots(
        acl,
        mlki,
        box_cols,
        out_path=out_dir / "boxplots_acl_vs_mlki.png",
        title="ACL vs MLKI: Boxplots for 3 indicators",
    )

    # Console summary
    tested_cols = ttest_df.shape[0]
    adjusted_cols = adjusted_df.shape[0]
    print(f"Saved outputs to: {out_dir.resolve()}")
    print(f"T-tests computed for {tested_cols} column(s): {', '.join(ttest_df['column'].tolist()) if tested_cols else 'none'}")
    print(
        "Robust tests computed for "
        f"{robust_df.shape[0]} column(s): {', '.join(robust_df['column'].tolist()) if not robust_df.empty else 'none'}"
    )
    print(
        "Adjusted regressions computed for "
        f"{adjusted_cols} column(s): {', '.join(adjusted_df['outcome'].tolist()) if adjusted_cols else 'none'}"
    )
    print(f"ACL corr columns: {acl_corr_cols}")
    print(f"MLKI corr columns: {mlki_corr_cols}")
    print(f"Boxplot columns: {box_cols}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
