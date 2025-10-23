
import re
import pandas as pd

def _find_columns(cols):
    res = {}
    for c in cols:
        cl = c.lower()
        if "raison" in cl and "sociale" in cl:
            res["raison"] = c
        elif "catég" in cl or "categorie" in cl or "cat\u00e9g" in cl:
            res["categorie"] = c
        elif ("référent" in cl and "moa" in cl) or ("referent" in cl and "moa" in cl):
            res["referent"] = c
        elif ("email" in cl and "referent" in cl) or ("email" in cl and "référent" in cl):
            res["email_referent"] = c
        elif "contacts" in cl:
            res["contacts"] = c
    return res

def _derive_contact_moa(row, colmap):
    email = None
    if "email_referent" in colmap:
        v = row.get(colmap["email_referent"], "")
        if isinstance(v, str) and "@" in v:
            email = v.strip()
    if (not email) and "contacts" in colmap:
        raw = str(row.get(colmap["contacts"], ""))
        emails = re.split(r"[,\s;]+", raw)
        emails = [e.strip().rstrip(".,;") for e in emails if "@" in e]
        name = str(row.get(colmap.get("referent", ""), "")).strip()
        tokens = [t for t in re.split(r"[\s\-]+", name.lower()) if t]
        best = None
        for e in emails:
            local = e.split("@",1)[0].lower()
            score = sum(tok in local for tok in tokens if len(tok) >= 2)
            if best is None or score > best[0]:
                best = (score, e)
        if best and best[0] > 0:
            email = best[1]
        elif emails:
            email = emails[0]
    return email or ""

def process_csv_to_moa_df(csv_bytes_or_path):
    """Return a dataframe with columns: Raison sociale, Référent MOA, Contact MOA, Catégories (single cell)."""
    df = pd.read_csv(csv_bytes_or_path, sep=None, engine="python")
    colmap = _find_columns(df.columns)
    if "raison" not in colmap:
        df["Raison sociale"] = None
        colmap["raison"] = "Raison sociale"
    if "categorie" not in colmap:
        df["Catégories"] = None
        colmap["categorie"] = "Catégories"
    if "referent" not in colmap:
        df["Référent MOA"] = ""
        colmap["referent"] = "Référent MOA"
    if "email_referent" not in colmap and "contacts" not in colmap:
        df["Contacts"] = ""
        colmap["contacts"] = "Contacts"
    out = pd.DataFrame()
    out["Raison sociale"] = df[colmap["raison"]]
    out["Référent MOA"] = df[colmap["referent"]]
    out["Contact MOA"] = df.apply(lambda r: _derive_contact_moa(r, colmap), axis=1)
    # keep categories as a single cell (trim spaces, but no split)
    out["Catégories"] = df[colmap["categorie"]].apply(lambda x: str(x).strip() if pd.notna(x) else "")
    return out

def export_moa_excel(df, out_path_or_buffer):
    with pd.ExcelWriter(out_path_or_buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="MOA")
        ws = writer.sheets["MOA"]
        for idx, col in enumerate(df.columns):
            max_len = max([len(str(x)) for x in df[col].astype(str).values] + [len(col)])
            ws.set_column(idx, idx, min(60, max(12, max_len + 2)))
