import pandas as pd
import re

src = "rs1019.xlsx"
out_file = "RS1019_resultado.xlsx"

def to_float(x):
    s = str(x).strip().replace(",",".")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    return float(m.group(0)) if m else None

xls = pd.ExcelFile(src)
frames = []

for sh in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sh, header=1)
    cols = ["Tipo","Qtd","lx(cm)","ly(cm)","Peso/m²"]
    if not all(c in df.columns for c in cols):
        continue
    sub = df[cols].copy()
    sub["lx(cm)"] = sub["lx(cm)"].map(to_float)
    sub["ly(cm)"] = sub["ly(cm)"].map(to_float)
    sub["Qtd"] = sub["Qtd"].map(to_float)
    sub["Peso/m²"] = sub["Peso/m²"].map(to_float)
    sub = sub.dropna(subset=["Tipo","lx(cm)","ly(cm)","Qtd","Peso/m²"])
    sub["Origem"] = sh
    frames.append(sub)

if not frames:
    raise SystemExit("Nenhuma aba valida encontrada")

unido = pd.concat(frames, ignore_index=True)

consol = unido.groupby(["Tipo","lx(cm)","ly(cm)","Peso/m²"], as_index=False)["Qtd"].sum()
consol["A(m2)"] = (consol["lx(cm)"]/100) * (consol["ly(cm)"]/100) * consol["Qtd"]
consol["Peso(kg)"] = consol["A(m2)"] * consol["Peso/m²"]
consol = consol.sort_values(["Tipo","lx(cm)","ly(cm)"]).reset_index(drop=True)

resumo = consol.groupby("Tipo", as_index=False).agg({"A(m2)":"sum","Peso(kg)":"sum"})
resumo["Peso/m2"] = (consol.groupby("Tipo")["Peso/m²"].mean().reindex(resumo["Tipo"]).values)
resumo = resumo[["Tipo","Peso/m2","A(m2)","Peso(kg)"]].sort_values("A(m2)", ascending=False).reset_index(drop=True)

with pd.ExcelWriter(out_file) as w:
    unido.to_excel(w, sheet_name="Unido", index=False)
    consol.to_excel(w, sheet_name="Consolidado", index=False)
    resumo.to_excel(w, sheet_name="Resumo", index=False)

print("OK:", out_file)
