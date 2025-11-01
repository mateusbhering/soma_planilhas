import pandas as pd
import re
from pathlib import Path
import math
from openpyxl.styles import Font, Alignment
import sys
import os

if getattr(sys, 'frozen', False):
    script_dir = Path(sys.executable).parent
else:
    script_dir = Path(__file__).parent

src = script_dir / "planilha.xlsx"
out_file = script_dir / "resultado.xlsx"

def to_float(x):
    s = str(x).strip().replace(",",".")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    return float(m.group(0)) if m else None

def extract_tipo_number(tipo):
    m = re.search(r"(\d+)", str(tipo))
    return int(m.group(1)) if m else 0

xls = pd.ExcelFile(src)
frames = []

for sh in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sh, header=1)
    cols = ["Tipo","Qtd","lx(cm)","ly(cm)","Peso/m²"]
    if not all(c in df.columns for c in cols):
        continue
    sub = df[cols].copy()
    sub["Qtd"] = sub["Qtd"].map(to_float)
    sub["lx(cm)"] = sub["lx(cm)"].map(to_float)
    sub["ly(cm)"] = sub["ly(cm)"].map(to_float)
    sub = sub.dropna(subset=["Tipo","lx(cm)","ly(cm)","Qtd","Peso/m²"])
    sub["Origem"] = sh
    frames.append(sub)

if not frames:
    raise SystemExit("Nenhuma aba valida encontrada")

unido = pd.concat(frames, ignore_index=True)

consol = unido.groupby(["Tipo","lx(cm)","ly(cm)","Peso/m²"], as_index=False)["Qtd"].sum()
consol["A(m2)"] = (consol["lx(cm)"]/100) * (consol["ly(cm)"]/100) * consol["Qtd"]
consol["Peso(kg)"] = consol["A(m2)"] * consol["Peso/m²"]

consol["tipo_num"] = consol["Tipo"].apply(extract_tipo_number)
consol = consol.sort_values(["tipo_num","lx(cm)","ly(cm)"]).reset_index(drop=True)
consol = consol.drop("tipo_num", axis=1)

consol_export = consol[["Tipo","Qtd","lx(cm)","ly(cm)"]].copy()

resumo = consol.groupby("Tipo", as_index=False).agg({"A(m2)":"sum","Peso(kg)":"sum"})
resumo["Peso/m2"] = (consol.groupby("Tipo")["Peso/m²"].mean().reindex(resumo["Tipo"]).values)
resumo = resumo[["Tipo","Peso/m2","A(m2)","Peso(kg)"]].sort_values("Peso/m2", ascending=True).reset_index(drop=True)

resumo["A(m2)"] = resumo["A(m2)"].apply(math.ceil)
resumo["Peso(kg)"] = resumo["Peso(kg)"].apply(math.ceil)

total_row = pd.DataFrame([{
    "Tipo": "TOTAL",
    "Peso/m2": None,
    "A(m2)": None,
    "Peso(kg)": math.ceil(resumo["Peso(kg)"].sum())
}])
resumo = pd.concat([resumo, total_row], ignore_index=True)

with pd.ExcelWriter(out_file, engine='openpyxl') as w:
    unido.to_excel(w, sheet_name="Unido", index=False)
    consol_export.to_excel(w, sheet_name="Consolidado", index=False)
    resumo.to_excel(w, sheet_name="Resumo", index=False, startrow=1)
    
    worksheet = w.sheets["Resumo"]
    worksheet.merge_cells('A1:D1')
    cell = worksheet['A1']
    cell.value = "Resumo Total"
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')

print("OK:", out_file)
