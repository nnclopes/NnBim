# -*- coding: utf-8 -*-
__title__ = "Gerar Materiais\n(Interativo)"
__doc__ = "Pede prefixo, paleta, quantidade de tons (N), Cor A, Cor B e percentual de branco para um material cinza. Cria N materiais do degradê + 1 cinza (Shaded)."

from Autodesk.Revit.DB import (FilteredElementCollector, Material, Color, Transaction)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document

# ----------------- helpers -----------------
def sanitize(name):
    """Substitui caracteres problemáticos para nomes de material/Revit ribbon."""
    bad = ['%', ':', '\\', '/', '|', '<', '>', '"']
    for b in bad:
        name = name.replace(b, '-')
    name = name.strip()
    return name

def parse_rgb(s):
    if s is None: raise ValueError("RGB vazio")
    s = s.strip()
    # HEX: #RRGGBB
    if s.startswith("#") and len(s)==7:
        h = s[1:].upper()
        try:
            r = int(h[0:2],16); g = int(h[2:4],16); b = int(h[4:6],16)
            return (r,g,b)
        except:
            raise ValueError("HEX inválido. Use #RRGGBB.")
    # "R,G,B"
    parts = s.replace(" ","").split(",")
    if len(parts)==3:
        try:
            r = int(parts[0]); g = int(parts[1]); b = int(parts[2])
        except:
            raise ValueError("RGB inválido. Use 255,114,16.")
        for v in (r,g,b):
            if v<0 or v>255: raise ValueError("Valores RGB devem estar entre 0 e 255.")
        return (r,g,b)
    raise ValueError("Formato inválido. Use 255,114,16 ou #FF7210.")

def ensure_material(name):
    for m in FilteredElementCollector(doc).OfClass(Material):
        if m.Name == name:
            return m
    mid = Material.Create(doc, name)
    return doc.GetElement(mid)

def lerp(a,b,t): 
    return int(round(a + (b-a)*t))

# ----------------- inputs -----------------
prefix = forms.ask_for_string(default="ARQ_Brise", prompt="Prefixo dos materiais (ex.: ARQ_Brise)")
if prefix is None: raise SystemExit
prefix = sanitize(prefix)

palette = forms.ask_for_string(default="Custom", prompt="Nome da paleta (ex.: Moss, Pastel, ProjetoX)")
if palette is None: raise SystemExit
palette = sanitize(palette)

n_str = forms.ask_for_string(default="21", prompt="Quantidade de tons (N) — mínimo 2, máximo 128")
if n_str is None: raise SystemExit
try:
    N = int(n_str)
    if N < 2 or N > 128: raise ValueError
except:
    forms.alert("Quantidade de tons inválida. Informe um inteiro entre 2 e 128.", exitscript=True)

rgb_a = forms.ask_for_string(default="85,107,47", prompt="Cor A (início do degradê) — RGB (R,G,B) ou #HEX")
rgb_b = forms.ask_for_string(default="205,236,203", prompt="Cor B (fim do degradê) — RGB (R,G,B) ou #HEX")
if rgb_a is None or rgb_b is None: raise SystemExit

pct = forms.ask_for_string(default="70", prompt="Percentual de branco para o CINZA (0 a 100) — 0=preto, 100=branco")
if pct is None: raise SystemExit

try:
    cA = parse_rgb(rgb_a)
    cB = parse_rgb(rgb_b)
    P = float(pct)
    if P < 0 or P > 100: raise ValueError
    gval = int(round(255.0*(P/100.0)))
    cG = (gval,gval,gval)
except Exception as ex:
    forms.alert("Entrada inválida: %s" % ex, exitscript=True)

# zero padding
pad = 2 if N <= 99 else 3

# ----------------- create materials -----------------
t = Transaction(doc, "Nn | Materiais (Interativo)")
t.Start()

# Gradient A->B
for i in range(N):
    tv = 0.0 if N==1 else float(i)/(N-1)
    r = lerp(cA[0], cB[0], tv)
    g = lerp(cA[1], cB[1], tv)
    b = lerp(cA[2], cB[2], tv)
    name = "%s_%s_%s" % (prefix, palette, str(i+1).zfill(pad))
    m = ensure_material(name)
    m.Color = Color(r,g,b)
    try:
        m.UseRenderAppearanceForShading = False
    except:
        pass

# Gray material from % white
gray_name = "%s_%s_Cinza_g%s" % (prefix, palette, str(gval).zfill(3))
mg = ensure_material(gray_name)
mg.Color = Color(cG[0], cG[1], cG[2])
try:
    mg.UseRenderAppearanceForShading = False
except:
    pass

t.Commit()

# Summary
msg = "Criados/atualizados %d materiais do degradê %s→%s\nNome base: %s_%s_XX (padding %d)\nCinza: %s (%%branco=%.1f%%, RGB %d,%d,%d)" % (
    N, str(cA), str(cB), prefix, palette, pad, gray_name, P, cG[0], cG[1], cG[2]
)
TaskDialog.Show("Nn | Gerar Materiais (Interativo)", msg)
