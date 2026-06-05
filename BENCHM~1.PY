from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side, GradientFill)
from openpyxl.utils import get_column_letter

wb = Workbook()

# ── Sheet 1: Feature Matrix ──────────────────────────────────────────
ws = wb.active
ws.title = "Feature Matrix S2"

# Colors
C_GREEN_BG = "D6F4E0"
C_YELLOW_BG = "FFF3CD"
C_RED_BG = "FDDEDE"
C_BLUE_HDR = "0059D5"
C_TN_COL = "EBF2FF"
C_GRAY = "F5F7FA"
C_WHITE = "FFFFFF"
C_DARK = "000B19"

def cell_style(ws, row, col, value, bg=None, bold=False, wrap=True, font_color="000000", size=9, halign="center"):
    c = ws.cell(row=row, column=col, value=value)
    if bg:
        c.fill = PatternFill("solid", start_color=bg, end_color=bg)
    c.font = Font(name="Arial", bold=bold, color=font_color, size=size)
    c.alignment = Alignment(horizontal=halign, vertical="center", wrap_text=wrap)
    return c

def border_all(ws, min_row, max_row, min_col, max_col):
    thin = Side(style="thin", color="D0D5DD")
    for r in range(min_row, max_row+1):
        for c in range(min_col, max_col+1):
            ws.cell(r, c).border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Title row
ws.merge_cells("A1:M1")
c = ws["A1"]
c.value = "Benchmark Competitivo · Sección 2 · Landing de Value Prop para Agencias"
c.font = Font(name="Arial", bold=True, color="FFFFFF", size=12)
c.fill = PatternFill("solid", start_color=C_BLUE_HDR)
c.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[1].height = 28

# Subtitle row
ws.merge_cells("A2:M2")
c = ws["A2"]
c.value = "Análisis basado en relevamiento directo de cada landing. Datos verificados: Junio 2026"
c.font = Font(name="Arial", italic=True, color="666666", size=8)
c.fill = PatternFill("solid", start_color="F0F4FF")
c.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[2].height = 18

# Platform data: name, url, verified
platforms = [
    ("Shopify",         "shopify.com/ar/partners",           "✅ Verificado"),
    ("Wix Studio",      "wix.com/studio/partner-program",    "✅ Verificado"),
    ("Adobe Commerce",  "business.adobe.com/support/adobe-partners.html", "✅ Verificado"),
    ("WooCommerce",     "woocommerce.com/for-agencies/",     "✅ Verificado"),
    ("VTEX",            "vtex.com/es-mx/partners/",          "✅ Verificado"),
    ("Tiendanube",      "tiendanube.com/socios/comerciales", "✅ Verificado"),
]

# Header row (row 4)
ws.row_dimensions[4].height = 50
cell_style(ws, 4, 1, "Feature", bg="1E3A5F", bold=True, font_color="FFFFFF", size=9, halign="left")
cell_style(ws, 4, 2, "Peso", bg="1E3A5F", bold=True, font_color="FFFFFF", size=9)

for i, (name, url, _) in enumerate(platforms):
    col = 3 + i
    bg = C_TN_COL if name == "Tiendanube" else "1E3A5F"
    fc = "000B19" if name == "Tiendanube" else "FFFFFF"
    cell_style(ws, 4, col, f"{name}\n{url}", bg=bg, bold=True, font_color=fc, size=8)

cell_style(ws, 4, 9, "Notas", bg="1E3A5F", bold=True, font_color="FFFFFF", size=9, halign="left")

# Feature matrix data
# Format: (feature_name, peso, shopify, wix, adobe, woo, vtex, tn, nota)
# Values: "✓" / "◐" / "✗" with optional note
features = [
    ("1. Landing dedicada para agencias",
     1,
     ("✓", C_GREEN_BG, "Landing muy completa. Hero: 'Construye. Gana. Crece.'"),
     ("✓", C_GREEN_BG, "Landing dedicada Wix Studio con hero claro y beneficios"),
     ("✓", C_GREEN_BG, "'Powerful partnerships, powerful results'. 3 tipos de programa."),
     ("✓", C_GREEN_BG, "Redirige a 'Automattic for Agencies'. Landing existe pero básica."),
     ("✓", C_GREEN_BG, "Landing existe: 'ISV e Integradores de Sistemas'. Foco en tech partners."),
     ("✓", C_TN_COL, "'Cuando tus clientes crecen, tu agencia también'. Landing completa."),
     "Todos tienen landing. La diferencia está en calidad y profundidad del contenido."),

    ("2. Tiers con beneficios explicados",
     1,
     ("✓", C_GREEN_BG, "Tiers implícitos en beneficios (apps, referrals, Plus). Sin tabla explícita."),
     ("◐", C_YELLOW_BG, "Beneficios escalan con la agencia. Sin tiers nombrados explícitamente."),
     ("✓", C_GREEN_BG, "3 programas: Solution Partners, Technology Partners, Channel Partners."),
     ("✗", C_RED_BG, "Sin tiers. Solo describe beneficios generales del programa Automattic."),
     ("✗", C_RED_BG, "Sin tiers explicados. Solo CTA 'Convertirse en socio'."),
     ("◐", C_YELLOW_BG, "4 pilares descriptos pero tiers nuevos (Silver/Gold/Platinum) no publicados."),
     "Shopify y Adobe tienen estructura de tiers más clara. TN tiene los pilares pero le falta publicar los tiers del nuevo programa."),

    ("3. Requisitos de ingreso / progresión",
     1,
     ("✓", C_GREEN_BG, "Signup gratuito. Proceso claro en 3 pasos en la misma página."),
     ("✓", C_GREEN_BG, "Proceso claro: Registrarse → crear sitios → ganar beneficios."),
     ("✓", C_GREEN_BG, "CTA 'Become a partner' → partners.adobe.com/join. Proceso claro."),
     ("◐", C_YELLOW_BG, "Solo dice 'Apply today at Automattic.com'. Redirige a otro sitio."),
     ("◐", C_YELLOW_BG, "Solo CTA 'Convertirse en socio'. Sin descripción del proceso."),
     ("✓", C_TN_COL, "4 pasos explicados: Registrate → Implementá → Crecé → Clientes e ingresos."),
     "TN tiene el proceso mejor documentado entre los que sí lo tienen."),

    ("4. Revenue share / modelo de ingresos",
     1,
     ("✓", C_GREEN_BG, "Explícito: $1.3B pagados a partners en 2025. Múltiples fuentes de ingresos."),
     ("✓", C_GREEN_BG, "20%+ en sitios Premium, comisión por Wix Payments, revenue de apps."),
     ("◐", C_YELLOW_BG, "Mencionado como beneficio pero sin % específico en esta página."),
     ("◐", C_YELLOW_BG, "'Special agency pricing and exclusive referral program'. Sin datos concretos."),
     ("✗", C_RED_BG, "Sin mención de revenue share en la landing."),
     ("◐", C_YELLOW_BG, "'Comisiones recurrentes' + 'co-inversión' mencionados sin % publicados."),
     "Solo Shopify ($1.3B) y Wix (20%+) tienen el modelo de ingresos cuantificado públicamente."),

    ("5. Habilitación / academia / certificaciones",
     1,
     ("✓", C_GREEN_BG, "Shopify Academy: cursos, certifications, Partner Blog, comunidad dev."),
     ("✓", C_GREEN_BG, "'Certifica tu experiencia con certificaciones de Wix Studio.' Webinarios."),
     ("✓", C_GREEN_BG, "Adobe Learning, Experience League, certificaciones por producto."),
     ("✗", C_RED_BG, "No mencionado en la landing. Sin academia visible."),
     ("✗", C_RED_BG, "No mencionado. La página es principalmente un directorio de socios."),
     ("✓", C_TN_COL, "'Rutas de certificación' y 'recorrido estructurado desde el primer día'."),
     "TN, Shopify, Wix y Adobe tienen habilitación documentada. VTEX y Woo (for agencies) no."),

    ("6. Dashboard / herramientas para el partner",
     1,
     ("✓", C_GREEN_BG, "Partner Dashboard: gestión de clientes, earnings, dev stores gratuitas."),
     ("✓", C_GREEN_BG, "'Gestiona todos tus sitios y clientes desde una sola plataforma.'"),
     ("✓", C_GREEN_BG, "Partner portal, Adobe Exchange, herramientas de integración."),
     ("◐", C_YELLOW_BG, "Menciona plataforma de gestión de sitios. Sin detalle."),
     ("✗", C_RED_BG, "No mencionado en la landing analizada."),
     ("✓", C_TN_COL, "Portal de Partners mencionado explícitamente."),
     ""),

    ("7. Co-marketing / generación de leads",
     1,
     ("✓", C_GREEN_BG, "Shopify Partner Directory, blog de partners, eventos, comunidad."),
     ("✓", C_GREEN_BG, "Wix Marketplace: listings y lead gen activo. 'Nuevos clientes en tu bandeja.'"),
     ("✓", C_GREEN_BG, "Adobe Exchange, partner directory, joint case studies."),
     ("◐", C_YELLOW_BG, "'Partner directory listings across multiple Automattic brands.' Sin más detalle."),
     ("✗", C_RED_BG, "No mencionado en esta landing."),
     ("✓", C_TN_COL, "'Presencia en marketplace, eventos conjuntos y acciones de marketing.'"),
     ""),

    ("8. Partner manager / soporte dedicado",
     1,
     ("◐", C_YELLOW_BG, "Soporte prioritario. Partner manager solo en tiers Shopify Plus."),
     ("✓", C_GREEN_BG, "'Tu gerente de éxito de Partners' disponible. Mencionado para todos."),
     ("◐", C_YELLOW_BG, "Partner manager en Platinum. Otros tiers sin mención explícita."),
     ("✗", C_RED_BG, "Solo soporte técnico general. Sin partner manager mencionado."),
     ("✗", C_RED_BG, "No mencionado en la landing."),
     ("◐", C_YELLOW_BG, "'Soporte humano regional dedicado' solo en niveles más altos."),
     "Wix es el único que ofrece partner manager a todos los partners explícitamente."),

    ("9. Testimonios / casos de éxito de agencias",
     1,
     ("✓", C_GREEN_BG, "5 testimonios: Zapiet, Domaine, Judge.me, The Working Party, Loop."),
     ("◐", C_YELLOW_BG, "No visible en la landing principal del partner program."),
     ("✓", C_GREEN_BG, "Customer success stories en la página de partners."),
     ("✗", C_RED_BG, "Sin testimonios en la landing."),
     ("✗", C_RED_BG, "Sin testimonios en la landing."),
     ("✓", C_TN_COL, "4 testimonios: Zur, Javier Maldonado, Volando Bajito, Minimal Web."),
     ""),

    ("10. CTA claro + proceso de aplicación",
     1,
     ("✓", C_GREEN_BG, "'Conviértete en partner' → partners.shopify.com/signup. Gratuito."),
     ("✓", C_GREEN_BG, "'Registrarse ahora' → proceso directo en la misma plataforma."),
     ("✓", C_GREEN_BG, "'Become a partner' → partners.adobe.com/join."),
     ("✓", C_GREEN_BG, "'Join the program' → Automattic.com. CTA claro pero redirige a otro dominio."),
     ("✓", C_GREEN_BG, "'Convertirse en socio' → formulario. CTA presente."),
     ("✓", C_TN_COL, "'Quiero ser parte' → partners.tiendanube.com/auth/signup."),
     ""),

    ("11. Foco explícito en LATAM",
     1,
     ("✗", C_RED_BG, "Plataforma global. Sin foco regional en LATAM."),
     ("◐", C_YELLOW_BG, "Disponible en español. Sin foco específico en LATAM."),
     ("✗", C_RED_BG, "Global. Sin mención de LATAM."),
     ("✗", C_RED_BG, "Global. Sin mención de LATAM."),
     ("✓", C_GREEN_BG, "40+ países. Regiones: Brazil, South/North LATAM mencionados explícitamente."),
     ("✓", C_TN_COL, "AR, BR, CL, CO, MX. 'Ecosistema de e-commerce más grande de LATAM.'"),
     "Solo TN y VTEX tienen foco real en LATAM."),
]

# Score row values (✓=1, ◐=0.5, ✗=0)
score_map = {"✓": 1.0, "◐": 0.5, "✗": 0.0}

# Write features
row = 5
for feat_name, peso, shopify, wix, adobe, woo, vtex, tn, nota in features:
    ws.row_dimensions[row].height = 36
    cell_style(ws, row, 1, feat_name, bg=C_GRAY if row % 2 == 0 else C_WHITE, halign="left", size=8, bold=True)
    cell_style(ws, row, 2, peso, bg=C_GRAY if row % 2 == 0 else C_WHITE, size=8)

    for col_i, (val, bg, note) in enumerate([(shopify, None, None), (wix, None, None),
                                              (adobe, None, None), (woo, None, None),
                                              (vtex, None, None), (tn, None, None)]):
        v, b, n = val
        actual_bg = b
        c = ws.cell(row=row, column=3+col_i, value=v)
        c.fill = PatternFill("solid", start_color=actual_bg)
        c.font = Font(name="Arial", bold=True, size=9,
                      color="16A34A" if v=="✓" else ("D97706" if v=="◐" else "DC2626"))
        c.alignment = Alignment(horizontal="center", vertical="center")
        if n:
            from openpyxl.comments import Comment
            comment = Comment(n, "Benchmark")
            c.comment = comment

    cell_style(ws, row, 9, nota, bg=C_GRAY if row % 2 == 0 else C_WHITE, halign="left", size=8, wrap=True)
    row += 1

# Score calculation rows
ws.row_dimensions[row].height = 8  # spacer
row += 1

# Score header
cell_style(ws, row, 1, "SCORE CALCULADO (✓=1 · ◐=0.5 · ✗=0 · escala /10)", 
           bg="1E3A5F", bold=True, font_color="FFFFFF", size=9, halign="left")
for col in range(2, 10):
    ws.cell(row=row, column=col).fill = PatternFill("solid", start_color="1E3A5F")
ws.merge_cells(f"A{row}:B{row}")
ws.row_dimensions[row].height = 22
row += 1

# Scores per platform
platform_scores = []
for p_idx in range(6):
    total = 0
    for feat in features:
        vals = [feat[2], feat[3], feat[4], feat[5], feat[6], feat[7]]
        v = vals[p_idx][0]
        total += score_map.get(v, 0)
    score = round(total / 11 * 10, 1)
    platform_scores.append(score)

score_colors = ["C_GREEN_BG", C_GREEN_BG, C_GREEN_BG, C_GREEN_BG, C_YELLOW_BG, C_YELLOW_BG, C_YELLOW_BG]

ws.row_dimensions[row].height = 30
cell_style(ws, row, 1, "Puntos brutos (/11 features)", bg=C_GRAY, size=8, bold=True, halign="left")
for p_idx in range(6):
    total = sum(score_map.get(f[[2,3,4,5,6,7][p_idx]][0], 0) for f in features)
    bg = C_TN_COL if p_idx == 5 else C_GRAY
    cell_style(ws, row, 3+p_idx, round(total, 1), bg=bg, size=9, bold=True)
row += 1

ws.row_dimensions[row].height = 35
cell_style(ws, row, 1, "Score /10 (calculado)", bg="1E3A5F", bold=True, font_color="FFFFFF", size=9, halign="left")
for p_idx, (score, (name, _, _)) in enumerate(zip(platform_scores, platforms)):
    bg = C_TN_COL if p_idx == 5 else ("D6F4E0" if score >= 7 else ("FFF3CD" if score >= 5 else "FDDEDE"))
    c = ws.cell(row=row, column=3+p_idx, value=score)
    c.fill = PatternFill("solid", start_color=bg)
    c.font = Font(name="Arial", bold=True, size=12, color="0059D5" if p_idx==5 else "000000")
    c.alignment = Alignment(horizontal="center", vertical="center")
row += 1

# Benchmark vs current
ws.row_dimensions[row].height = 22
old_scores = [9.5, 8.0, 7.5, 6.5, 7.0, 3.5]
cell_style(ws, row, 1, "Score previo en benchmark (antes del relevamiento)", bg=C_GRAY, size=8, bold=False, halign="left")
for p_idx, old in enumerate(old_scores):
    bg = C_TN_COL if p_idx == 5 else C_GRAY
    cell_style(ws, row, 3+p_idx, old, bg=bg, size=9)
row += 1

ws.row_dimensions[row].height = 22
cell_style(ws, row, 1, "Δ vs. previo", bg="1E3A5F", bold=True, font_color="FFFFFF", size=9, halign="left")
for p_idx, (new, old) in enumerate(zip(platform_scores, old_scores)):
    diff = round(new - old, 1)
    bg = "D6F4E0" if diff > 0 else ("FDDEDE" if diff < 0 else C_GRAY)
    bg = C_TN_COL if p_idx == 5 else bg
    cell_style(ws, row, 3+p_idx, f"+{diff}" if diff > 0 else str(diff), bg=bg, size=9, bold=True,
              font_color="16A34A" if diff > 0 else ("DC2626" if diff < 0 else "000000"))
row += 2

# Key observations
cell_style(ws, row, 1, "🔑 OBSERVACIONES CLAVE", bg="FEF3C7", bold=True, size=9, halign="left")
ws.merge_cells(f"A{row}:I{row}")
row += 1

obs = [
    ("⚠️ VTEX", "La landing /es-mx/partners/ es para ISV e Integradores de Sistemas, NO para agencias de servicio. Tiene muy poca value prop: sin tiers, sin revenue share, sin testimonios. Score cae de 7.0 → " + str(platform_scores[4])),
    ("⚠️ WooCommerce", "woocommerce.com/for-agencies/ redirige a 'Automattic for Agencies'. Landing básica: no menciona tiers ni academia. Score baja de 6.5 → " + str(platform_scores[3])),
    ("✅ Wix Studio", "Programa muy sólido: revenue share 20%+, partner manager para todos, certificaciones, Wix Marketplace como lead gen. Score sube a " + str(platform_scores[1])),
    ("✅ Tiendanube", "tiendanube.com/socios/comerciales SÍ existe y es completa. 4 pilares, 4 pasos claros, testimonios, CTA, co-marketing, certificaciones. GAP: tiers nuevos no publicados aún. Score real: " + str(platform_scores[5])),
    ("💡 Recomendación", "Publicar los tiers Silver/Gold/Platinum/Titanium/Blue Titanium con beneficios concretos en /socios/comerciales subiría el score de " + str(platform_scores[5]) + " a ~9.0 — igualando a Shopify."),
]

for label, text in obs:
    ws.row_dimensions[row].height = 28
    cell_style(ws, row, 1, label, bg="FFFBEB", bold=True, size=8, halign="left")
    ws.cell(row=row, column=2).value = text
    ws.merge_cells(f"B{row}:I{row}")
    ws.cell(row=row, column=2).font = Font(name="Arial", size=8)
    ws.cell(row=row, column=2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.cell(row=row, column=2).fill = PatternFill("solid", start_color="FFFBEB")
    row += 1

# ── Column widths ──
ws.column_dimensions['A'].width = 32
ws.column_dimensions['B'].width = 7
for i in range(6):
    ws.column_dimensions[get_column_letter(3+i)].width = 18
ws.column_dimensions['I'].width = 40

# Borders
border_all(ws, 4, row-1, 1, 9)

# Freeze panes
ws.freeze_panes = "C5"

# ── Sheet 2: Source URLs ──────────────────────────────────────────────
ws2 = wb.create_sheet("Fuentes")
ws2.column_dimensions['A'].width = 25
ws2.column_dimensions['B'].width = 55
ws2.column_dimensions['C'].width = 15

headers = ["Plataforma", "URL relevada", "Estado"]
for col, h in enumerate(headers, 1):
    c = ws2.cell(1, col, h)
    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    c.fill = PatternFill("solid", start_color="1E3A5F")
    c.alignment = Alignment(horizontal="left", vertical="center")

sources = [
    ("Shopify", "https://www.shopify.com/ar/partners", "✅ Verificado via web_fetch"),
    ("Wix Studio", "https://es.wix.com/studio/partner-program", "✅ Verificado via Chrome"),
    ("Adobe Commerce", "https://business.adobe.com/support/adobe-partners.html", "✅ Verificado via web_fetch"),
    ("WooCommerce", "https://woocommerce.com/for-agencies/", "✅ Verificado via Chrome"),
    ("VTEX", "https://www.vtex.com/es-mx/partners/", "✅ Verificado via Chrome"),
    ("Tiendanube", "https://www.tiendanube.com/socios/comerciales", "✅ Verificado via web_fetch"),
    ("", "", ""),
    ("⚠️ NOTA", "woocommerce.com/es/partners/ (URL original) → es para ISVs de marketplace, NO para agencias", "⚠️ Revisar"),
    ("⚠️ NOTA", "vtex.com/es-mx/partners/ → foco en ISV & System Integrators, no exclusivo para agencias", "⚠️ Revisar"),
]

for r_idx, (p, url, status) in enumerate(sources, 2):
    ws2.cell(r_idx, 1, p).font = Font(name="Arial", size=9, bold=True)
    ws2.cell(r_idx, 2, url).font = Font(name="Arial", size=9, color="0059D5")
    ws2.cell(r_idx, 3, status).font = Font(name="Arial", size=9)
    for c in range(1,4):
        ws2.cell(r_idx, c).alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[r_idx].height = 18

wb.save("/sessions/eager-cool-gauss/mnt/outputs/benchmark_S2_feature_matrix.xlsx")
print("Excel saved")
