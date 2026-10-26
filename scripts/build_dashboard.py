"""
Gera a aba "Dashboard" do Controle_Recebimentos.xlsx a partir dos dados
já lançados nas abas "Recebimentos" e "Cadastro de Clientes".

Reconstrução do script original usado no projeto (o arquivo de origem foi
perdido numa troca de sessão — este é o mesmo resultado, reconstruído a
partir da estrutura final do workbook: fórmulas, cores, gráficos e
layout de impressão idênticos ao que está publicado).

Uso:
    pip install openpyxl
    python scripts/build_dashboard.py Controle_Recebimentos.xlsx

O script é idempotente: se a aba "Dashboard" já existir, ela é recriada
do zero antes de ser preenchida de novo.
"""

import sys

from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.drawing.fill import GradientStop, LinearShadeProperties, GradientFillProperties
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.data_source import StrRef, NumRef
from openpyxl.chart.series import SeriesLabel
from openpyxl.drawing.colors import ColorChoice
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# --- Paleta (mesma linha de cor do print de referência: fundo roxo/vinho
# escuro com gradientes pink / laranja / dourado / roxo / vermelho) ---
BG = "241527"          # fundo da página
CARD_BG = "2E1833"     # fundo dos cartões / área dos gráficos
BORDER = "3B1E3F"
TEXT_LIGHT = "F5EAF0"
TEXT_MUTED = "C9A8C9"

ROXO = "9B5DE5"
PINK = "E5336B"
LARANJA = "FF9A2E"
DOURADO = "FFC93C"
VERMELHO = "E5383B"

MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun"]
MESES_EXTENSO = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho"]

REGIOES_ORDEM = [
    "Norte/PA", "Sudeste/SP", "Nordeste/PE", "Centro-Oeste/GO", "Sul/RS",
    "Nordeste/CE", "Centro-Oeste/MT", "Sul/PR", "Norte/AM", "Não identificado",
]
CATEGORIAS_ORDEM = [
    "Hortifruti", "Bebidas", "Mercearia", "Congelados",
    "Higiene e Limpeza", "Não identificado",
]


def gradient_fill(c1: str, c2: str, angle: int = 5400000) -> GraphicalProperties:
    """Gradiente linear (vertical por padrão) usado no preenchimento das barras."""
    stops = [
        GradientStop(pos=0, srgbClr=ColorChoice(srgbClr=c1)),
        GradientStop(pos=100000, srgbClr=ColorChoice(srgbClr=c2)),
    ]
    grad = GradientFillProperties(gsLst=stops, lin=LinearShadeProperties(ang=angle, scaled=True))
    gp = GraphicalProperties()
    gp.gradFill = grad
    return gp


def build_dashboard(path: str) -> None:
    wb = load_workbook(path)

    if "Recebimentos" not in wb.sheetnames or "Cadastro de Clientes" not in wb.sheetnames:
        raise SystemExit('O workbook precisa das abas "Recebimentos" e "Cadastro de Clientes".')

    rec = wb["Recebimentos"]
    last_row = rec.max_row  # linha 4 é o cabeçalho; dados começam na 5

    cad = wb["Cadastro de Clientes"]
    clientes = [row[0].value for row in cad.iter_rows(min_row=5, max_col=1) if row[0].value]

    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    ws = wb.create_sheet("Dashboard", 0)

    # ---------- layout base ----------
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_area = "A1:T50"

    bg_fill = PatternFill("solid", fgColor=BG)
    for row in ws.iter_rows(min_row=1, max_row=50, max_col=20):
        for cell in row:
            cell.fill = bg_fill

    ws.column_dimensions["A"].width = 2

    # ---------- título ----------
    ws["B1"] = "Distribuidora Aurora Ltda — Dashboard Financeiro"
    ws["B1"].font = Font(name="Calibri", size=20, bold=True, color=TEXT_LIGHT)
    ws["B2"] = "1º Semestre de 2026 (Janeiro a Junho)  ·  Fonte: aba Recebimentos"
    ws["B2"].font = Font(name="Calibri", size=11, color=TEXT_MUTED)

    # ---------- cartões de KPI ----------
    cards = [
        ("B", "TOTAL RECEBIDO", PINK, "=Dashboard!$AC$81", "jan a jun / 2026"),
        ("F", "Nº DE RECEBIMENTOS", LARANJA, "=Dashboard!$AC$82", "pagamentos no semestre"),
        ("J", "TICKET MÉDIO", DOURADO, "=Dashboard!$AC$83",
         '="por recebimento ("&Dashboard!$AC$82&" pagtos)"'),
        ("N", "MELHOR MÊS", ROXO, "=Dashboard!$AD$84", "=Dashboard!$AE$84"),
        ("R", "MAIOR CLIENTE", VERMELHO, "=Dashboard!$AD$85", "=Dashboard!$AE$85"),
    ]
    left_border = Border(left=Side(style="thin", color=BORDER))
    for col, label, accent, value_formula, detail_formula in cards:
        ws[f"{col}5"] = label
        ws[f"{col}5"].font = Font(name="Calibri", size=9, bold=True, color=accent)
        ws[f"{col}5"].border = left_border
        ws[f"{col}6"] = value_formula
        ws[f"{col}6"].font = Font(name="Calibri", size=16, bold=True, color=TEXT_LIGHT)
        ws[f"{col}6"].border = left_border
        ws[f"{col}8"] = detail_formula
        ws[f"{col}8"].font = Font(name="Calibri", size=9, color=TEXT_MUTED)
        ws[f"{col}8"].border = left_border

    # ---------- tabelas de apoio (fora da área de impressão: A1:T50) ----------
    ws["AB66"] = "Tabelas de apoio dos gráficos (fora da área de impressão)"
    ws["AB66"].font = Font(name="Calibri", size=9, italic=True, color=TEXT_MUTED)

    headers = ["Mês", "Valor", "Mês (extenso)", None, "Região/UF", "Valor",
               None, "Categoria", "Valor", None, "Cliente", "Valor"]
    for offset, text in enumerate(headers):
        if text is None:
            continue
        cell = ws.cell(row=67, column=28 + offset, value=text)  # coluna 28 = AB
        cell.font = Font(name="Calibri", size=9, bold=True, color=TEXT_LIGHT)

    # Caixa por mês (AB:AD)
    for i, (mes, extenso) in enumerate(zip(MESES, MESES_EXTENSO)):
        r = 68 + i
        ws[f"AB{r}"] = mes
        ws[f"AC{r}"] = (
            f"=SUMPRODUCT((MONTH(Recebimentos!$A$5:$A${last_row})={i + 1})"
            f"*Recebimentos!$C$5:$C${last_row})"
        )
        ws[f"AD{r}"] = extenso

    # Caixa por região (AF:AG)
    for i, regiao in enumerate(REGIOES_ORDEM):
        r = 68 + i
        ws[f"AF{r}"] = regiao
        crit = "" if regiao == "Não identificado" else regiao
        ws[f"AG{r}"] = (
            f'=SUMIF(Recebimentos!$E$5:$E${last_row},"{crit}",Recebimentos!$C$5:$C${last_row})'
        )

    # Caixa por categoria (AI:AJ)
    for i, categoria in enumerate(CATEGORIAS_ORDEM):
        r = 68 + i
        ws[f"AI{r}"] = categoria
        crit = "" if categoria == "Não identificado" else categoria
        ws[f"AJ{r}"] = (
            f'=SUMIF(Recebimentos!$F$5:$F${last_row},"{crit}",Recebimentos!$C$5:$C${last_row})'
        )

    # Ranking de clientes (AL:AM) — só clientes cadastrados; quem não está no
    # Cadastro de Clientes entra no total mas não aparece neste ranking.
    for i, cliente in enumerate(clientes):
        r = 68 + i
        ws[f"AL{r}"] = cliente
        ws[f"AM{r}"] = (
            f'=SUMIF(Recebimentos!$B$5:$B${last_row},"{cliente}",Recebimentos!$C$5:$C${last_row})'
        )
    last_client_row = 68 + len(clientes) - 1

    # Resumo auxiliar dos cartões
    ws["AB80"] = "Resumo (cálculo auxiliar dos cartões)"
    ws["AB80"].font = Font(name="Calibri", size=9, italic=True, color=TEXT_MUTED)
    ws["AB81"] = "Total Recebido"
    ws["AC81"] = f"=SUM(Recebimentos!$C$5:$C${last_row})"
    ws["AB82"] = "Nº de Recebimentos"
    ws["AC82"] = f"=COUNT(Recebimentos!$C$5:$C${last_row})"
    ws["AB83"] = "Ticket Médio"
    ws["AC83"] = "=$AC$81/$AC$82"
    ws["AB84"] = "Melhor Mês (nome / valor)"
    ws["AD84"] = "=INDEX($AD$68:$AD$73,MATCH(MAX($AC$68:$AC$73),$AC$68:$AC$73,0))"
    ws["AE84"] = "=MAX($AC$68:$AC$73)"
    ws["AB85"] = "Maior Cliente (nome / valor)"
    ws["AD85"] = f"=INDEX($AL$68:$AL${last_client_row},MATCH(MAX($AM$68:$AM${last_client_row}),$AM$68:$AM${last_client_row},0))"
    ws["AE85"] = f"=MAX($AM$68:$AM${last_client_row})"

    # ---------- gráficos ----------
    card_fill = PatternFill("solid", fgColor=CARD_BG)
    for row in ws.iter_rows(min_row=10, max_row=48, min_col=2, max_col=19):
        for cell in row:
            cell.fill = card_fill

    def style_axes(chart):
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        for axis in (chart.x_axis, chart.y_axis):
            axis.spPr = GraphicalProperties()
            axis.spPr.line = axis.spPr.line
            axis.txPr = None

    # 1) Caixa por mês — linha simples (sem área), ponto dourado, traço laranja
    line = LineChart()
    line.title = "Caixa por Mês (Jan–Jun)"
    line.style = None
    line.y_axis.majorGridlines = None
    cats = Reference(ws, min_col=28, min_row=68, max_row=73)        # AB68:AB73
    vals = Reference(ws, min_col=29, min_row=68, max_row=73)        # AC68:AC73
    line.add_data(vals, titles_from_data=False)
    line.set_categories(cats)
    s = line.series[0]
    s.smooth = True
    s.marker = Marker(symbol="circle", size=6)
    s.marker.graphicalProperties.solidFill = DOURADO
    s.graphicalProperties.line.solidFill = LARANJA
    s.graphicalProperties.line.width = 22000
    line.x_axis.axPos = "b"
    line.y_axis.axPos = "l"
    line.width, line.height = 14.8, 8.2
    ws.add_chart(line, "B11")

    # 2) Caixa por região — barras com gradiente roxo → pink
    regiao_chart = BarChart()
    regiao_chart.type = "col"
    regiao_chart.title = "Caixa por Região"
    regiao_chart.gapWidth = 30
    cats = Reference(ws, min_col=32, min_row=68, max_row=77)        # AF68:AF77
    vals = Reference(ws, min_col=33, min_row=68, max_row=77)        # AG68:AG77
    regiao_chart.add_data(vals, titles_from_data=False)
    regiao_chart.set_categories(cats)
    regiao_chart.series[0].graphicalProperties = gradient_fill(ROXO, PINK)
    regiao_chart.width, regiao_chart.height = 14.8, 8.2
    ws.add_chart(regiao_chart, "K11")

    # 3) Caixa por categoria — barras com gradiente dourado → vermelho
    categoria_chart = BarChart()
    categoria_chart.type = "col"
    categoria_chart.title = "Caixa por Categoria"
    categoria_chart.gapWidth = 40
    cats = Reference(ws, min_col=35, min_row=68, max_row=73)        # AI68:AI73
    vals = Reference(ws, min_col=36, min_row=68, max_row=73)        # AJ68:AJ73
    categoria_chart.add_data(vals, titles_from_data=False)
    categoria_chart.set_categories(cats)
    categoria_chart.series[0].graphicalProperties = gradient_fill(DOURADO, VERMELHO)
    categoria_chart.width, categoria_chart.height = 14.8, 8.2
    ws.add_chart(categoria_chart, "B31")

    # 4) Ranking de clientes — barras horizontais, maior → menor, gradiente
    #    laranja → pink (eixo explicitamente configurado para não embaralhar
    #    a ordem das categorias, que é o bug clássico desse tipo de gráfico
    #    no openpyxl)
    ranking_chart = BarChart()
    ranking_chart.type = "bar"
    ranking_chart.title = "Ranking de Clientes (maior → menor)"
    ranking_chart.gapWidth = 30
    cats = Reference(ws, min_col=38, min_row=68, max_row=last_client_row)   # AL68:AL..
    vals = Reference(ws, min_col=39, min_row=68, max_row=last_client_row)   # AM68:AM..
    ranking_chart.add_data(vals, titles_from_data=False)
    ranking_chart.set_categories(cats)
    ranking_chart.series[0].graphicalProperties = gradient_fill(LARANJA, PINK, angle=0)
    ranking_chart.x_axis.axPos = "l"   # categorias na esquerda
    ranking_chart.y_axis.axPos = "b"   # valores embaixo
    ranking_chart.y_axis.crosses = "min"
    ranking_chart.x_axis.delete = False
    ranking_chart.y_axis.delete = False
    ranking_chart.width, ranking_chart.height = 14.8, 8.2
    ws.add_chart(ranking_chart, "K31")

    wb.save(path)
    print(f"Aba Dashboard recriada em {path}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("uso: python scripts/build_dashboard.py Controle_Recebimentos.xlsx")
        raise SystemExit(1)
    build_dashboard(sys.argv[1])
