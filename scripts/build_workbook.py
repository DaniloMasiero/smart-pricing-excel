"""Builds the Smart Pricing workbook (Excel) for a fictitious Brazilian hair care brand.

Sheets: Dashboard, Settings, Products, Pricing, Simulator. Every number on Pricing, Simulator and
Dashboard is an Excel formula driven by the inputs on Settings and Products (blue cells on yellow).
Usage: python scripts/build_workbook.py   ->  SmartPricing.xlsx
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = Path(__file__).resolve().parent.parent / "SmartPricing.xlsx"

NAVY, TEAL, CORAL, AMBER = "1F3A5F", "2A9D8F", "E76F51", "E9C46A"
GREY, LIGHT, INPUT = "6B7785", "F3F5F8", "FFF7D6"
F = "Segoe UI"
MONEY, PCT = '"R$" #,##0.00', "0.0%"

title_font = Font(name=F, size=18, bold=True, color=NAVY)
sub_font = Font(name=F, size=9, color=GREY)
head_font = Font(name=F, size=10, bold=True, color="FFFFFF")
body_font = Font(name=F, size=10, color="1F2937")
input_font = Font(name=F, size=10, color="1F4FBF")      # blue = input (modeling convention)
bold = Font(name=F, size=10, bold=True, color=NAVY)
head_fill = PatternFill("solid", fgColor=NAVY)
input_fill = PatternFill("solid", fgColor=INPUT)
light_fill = PatternFill("solid", fgColor=LIGHT)
thin = Side(style="thin", color="D9DEE5")
box = Border(top=thin, bottom=thin, left=thin, right=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

CHANNELS = [  # name, commission %, fixed fee per unit, payment fee %, shipping per unit, marketing %
    ("Store", 0.00, 0.00, 0.025, 0.00, 0.02),
    ("Website", 0.00, 0.00, 0.040, 6.00, 0.05),
    ("Mercado Livre", 0.12, 6.50, 0.000, 0.00, 0.03),
    ("Amazon", 0.13, 4.50, 0.000, 0.00, 0.03),
    ("Shopee", 0.14, 4.00, 0.000, 0.00, 0.03),
]
# sku, product, category, unit cost, packaging, current prices per channel (Store, Website, ML, Amazon, Shopee)
PRODUCTS = [
    ("NT-SH01", "Repair Shampoo 300 ml", "Shampoo", 14.70, 1.80, (39.90, 39.90, 44.90, 44.90, 42.90)),
    ("NT-SH02", "Curl Shampoo 300 ml", "Shampoo", 15.60, 1.80, (42.90, 42.90, 47.90, 46.90, 45.90)),
    ("NT-SH03", "Anti-Frizz Shampoo 300 ml", "Shampoo", 14.40, 1.80, (34.90, 34.90, 36.90, 36.90, 34.90)),
    ("NT-SH04", "Detox Shampoo 1 L", "Shampoo", 33.75, 3.15, (89.90, 89.90, 99.90, 97.90, 94.90)),
    ("NT-CD01", "Repair Conditioner 300 ml", "Conditioner", 15.30, 1.80, (42.90, 42.90, 46.90, 46.90, 44.90)),
    ("NT-CD02", "Curl Conditioner 300 ml", "Conditioner", 16.35, 1.80, (44.90, 44.90, 49.90, 49.90, 47.90)),
    ("NT-CD03", "Anti-Frizz Conditioner 300 ml", "Conditioner", 15.15, 1.80, (36.90, 36.90, 38.90, 38.90, 36.90)),
    ("NT-MK01", "Hydration Mask 250 g", "Mask", 20.55, 2.70, (59.90, 59.90, 64.90, 64.90, 62.90)),
    ("NT-MK02", "Reconstruction Mask 250 g", "Mask", 25.35, 2.70, (69.90, 69.90, 74.90, 74.90, 72.90)),
    ("NT-MK03", "Nutrition Mask 500 g", "Mask", 37.20, 3.60, (99.90, 99.90, 109.90, 107.90, 104.90)),
    ("NT-MK04", "Curl Mask 1 kg", "Mask", 57.75, 4.65, (139.90, 139.90, 149.90, 149.90, 144.90)),
    ("NT-LV01", "Leave-in Cream 200 ml", "Leave-in", 13.35, 1.65, (34.90, 34.90, 37.90, 37.90, 35.90)),
    ("NT-LV02", "Thermal Protector 150 ml", "Leave-in", 16.80, 2.10, (49.90, 49.90, 52.90, 52.90, 49.90)),
    ("NT-LV03", "Curl Activator 300 ml", "Leave-in", 14.55, 1.80, (39.90, 39.90, 44.90, 42.90, 41.90)),
    ("NT-OL01", "Argan Oil 60 ml", "Oil", 21.45, 3.30, (64.90, 64.90, 69.90, 69.90, 66.90)),
    ("NT-OL02", "Finishing Serum 40 ml", "Oil", 18.90, 3.00, (54.90, 54.90, 57.90, 57.90, 55.90)),
    ("NT-AM01", "Hair Ampoule (unit)", "Treatment", 5.85, 0.90, (14.90, 14.90, 16.90, 16.90, 15.90)),
    ("NT-AM02", "Hair Ampoule Box (12 units)", "Treatment", 61.50, 5.25, (159.90, 159.90, 169.90, 169.90, 164.90)),
    ("NT-TN01", "Hair Tonic 100 ml", "Treatment", 23.70, 2.40, (69.90, 69.90, 74.90, 74.90, 72.90)),
    ("NT-KT01", "Repair Kit (Shampoo + Conditioner + Mask)", "Kit", 50.55, 9.75, (129.90, 129.90, 139.90, 139.90, 134.90)),
    ("NT-KT02", "Curl Kit (Shampoo + Conditioner + Activator)", "Kit", 46.50, 9.75, (119.90, 119.90, 124.90, 124.90, 119.90)),
    ("NT-KT03", "Travel Kit (4 minis)", "Kit", 18.60, 4.80, (49.90, 49.90, 49.90, 49.90, 47.90)),
    ("NT-ST01", "Styling Gel 250 g", "Styling", 10.20, 1.35, (24.90, 24.90, 26.90, 26.90, 24.90)),
    ("NT-ST02", "Texturizing Spray 200 ml", "Styling", 17.85, 2.25, (44.90, 44.90, 46.90, 46.90, 44.90)),
]
NP, NC = len(PRODUCTS), len(CHANNELS)


def style_range(ws, ref, font=body_font, fill=None, fmt=None, align=None, border=box):
    for row in ws[ref]:
        for c in row:
            c.font = font
            if fill:
                c.fill = fill
            if fmt:
                c.number_format = fmt
            if align:
                c.alignment = align
            if border:
                c.border = border


def header(ws, row, col, labels, widths=None, height=30):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        c.font, c.fill, c.alignment, c.border = head_font, head_fill, center, box
    ws.row_dimensions[row].height = height
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[ws.cell(row=row, column=col + i).column_letter].width = w


def page_title(ws, title, sub, back=True):
    ws.sheet_view.showGridLines = False
    ws["B2"] = title
    ws["B2"].font = title_font
    ws["B3"] = sub
    ws["B3"].font = sub_font
    ws.column_dimensions["A"].width = 2
    if back:
        c = ws["B4"]
        c.value = "← Back to Dashboard"
        c.hyperlink = "#'Dashboard'!A1"
        c.font = Font(name=F, size=9, color="1F4FBF", underline="single")


def value_labels():
    """Value-only data labels (openpyxl otherwise shows series and category names too)."""
    return DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)


def show_axes(chart):
    """openpyxl marks axes as deleted by default and lets title/legend overlap the plot; fix both."""
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if chart.legend is not None:
        chart.legend.overlay = False
    if chart.title is not None:
        chart.title.overlay = False


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def build():
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    dash = wb.active
    dash.title = "Dashboard"
    st = wb.create_sheet("Settings")
    pr = wb.create_sheet("Products")
    pc = wb.create_sheet("Pricing")
    sim = wb.create_sheet("Simulator")
    sub = "Nativa Hair Care (fictitious brand, Brazil) · amounts in BRL · channel fees are illustrative · portfolio project"

    # ---------------- Settings ----------------
    page_title(st, "Settings", "Blue cells on yellow are inputs. Everything else is calculated. " + sub)
    st["B6"], st["B6"].font = "Company parameters", bold
    params = [
        ("Sales tax (Simples Nacional, effective rate)", 0.08, PCT, "TaxRate"),
        ("Overhead (fixed costs as % of revenue)", 0.12, PCT, "Overhead"),
        ("Target net margin", 0.15, PCT, "TargetMargin"),
        ("Price ending (psychological pricing)", 0.90, '0.00', "PriceEnding"),
    ]
    for i, (lab, val, fmt, nm) in enumerate(params):
        r = 7 + i
        st.cell(row=r, column=2, value=lab).font = body_font
        c = st.cell(row=r, column=3, value=val)
        c.font, c.fill, c.number_format, c.border = input_font, input_fill, fmt, box
        name(wb, nm, f"Settings!$C${r}")
    st.column_dimensions["B"].width = 44
    st.column_dimensions["C"].width = 16

    st["B13"], st["B13"].font = "Sales channels", bold
    header(st, 14, 2, ["Channel", "Commission %", "Fixed fee per unit (R$)", "Payment fee %",
                       "Shipping per unit (R$)", "Marketing %"], [44, 16, 18, 16, 18, 14])
    for i, ch in enumerate(CHANNELS):
        r = 15 + i
        for j, v in enumerate(ch):
            c = st.cell(row=r, column=2 + j, value=v)
            c.border = box
            if j == 0:
                c.font = body_font
            else:
                c.font, c.fill = input_font, input_fill
                c.number_format = PCT if j in (1, 3, 5) else MONEY
    last_ch = 14 + NC
    name(wb, "ChannelList", f"Settings!$B$15:$B${last_ch}")
    st[f"B{last_ch + 2}"] = ("Fees are illustrative, close to public rates of each marketplace in 2026. "
                             "Update them with the client's real contract before using the prices.")
    st[f"B{last_ch + 2}"].font = sub_font

    # ---------------- Products ----------------
    page_title(pr, "Products", "Landed cost = unit cost + packaging + inbound freight. Current prices per channel are inputs. " + sub)
    pr["C5"], pr["C5"].font = "Inbound freight (% of unit cost)", body_font
    pr["E5"] = 0.05
    pr["E5"].font, pr["E5"].fill, pr["E5"].number_format, pr["E5"].border = input_font, input_fill, PCT, box
    name(wb, "FreightIn", "Products!$E$5")
    ch_names = [c[0] for c in CHANNELS]
    cols = ["SKU", "Product", "Category", "Unit cost", "Packaging", "Landed cost"] + [f"{c}" for c in ch_names]
    header(pr, 7, 2, cols, [11, 40, 13, 12, 12, 13] + [14] * NC)
    pr.cell(row=6, column=8, value="Current price by channel").font = bold
    for i, (sku, prod, cat, cost, pack, prices) in enumerate(PRODUCTS):
        r = 8 + i
        vals = [sku, prod, cat, cost, pack, f"=E{r}+F{r}+E{r}*FreightIn"] + list(prices)
        for j, v in enumerate(vals):
            c = pr.cell(row=r, column=2 + j, value=v)
            c.border = box
            if j in (3, 4) or j >= 6:
                c.font, c.fill, c.number_format = input_font, input_fill, MONEY
            elif j == 5:
                c.font, c.number_format = bold, MONEY
            else:
                c.font = body_font
    last_p = 7 + NP
    pr.freeze_panes = "D8"
    name(wb, "ProductList", f"Products!$C$8:$C${last_p}")

    # ---------------- Pricing (one row per product x channel) ----------------
    page_title(pc, "Pricing by Product and Channel",
               "Suggested price = (landed cost + fixed costs per unit) / (1 - % costs - target margin). " + sub)
    pcols = ["SKU", "Product", "Category", "Channel", "Landed cost", "% costs on price", "Fixed costs per unit",
             "Break-even price", "Suggested price", "Current price", "Margin at current price (R$)",
             "Margin at current price %", "Price change needed", "Status", "_gap_rank"]
    header(pc, 6, 2, pcols, [11, 38, 13, 15, 13, 13, 13, 13, 13, 13, 15, 15, 14, 15, 6], height=48)
    r = 7
    P = f"Products!$B$8:$B${last_p}"
    for i, (sku, prod, cat, *_rest) in enumerate(PRODUCTS):
        for j, ch in enumerate(ch_names):
            chrow = f"MATCH($E{r},ChannelList,0)"
            pct = (f"=TaxRate+Overhead+INDEX(Settings!$C$15:$C${last_ch},{chrow})"
                   f"+INDEX(Settings!$E$15:$E${last_ch},{chrow})+INDEX(Settings!$G$15:$G${last_ch},{chrow})")
            fixed = f"=INDEX(Settings!$D$15:$D${last_ch},{chrow})+INDEX(Settings!$F$15:$F${last_ch},{chrow})"
            vals = [
                sku, prod, cat, ch,
                f"=INDEX(Products!$G$8:$G${last_p},MATCH($B{r},{P},0))",
                pct, fixed,
                f"=(F{r}+H{r})/(1-G{r})",
                f"=ROUNDUP((F{r}+H{r})/(1-G{r}-TargetMargin)-PriceEnding,0)+PriceEnding",
                f"=INDEX(Products!$H$8:$L${last_p},MATCH($B{r},{P},0),MATCH($E{r},Products!$H$7:$L$7,0))",
                f"=K{r}*(1-G{r})-F{r}-H{r}",
                f"=IFERROR(L{r}/K{r},0)",
                f"=IFERROR(J{r}/K{r}-1,0)",
                f'=IF(M{r}<0,"Loss",IF(M{r}<TargetMargin,"Below target","OK"))',
                f"=N{r}+ROW()/10^9",
            ]
            for k, v in enumerate(vals):
                c = pc.cell(row=r, column=2 + k, value=v)
                c.font, c.border = body_font, box
                if k in (4, 6, 7, 8, 9, 10):
                    c.number_format = MONEY
                if k in (5, 11, 12):
                    c.number_format = PCT
                if k == 8:
                    c.font = bold
            r += 1
    last_r = r - 1
    pc.column_dimensions["P"].hidden = True
    tab = Table(displayName="tblPricing", ref=f"B6:P{last_r}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    pc.add_table(tab)
    pc.freeze_panes = "D7"
    stat = f"O7:O{last_r}"
    pc.conditional_formatting.add(stat, CellIsRule(operator="equal", formula=['"Loss"'],
                                  fill=PatternFill("solid", fgColor="F8D7CF"), font=Font(name=F, color="9B2C14", bold=True)))
    pc.conditional_formatting.add(stat, CellIsRule(operator="equal", formula=['"Below target"'],
                                  fill=PatternFill("solid", fgColor="FCF0CF"), font=Font(name=F, color="8A6100", bold=True)))
    pc.conditional_formatting.add(stat, CellIsRule(operator="equal", formula=['"OK"'],
                                  fill=PatternFill("solid", fgColor="D5EFEA"), font=Font(name=F, color="1B6B61", bold=True)))
    pc.conditional_formatting.add(f"N7:N{last_r}", FormulaRule(formula=["N7>0.05"], font=Font(name=F, color="9B2C14", bold=True)))
    rng = lambda col: f"Pricing!${col}$7:${col}${last_r}"

    # ---------------- Simulator ----------------
    page_title(sim, "Discount Simulator", "Pick a product and a channel, try a discount, and see what is left as margin. " + sub)
    sim.column_dimensions["B"].width = 34
    sim.column_dimensions["C"].width = 18
    sim.column_dimensions["D"].width = 14
    labels = [("Product", PRODUCTS[7][1]), ("Channel", "Mercado Livre"), ("Discount", 0.10)]
    for i, (lab, val) in enumerate(labels):
        rr = 6 + i
        sim.cell(row=rr, column=2, value=lab).font = bold
        c = sim.cell(row=rr, column=3, value=val)
        c.font, c.fill, c.border = input_font, input_fill, box
        if lab == "Discount":
            c.number_format = PCT
    sim.merge_cells("C6:F6")
    dv_p = DataValidation(type="list", formula1="=ProductList", allow_blank=False)
    dv_c = DataValidation(type="list", formula1="=ChannelList", allow_blank=False)
    dv_d = DataValidation(type="decimal", operator="between", formula1="0", formula2="0.9")
    for dv, cell in ((dv_p, "C6"), (dv_c, "C7"), (dv_d, "C8")):
        sim.add_data_validation(dv)
        dv.add(cell)
    chrow = "MATCH($C$7,ChannelList,0)"
    prow = f"MATCH($C$6,Products!$C$8:$C${last_p},0)"
    helpers = [
        ("Landed cost", f"=INDEX(Products!$G$8:$G${last_p},{prow})", MONEY),
        ("Current price", f"=INDEX(Products!$H$8:$L${last_p},{prow},MATCH($C$7,Products!$H$7:$L$7,0))", MONEY),
        ("Price after discount", "=C12*(1-C8)", MONEY),
        ("Suggested price (target margin)",
         f"=INDEX({rng('J')},MATCH(1,INDEX(({rng('C')}=$C$6)*({rng('E')}=$C$7),0),0))", MONEY),
        ("Maximum discount that keeps the target margin",
         f"=MAX(0,1-(C11+INDEX(Settings!$D$15:$D${last_ch},{chrow})+INDEX(Settings!$F$15:$F${last_ch},{chrow}))"
         f"/(C12*(1-TargetMargin-TaxRate-Overhead-INDEX(Settings!$C$15:$C${last_ch},{chrow})"
         f"-INDEX(Settings!$E$15:$E${last_ch},{chrow})-INDEX(Settings!$G$15:$G${last_ch},{chrow}))))", PCT),
    ]
    for i, (lab, f, fmt) in enumerate(helpers):
        rr = 11 + i
        sim.cell(row=rr, column=2, value=lab).font = body_font
        c = sim.cell(row=rr, column=3, value=f)
        c.font, c.number_format, c.border = bold, fmt, box
    sim["B10"], sim["B10"].font = "Price", bold

    sim["B17"], sim["B17"].font = "Where each R$ of the discounted price goes", bold
    header(sim, 18, 2, ["Component", "R$ per unit", "% of price"])
    price = "$C$13"
    comp = [
        ("Landed cost", "=C11"),
        ("Sales tax", f"={price}*TaxRate"),
        ("Channel commission", f"={price}*INDEX(Settings!$C$15:$C${last_ch},{chrow})"),
        ("Payment fee", f"={price}*INDEX(Settings!$E$15:$E${last_ch},{chrow})"),
        ("Marketing", f"={price}*INDEX(Settings!$G$15:$G${last_ch},{chrow})"),
        ("Fixed fee + shipping", f"=INDEX(Settings!$D$15:$D${last_ch},{chrow})+INDEX(Settings!$F$15:$F${last_ch},{chrow})"),
        ("Overhead", f"={price}*Overhead"),
        ("Net margin", f"={price}-SUM(C19:C25)"),
    ]
    for i, (lab, f) in enumerate(comp):
        rr = 19 + i
        sim.cell(row=rr, column=2, value=lab).font = bold if lab == "Net margin" else body_font
        c = sim.cell(row=rr, column=3, value=f)
        c.number_format, c.border, c.font = MONEY, box, bold if lab == "Net margin" else body_font
        c2 = sim.cell(row=rr, column=4, value=f"=IFERROR(C{rr}/{price},0)")
        c2.number_format, c2.border, c2.font = PCT, box, body_font
        sim.cell(row=rr, column=2).border = box
    sim["B28"] = '=IF(C26<0,"This discount loses money on every unit.",IF(D26<TargetMargin,"Margin below target: reduce the discount or negotiate the channel fee.","Margin within target."))'
    sim["B28"].font = Font(name=F, size=11, bold=True, color=CORAL)
    sim.conditional_formatting.add("C26:D26", CellIsRule(operator="lessThan", formula=["0"],
                                   font=Font(name=F, color="9B2C14", bold=True), fill=PatternFill("solid", fgColor="F8D7CF")))
    ch = BarChart()
    ch.type, ch.style = "bar", 10
    ch.title = "Breakdown of the discounted price (R$)"
    ch.add_data(Reference(sim, min_col=3, min_row=18, max_row=26), titles_from_data=True)
    ch.set_categories(Reference(sim, min_col=2, min_row=19, max_row=26))
    ch.y_axis.scaling.min = None
    ch.x_axis.scaling.orientation = "maxMin"
    ch.legend = None
    ch.series[0].dLbls = value_labels()
    show_axes(ch)
    ch.series[0].graphicalProperties.solidFill = NAVY
    ch.height, ch.width = 9, 16
    sim.add_chart(ch, "F10")

    # ---------------- Dashboard ----------------
    page_title(dash, "Smart Pricing Dashboard", sub, back=False)
    dash["B4"] = "Go to:"
    dash["B4"].font = sub_font
    for i, (lab, sh) in enumerate([("Settings", "Settings"), ("Products", "Products"), ("Pricing", "Pricing"),
                                   ("Simulator", "Simulator")]):
        c = dash.cell(row=4, column=3 + i, value=lab)
        c.hyperlink = f"#'{sh}'!A1"
        c.font = Font(name=F, size=10, color="FFFFFF", bold=True)
        c.fill = PatternFill("solid", fgColor=TEAL)
        c.alignment = center
    for col, w in zip("BCDEFGHIJ", [30, 16, 16, 16, 16, 16, 16, 16, 16]):
        dash.column_dimensions[col].width = w
    kpis = [
        ("Products", f"=COUNTA(Products!$B$8:$B${last_p})", "0"),
        ("Channels", f"=COUNTA(ChannelList)", "0"),
        ("Avg margin at current prices", f"=AVERAGE({rng('M')})", PCT),
        ("Target margin", "=TargetMargin", PCT),
        ("Prices below target", f'=COUNTIF({rng("O")},"Below target")', "0"),
        ("Prices with loss", f'=COUNTIF({rng("O")},"Loss")', "0"),
    ]
    for i, (lab, f, fmt) in enumerate(kpis):
        col = 2 + i if i < 1 else 2 + i
        cl = dash.cell(row=6, column=2 + i, value=lab)
        cv = dash.cell(row=7, column=2 + i, value=f)
        cl.font, cl.alignment, cl.fill = sub_font, center, light_fill
        cv.font = Font(name=F, size=20, bold=True, color=CORAL if "loss" in lab or "below" in lab else NAVY)
        cv.alignment, cv.number_format, cv.fill = center, fmt, light_fill
    dash.row_dimensions[7].height = 36

    dash["B9"], dash["B9"].font = "By channel", bold
    header(dash, 10, 2, ["Channel", "Avg margin (current)", "Target", "Below target", "Loss", "Avg price change needed"])
    for i, (chn, *_r) in enumerate(CHANNELS):
        rr = 11 + i
        vals = [chn,
                f'=AVERAGEIFS({rng("M")},{rng("E")},B{rr})', "=TargetMargin",
                f'=COUNTIFS({rng("E")},B{rr},{rng("O")},"Below target")',
                f'=COUNTIFS({rng("E")},B{rr},{rng("O")},"Loss")',
                f'=AVERAGEIFS({rng("N")},{rng("E")},B{rr})']
        for j, v in enumerate(vals):
            c = dash.cell(row=rr, column=2 + j, value=v)
            c.border, c.font = box, body_font
            if j in (1, 2, 5):
                c.number_format = PCT
    end_ch = 10 + NC
    dash.conditional_formatting.add(f"C11:C{end_ch}", FormulaRule(formula=["C11<TargetMargin"],
                                    font=Font(name=F, color="9B2C14", bold=True)))

    top_row = end_ch + 3
    dash.cell(row=top_row - 1, column=2, value="Top 10 prices that need the biggest increase").font = bold
    header(dash, top_row, 2, ["Product", "Channel", "Current price", "Suggested price", "Change needed", "Status"])
    for k in range(10):
        rr = top_row + 1 + k
        m = f"MATCH(LARGE({rng('P')},{k + 1}),{rng('P')},0)"
        vals = [f"=INDEX({rng('C')},{m})", f"=INDEX({rng('E')},{m})", f"=INDEX({rng('K')},{m})",
                f"=INDEX({rng('J')},{m})", f"=INDEX({rng('N')},{m})", f"=INDEX({rng('O')},{m})"]
        for j, v in enumerate(vals):
            c = dash.cell(row=rr, column=2 + j, value=v)
            c.border, c.font = box, body_font
            c.number_format = MONEY if j in (2, 3) else (PCT if j == 4 else "General")
    st_rng = f"G{top_row + 1}:G{top_row + 10}"
    dash.conditional_formatting.add(st_rng, CellIsRule(operator="equal", formula=['"Loss"'],
                                    font=Font(name=F, color="9B2C14", bold=True)))
    dash.conditional_formatting.add(st_rng, CellIsRule(operator="equal", formula=['"Below target"'],
                                    font=Font(name=F, color="8A6100", bold=True)))

    bar = BarChart()
    bar.type, bar.style = "col", 10
    bar.title = "Average margin at current prices vs target"
    bar.add_data(Reference(dash, min_col=3, min_row=10, max_row=end_ch), titles_from_data=True)
    bar.set_categories(Reference(dash, min_col=2, min_row=11, max_row=end_ch))
    bar.y_axis.number_format = "0%"
    bar.y_axis.scaling.min = 0
    bar.series[0].graphicalProperties.solidFill = NAVY
    line = LineChart()
    line.add_data(Reference(dash, min_col=4, min_row=10, max_row=end_ch), titles_from_data=True)
    line.series[0].graphicalProperties.line.solidFill = CORAL
    line.series[0].graphicalProperties.line.dashStyle = "dash"
    line.series[0].graphicalProperties.line.width = 28000
    bar.series[0].dLbls = value_labels()
    show_axes(bar)
    line.y_axis.delete = True
    bar += line
    bar.height, bar.width = 7, 15
    bar.legend.position = "b"
    dash.add_chart(bar, "I6")

    cnt = BarChart()
    cnt.type, cnt.grouping, cnt.overlap, cnt.style = "col", "stacked", 100, 10
    cnt.title = "Prices below target and with loss, by channel"
    cnt.add_data(Reference(dash, min_col=5, max_col=6, min_row=10, max_row=end_ch), titles_from_data=True)
    cnt.set_categories(Reference(dash, min_col=2, min_row=11, max_row=end_ch))
    cnt.series[0].graphicalProperties.solidFill = AMBER
    cnt.series[1].graphicalProperties.solidFill = CORAL
    cnt.y_axis.scaling.min = 0
    cnt.height, cnt.width = 7, 15
    show_axes(cnt)
    for s_ in cnt.series:
        s_.dLbls = value_labels()
    cnt.legend.position = "b"
    dash.add_chart(cnt, "I20")

    dash["B" + str(top_row + 13)] = ("How to use: update fees and targets on Settings, costs and current prices on Products. "
                                      "Pricing and this dashboard recalculate. Use the Simulator before giving a discount.")
    dash["B" + str(top_row + 13)].font = sub_font
    dash["B" + str(top_row + 13)].alignment = Alignment(wrap_text=True, vertical="top")
    dash.merge_cells(f"B{top_row + 13}:G{top_row + 13}")
    dash.row_dimensions[top_row + 13].height = 30

    for ws in (dash, st, pr, pc, sim):
        ws.sheet_properties.tabColor = {"Dashboard": NAVY, "Settings": AMBER, "Products": AMBER,
                                        "Pricing": TEAL, "Simulator": CORAL}[ws.title]
    wb.save(OUT)
    print("ok:", OUT, f"{NP} products x {NC} channels = {NP * NC} prices")


if __name__ == "__main__":
    build()
