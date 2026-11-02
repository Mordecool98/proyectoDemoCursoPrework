# -*- coding: utf-8 -*-
"""Genera el libro de Excel de registro e indicadores de bombeo.

Uso:
    python generar_libro.py            -> crea la versión vacía y la de ejemplo
"""
import random
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- estilos
AZUL = "1F4E79"
CELESTE = "DDEBF7"
AMARILLO = "FFF2CC"
GRIS = "F2F2F2"
ROJO = "F8CBAD"
VERDE = "C6EFCE"

F_TIT = Font(bold=True, size=14, color=AZUL)
F_SUB = Font(bold=True, size=11, color=AZUL)
F_HDR = Font(bold=True, color="FFFFFF")
FILL_HDR = PatternFill("solid", fgColor=AZUL)
FILL_IN = PatternFill("solid", fgColor=AMARILLO)   # celdas de captura
FILL_CALC = PatternFill("solid", fgColor=GRIS)     # celdas calculadas
FILL_ROJO = PatternFill("solid", fgColor=ROJO)
FILL_VERDE = PatternFill("solid", fgColor=VERDE)
THIN = Side(style="thin", color="BFBFBF")
BORDE = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")

N_REG = 5000      # filas del registro diario
N_ENE = 1000      # filas de energía
N_MAN = 1000      # filas de mantenimiento
N_EQ = 12         # equipos posibles
N_MED = 8         # medidores de energía posibles
N_MESES = 36      # meses en histórico

R0 = 5            # primera fila de datos en hojas de captura
REG_END = R0 + N_REG - 1
ENE_END = R0 + N_ENE - 1
MAN_END = R0 + N_MAN - 1
EQ_END = R0 + N_EQ - 1
MED_END = R0 + N_MED - 1


def rng(sheet, col, start=R0, end=REG_END):
    return f"{sheet}!${col}${start}:${col}${end}"


REG = lambda c: rng("Registro_Diario", c, R0, REG_END)
ENE = lambda c: rng("Energia", c, R0, ENE_END)
MAN = lambda c: rng("Mantenimiento", c, R0, MAN_END)


def encabezado(ws, fila, columnas, anchos=None, alto=45):
    for i, txt in enumerate(columnas, start=1):
        c = ws.cell(row=fila, column=i, value=txt)
        c.font = F_HDR
        c.fill = FILL_HDR
        c.alignment = WRAP
        c.border = BORDE
    ws.row_dimensions[fila].height = alto
    if anchos:
        for i, w in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def titulo(ws, texto, sub=None):
    ws["A1"] = texto
    ws["A1"].font = F_TIT
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, color="595959")


def estilo_rango(ws, r1, r2, c1, c2, fill=None, fmt=None):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDE
            if fill:
                cell.fill = fill
            if fmt:
                cell.number_format = fmt


def nombre(wb, n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


# ================================================================ libro
def construir(ejemplo=False):
    wb = Workbook()

    # ------------------------------------------------------- Instrucciones
    ws = wb.active
    ws.title = "Instrucciones"
    titulo(ws, "Sistema de registro e indicadores de bombeo",
           "Pozos 1 y 2, Fosa 1 (2 líneas) y Fosa 2 (2 líneas)")
    ws.column_dimensions["A"].width = 120
    lineas = [
        ("CÓMO ESTÁ ORGANIZADO", True),
        ("Celdas AMARILLAS = se escriben a mano. Celdas GRISES = fórmulas, no tocar.", False),
        ("1. Parametros: umbrales de alerta (cloro, desbalance, eficiencia, etc.). Ajústelos a su realidad.", False),
        ("2. Equipos: datos de placa de cada pozo y línea de bombeo. Llenar UNA vez y actualizar si cambia un equipo.", False),
        ("3. Medidores: contadores de energía y su línea base de kWh/m³.", False),
        ("4. Registro_Diario: una fila por equipo por día (o por turno). Se copia del formato de papel.", False),
        ("5. Energia: kWh y costo por medidor, tomados del portal web de la distribuidora eléctrica.", False),
        ("6. Mantenimiento: cada intervención, preventiva o correctiva, con horas fuera de servicio.", False),
        ("7. Indicadores_Mes: escriba el mes en la celda amarilla y verá indicadores y alertas por equipo.", False),
        ("8. Historico_Proyeccion: producción y energía por mes, con proyección de los meses siguientes.", False),
        ("9. Tendencias_Equipos: caudal medio por equipo y kWh/m³ por medidor mes a mes, para ver desgaste.", False),
        ("10. Formato_Papel: hoja imprimible con las mismas columnas que Registro_Diario.", False),
        ("", False),
        ("RUTINA SUGERIDA", True),
        ("Diario: el operador llena el formato de papel. Al final del día (o de la semana) se pasa a Registro_Diario.", False),
        ("Mensual: cargar kWh y costo del portal eléctrico en Energia, con fecha inicio = día 1 y fecha fin = último día del mes.", False),
        ("Mensual: revisar Indicadores_Mes y Tendencias_Equipos. Cada alerta debe terminar en una orden de trabajo o en una nota que explique por qué no.", False),
        ("Al cerrar cada mantenimiento: registrarlo en Mantenimiento con el horómetro del equipo.", False),
        ("", False),
        ("SUPUESTOS QUE DEBE CONFIRMAR", True),
        ("Cada pozo y cada línea de bombeo tiene su propio macromedidor y horómetro. Si una fosa tiene un solo macromedidor para sus 2 líneas, registre el volumen en la línea que operó o reparta según horas.", False),
        ("Los pozos cuentan como PRODUCCIÓN y las fosas como REBOMBEO (no se suman dos veces). Se cambia en Equipos, columna 'Cuenta como producción'.", False),
        ("Hay un medidor de energía por pozo y por fosa. Si un medidor alimenta varias estaciones, asigne el mismo medidor a esos equipos en la hoja Equipos.", False),
        ("Los umbrales de cloro (0.5 a 1.0 mg/L) son una referencia; confírmelos con la norma COGUANOR NTG 29001 vigente y con su supervisor del MSPAS.", False),
        ("", False),
        ("CÓMO LEER LOS INDICADORES CLAVE", True),
        ("Caudal medio (m³/h) = volumen / horas. Si baja con el tiempo a igual nivel de agua, la bomba se está desgastando o hay obstrucción.", False),
        ("Capacidad específica (m³/h por metro) = caudal / (nivel dinámico - nivel estático). Si baja, el pozo se está colmatando o la rejilla se incrusta.", False),
        ("kWh/m³ = energía / volumen. Si sube, se gasta más energía para mover la misma agua: baja eficiencia de bomba o motor.", False),
        ("Desbalance de corriente = máxima desviación de una fase respecto al promedio. Más de 5 % calienta el motor y acorta su vida.", False),
        ("Disponibilidad = 1 - horas fuera de servicio / horas del mes.", False),
        ("Proyección: tendencia lineal de los meses con datos. Necesita al menos 3 meses; con 12 o más meses es más confiable. No capta la estacionalidad (época seca y lluviosa); compare con el mismo mes del año anterior.", False),
        ("No deje meses incompletos en el histórico al proyectar: un mes a medias baja la tendencia.", False),
    ]
    for i, (t, b) in enumerate(lineas, start=4):
        c = ws.cell(row=i, column=1, value=t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if b:
            c.font = F_SUB

    # ------------------------------------------------------- Parametros
    wp = wb.create_sheet("Parametros")
    titulo(wp, "Parámetros y umbrales de alerta", "Modifique solo la columna C")
    encabezado(wp, 4, ["Nombre", "Descripción", "Valor", "Unidad / nota"], [14, 60, 12, 50], 25)
    params = [
        ("P_KWH", "Aumento máximo de kWh/m³ sobre la línea base", 0.15, "15 %", "0%"),
        ("P_DESB", "Desbalance máximo de corriente entre fases", 0.05, "5 % (referencia NEMA)", "0%"),
        ("P_CORR", "Corriente máxima respecto a la nominal de placa", 1.00, "100 %", "0%"),
        ("P_CAUDAL", "Caudal mínimo respecto al nominal de la bomba", 0.85, "85 %", "0%"),
        ("P_CESP", "Caída máxima de capacidad específica del pozo vs base", 0.15, "15 %", "0%"),
        ("P_MANT", "Avisar mantenimiento preventivo al llegar a", 0.90, "90 % del intervalo en horas", "0%"),
        ("P_DISP", "Disponibilidad mínima mensual del equipo", 0.95, "95 %", "0%"),
        ("P_CLMIN", "Cloro residual libre mínimo", 0.5, "mg/L - verificar NTG 29001", "0.00"),
        ("P_CLMAX", "Cloro residual libre máximo", 1.0, "mg/L - verificar NTG 29001", "0.00"),
    ]
    for i, (n, d, v, u, fmt) in enumerate(params, start=5):
        wp.cell(row=i, column=1, value=n)
        wp.cell(row=i, column=2, value=d)
        c = wp.cell(row=i, column=3, value=v)
        c.fill = FILL_IN
        c.number_format = fmt
        wp.cell(row=i, column=4, value=u)
        nombre(wb, n, f"Parametros!$C${i}")
    estilo_rango(wp, 5, 4 + len(params), 1, 4)

    # ------------------------------------------------------- Equipos
    we = wb.create_sheet("Equipos")
    titulo(we, "Catálogo de equipos de bombeo", "Datos de placa y líneas base. Una fila por pozo o línea de bombeo.")
    cols_eq = ["ID equipo", "Nombre", "Ubicación", "Tipo", "Cuenta como producción",
               "Medidor de energía", "Potencia (HP)", "Corriente nominal (A)", "Voltaje nominal (V)",
               "Caudal nominal (m³/h)", "Altura nominal (m)", "Intervalo mant. preventivo (h)",
               "Capacidad específica base (m³/h/m)", "Profundidad de bomba (m)", "Fecha instalación",
               "Marca / modelo", "Horómetro en último preventivo (al iniciar)", "Notas"]
    encabezado(we, 4, cols_eq, [11, 20, 16, 9, 12, 14, 9, 11, 10, 11, 10, 13, 14, 11, 12, 18, 15, 25], 60)
    equipos = [
        ("POZO-1", "Pozo 1", "Pozo 1", "Pozo", "Sí", "MED-POZO1"),
        ("POZO-2", "Pozo 2", "Pozo 2", "Pozo", "Sí", "MED-POZO2"),
        ("F1-L1", "Fosa 1 línea 1", "Fosa 1", "Fosa", "No", "MED-FOSA1"),
        ("F1-L2", "Fosa 1 línea 2", "Fosa 1", "Fosa", "No", "MED-FOSA1"),
        ("F2-L1", "Fosa 2 línea 1", "Fosa 2", "Fosa", "No", "MED-FOSA2"),
        ("F2-L2", "Fosa 2 línea 2", "Fosa 2", "Fosa", "No", "MED-FOSA2"),
    ]
    # datos de placa de ejemplo (solo en el libro de ejemplo)
    placa = {
        "POZO-1": (75, 92, 460, 120, 150, 2000, 6.0, 120, 2015, 0),
        "POZO-2": (60, 75, 460, 100, 140, 2000, 5.0, 110, 2018, 0),
        "F1-L1": (40, 50, 460, 140, 60, 2500, None, None, 2016, 0),
        "F1-L2": (40, 50, 460, 140, 60, 2500, None, None, 2016, 0),
        "F2-L1": (30, 38, 460, 110, 55, 2500, None, None, 2019, 0),
        "F2-L2": (30, 38, 460, 110, 55, 2500, None, None, 2019, 0),
    }
    for i, e in enumerate(equipos, start=R0):
        for j, v in enumerate(e, start=1):
            we.cell(row=i, column=j, value=v)
        if ejemplo:
            hp, a, v, q, h, im, ce, prof, anio, hor = placa[e[0]]
            vals = [hp, a, v, q, h, im, ce, prof, date(anio, 1, 15), "Ejemplo", 10000 + 1500 * (i - R0) - 400]
            for j, x in enumerate(vals, start=7):
                we.cell(row=i, column=j, value=x)
    estilo_rango(we, R0, EQ_END, 1, len(cols_eq), FILL_IN)
    for r in range(R0, EQ_END + 1):
        we.cell(row=r, column=15).number_format = "dd/mm/yyyy"
    dv = DataValidation(type="list", formula1='"Pozo,Fosa,Tanque,Otro"', allow_blank=True)
    we.add_data_validation(dv); dv.add(f"D{R0}:D{EQ_END}")
    dv = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
    we.add_data_validation(dv); dv.add(f"E{R0}:E{EQ_END}")
    dv = DataValidation(type="list", formula1=f"=Medidores!$A${R0}:$A${MED_END}", allow_blank=True)
    we.add_data_validation(dv); dv.add(f"F{R0}:F{EQ_END}")
    we.freeze_panes = "B5"

    # ------------------------------------------------------- Medidores
    wm = wb.create_sheet("Medidores")
    titulo(wm, "Medidores de energía eléctrica",
           "La línea base de kWh/m³ es el valor normal cuando los equipos están en buen estado (p. ej. promedio de los primeros 3 meses o después de un mantenimiento mayor).")
    encabezado(wm, 4, ["ID medidor", "No. de contador / NIS", "Línea base (kWh/m³)", "Distribuidora / tarifa", "Notas"],
               [14, 22, 14, 24, 40], 35)
    meds = ["MED-POZO1", "MED-POZO2", "MED-FOSA1", "MED-FOSA2"]
    base_ej = {"MED-POZO1": 0.62, "MED-POZO2": 0.58, "MED-FOSA1": 0.22, "MED-FOSA2": 0.20}
    for i, m in enumerate(meds, start=R0):
        wm.cell(row=i, column=1, value=m)
        if ejemplo:
            wm.cell(row=i, column=3, value=base_ej[m])
    estilo_rango(wm, R0, MED_END, 1, 5, FILL_IN)
    for r in range(R0, MED_END + 1):
        wm.cell(row=r, column=3).number_format = "0.000"

    # ------------------------------------------------------- Registro diario
    wr = wb.create_sheet("Registro_Diario")
    titulo(wr, "Registro diario de operación", "Una fila por equipo por día. Amarillo = captura; gris = cálculo automático.")
    cols_r = ["Fecha", "Equipo", "Horómetro inicial (h)", "Horómetro final (h)", "Horas operadas",
              "Macromedidor inicial (m³)", "Macromedidor final (m³)", "Volumen (m³)", "Caudal medio (m³/h)",
              "Corriente L1 (A)", "Corriente L2 (A)", "Corriente L3 (A)", "Corriente promedio (A)",
              "Desbalance de corriente", "Voltaje (V)", "Presión descarga (psi)",
              "Nivel estático (m)", "Nivel dinámico (m)", "Capacidad específica (m³/h/m)",
              "Cloro residual (mg/L)", "Operador", "Observaciones",
              "Producción (aux)", "Medidor (aux)", "Revisar"]
    encabezado(wr, 4, cols_r, [11, 10, 11, 11, 9, 13, 13, 10, 10, 9, 9, 9, 10, 11, 9, 10, 9, 9, 12, 10, 14, 30, 10, 12, 30], 60)
    inputs = [1, 2, 3, 4, 6, 7, 10, 11, 12, 15, 16, 17, 18, 20, 21, 22]
    for r in range(R0, REG_END + 1):
        wr.cell(row=r, column=5, value=f'=IF(OR(C{r}="",D{r}=""),"",D{r}-C{r})')
        wr.cell(row=r, column=8, value=f'=IF(OR(F{r}="",G{r}=""),"",G{r}-F{r})')
        wr.cell(row=r, column=9, value=f'=IFERROR(H{r}/E{r},"")')
        wr.cell(row=r, column=13, value=f'=IF(COUNT(J{r}:L{r})=3,AVERAGE(J{r}:L{r}),"")')
        wr.cell(row=r, column=14, value=f'=IF(COUNT(J{r}:L{r})=3,IFERROR(MAX(MAX(J{r}:L{r})-M{r},M{r}-MIN(J{r}:L{r}))/M{r},""),"")')
        wr.cell(row=r, column=19, value=f'=IF(AND(ISNUMBER(I{r}),ISNUMBER(Q{r}),ISNUMBER(R{r})),IFERROR(I{r}/(R{r}-Q{r}),""),"")')
        wr.cell(row=r, column=23, value=f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},Equipos!$A${R0}:$E${EQ_END},5,FALSE),"?"))')
        wr.cell(row=r, column=24, value=f'=IF(B{r}="","",IFERROR(VLOOKUP(B{r},Equipos!$A${R0}:$F${EQ_END},6,FALSE),"?"))')
        wr.cell(row=r, column=25, value=(
            f'=IF(B{r}="","",TRIM('
            f'IF(AND(ISNUMBER(T{r}),OR(T{r}<P_CLMIN,T{r}>P_CLMAX)),"Cloro fuera de rango. ","")'
            f'&IF(AND(ISNUMBER(N{r}),N{r}>P_DESB),"Desbalance alto. ","")'
            f'&IF(AND(ISNUMBER(E{r}),OR(E{r}<0,E{r}>24)),"Revisar horómetro. ","")'
            f'&IF(AND(ISNUMBER(H{r}),H{r}<0),"Revisar macromedidor. ","")'
            f'&IF(W{r}="?","Equipo no existe en catálogo. ","")))'))
        for c in inputs:
            wr.cell(row=r, column=c).fill = FILL_IN
        for c in (5, 8, 9, 13, 14, 19, 23, 24, 25):
            wr.cell(row=r, column=c).fill = FILL_CALC
        wr.cell(row=r, column=1).number_format = "dd/mm/yyyy"
        for c in (5, 9, 13):
            wr.cell(row=r, column=c).number_format = "0.0"
        wr.cell(row=r, column=8).number_format = "#,##0"
        wr.cell(row=r, column=14).number_format = "0.0%"
        wr.cell(row=r, column=19).number_format = "0.00"
        wr.cell(row=r, column=20).number_format = "0.00"
    dv = DataValidation(type="list", formula1=f"=Equipos!$A${R0}:$A${EQ_END}", allow_blank=True)
    wr.add_data_validation(dv); dv.add(f"B{R0}:B{REG_END}")
    dvf = DataValidation(type="date", operator="greaterThan", formula1="DATE(2000,1,1)", allow_blank=True)
    dvf.error = "Escriba una fecha válida (dd/mm/aaaa)"
    wr.add_data_validation(dvf); dvf.add(f"A{R0}:A{REG_END}")
    wr.conditional_formatting.add(f"Y{R0}:Y{REG_END}", FormulaRule(formula=[f'LEN(Y{R0})>0'], fill=FILL_ROJO))
    wr.freeze_panes = "C5"
    wr.auto_filter.ref = f"A4:Y{REG_END}"

    # ------------------------------------------------------- Energía
    wn = wb.create_sheet("Energia")
    titulo(wn, "Consumo de energía eléctrica por medidor",
           "Datos del portal web de la distribuidora. Recomendado: un registro por medidor por mes calendario.")
    encabezado(wn, 4, ["Fecha inicio", "Fecha fin", "Medidor", "Energía (kWh)", "Demanda máxima (kW)",
                       "Costo total (Q)", "Fuente / observaciones"], [12, 12, 14, 12, 12, 13, 40], 35)
    for r in range(R0, ENE_END + 1):
        for c in range(1, 8):
            wn.cell(row=r, column=c).fill = FILL_IN
        wn.cell(row=r, column=1).number_format = "dd/mm/yyyy"
        wn.cell(row=r, column=2).number_format = "dd/mm/yyyy"
        wn.cell(row=r, column=4).number_format = "#,##0"
        wn.cell(row=r, column=6).number_format = '"Q"#,##0.00'
    dv = DataValidation(type="list", formula1=f"=Medidores!$A${R0}:$A${MED_END}", allow_blank=True)
    wn.add_data_validation(dv); dv.add(f"C{R0}:C{ENE_END}")
    wn.freeze_panes = "A5"

    # ------------------------------------------------------- Mantenimiento
    wt = wb.create_sheet("Mantenimiento")
    titulo(wt, "Historial de mantenimiento", "Registre todas las intervenciones. El horómetro permite calcular horas desde el último preventivo.")
    cols_t = ["Fecha", "Equipo", "Tipo", "Componente", "Problema / motivo", "Causa", "Trabajo realizado",
              "Horas fuera de servicio", "Horómetro al intervenir (h)", "Costo repuestos (Q)",
              "Costo mano de obra (Q)", "Costo total (Q)", "Responsable", "Observaciones"]
    encabezado(wt, 4, cols_t, [11, 10, 12, 13, 30, 25, 30, 10, 12, 12, 12, 12, 16, 25], 45)
    for r in range(R0, MAN_END + 1):
        for c in range(1, 15):
            if c != 12:
                wt.cell(row=r, column=c).fill = FILL_IN
        wt.cell(row=r, column=12, value=f'=IF(AND(J{r}="",K{r}=""),"",N(J{r})+N(K{r}))').fill = FILL_CALC
        wt.cell(row=r, column=1).number_format = "dd/mm/yyyy"
        for c in (10, 11, 12):
            wt.cell(row=r, column=c).number_format = '"Q"#,##0.00'
    dv = DataValidation(type="list", formula1=f"=Equipos!$A${R0}:$A${EQ_END}", allow_blank=True)
    wt.add_data_validation(dv); dv.add(f"B{R0}:B{MAN_END}")
    dv = DataValidation(type="list", formula1='"Preventivo,Correctivo,Predictivo,Inspección"', allow_blank=True)
    wt.add_data_validation(dv); dv.add(f"C{R0}:C{MAN_END}")
    dv = DataValidation(type="list", formula1='"Motor,Bomba,Tablero eléctrico,Arrancador/variador,Válvulas,Tubería,Macromedidor,Pozo/rejilla,Otro"', allow_blank=True)
    wt.add_data_validation(dv); dv.add(f"D{R0}:D{MAN_END}")
    wt.freeze_panes = "C5"

    # ------------------------------------------------------- Indicadores mes
    wi = wb.create_sheet("Indicadores_Mes")
    titulo(wi, "Indicadores del mes y alertas")
    wi["A3"] = "Mes a revisar (día 1):"
    wi["C3"] = date(2026, 10, 1) if ejemplo else date(2026, 11, 1)
    wi["C3"].fill = FILL_IN
    wi["C3"].number_format = "mmmm yyyy"
    wi["A4"] = "Último día:"
    wi["C4"] = "=EOMONTH(C3,0)"
    wi["C4"].number_format = "dd/mm/yyyy"
    wi["A5"] = "Días del mes:"
    wi["C5"] = "=DAY(C4)"
    for c in ("A3", "A4", "A5"):
        wi[c].font = Font(bold=True)
    D1, D2 = "$C$3", "$C$4"
    fechas_reg = f'{REG("A")},">="&{D1},{REG("A")},"<="&{D2}'
    fechas_ene = f'{ENE("B")},">="&{D1},{ENE("B")},"<="&{D2}'
    fechas_man = f'{MAN("A")},">="&{D1},{MAN("A")},"<="&{D2}'

    # Resumen general
    wi["A7"] = "RESUMEN GENERAL"
    wi["A7"].font = F_SUB
    resumen = [
        ("Producción del mes (pozos) m³", f'=SUMIFS({REG("H")},{REG("W")},"Sí",{fechas_reg})', "#,##0"),
        ("Producción promedio diaria m³/día", "=C8/C5", "#,##0"),
        ("Volumen rebombeado (fosas) m³", f'=SUMIFS({REG("H")},{REG("W")},"No",{fechas_reg})', "#,##0"),
        ("Energía total kWh", f'=SUMIFS({ENE("D")},{fechas_ene})', "#,##0"),
        ("kWh por m³ producido (global)", '=IF(OR(C8=0,C11=0),"",C11/C8)', "0.000"),
        ("Costo de energía Q", f'=SUMIFS({ENE("F")},{fechas_ene})', '"Q"#,##0.00'),
        ("Costo de energía por m³ producido Q/m³", '=IF(OR(C8=0,C13=0),"",C13/C8)', '"Q"0.000'),
        ("Lecturas de cloro residual", f'=COUNTIFS({fechas_reg},{REG("T")},">=0")', "0"),
        ("Lecturas de cloro fuera de rango", f'=COUNTIFS({fechas_reg},{REG("T")},"<"&P_CLMIN)+COUNTIFS({fechas_reg},{REG("T")},">"&P_CLMAX)', "0"),
        ("Cloro residual promedio mg/L", f'=IFERROR(AVERAGEIFS({REG("T")},{fechas_reg}),"")', "0.00"),
        ("Fallas correctivas en el mes", f'=COUNTIFS({MAN("C")},"Correctivo",{fechas_man})', "0"),
        ("Horas fuera de servicio (todos los equipos)", f'=SUMIFS({MAN("H")},{fechas_man})', "0.0"),
        ("Costo de mantenimiento Q", f'=SUMIFS({MAN("L")},{fechas_man})', '"Q"#,##0.00'),
    ]
    for i, (lbl, f, fmt) in enumerate(resumen, start=8):
        wi.cell(row=i, column=1, value=lbl)
        c = wi.cell(row=i, column=3, value=f)
        c.number_format = fmt
        c.fill = FILL_CALC
        c.border = BORDE
    wi.conditional_formatting.add("C16", CellIsRule(operator="greaterThan", formula=["0"], fill=FILL_ROJO))

    # Tabla por equipo
    H_EQ = 23
    wi.cell(row=H_EQ - 1, column=1, value="INDICADORES POR EQUIPO").font = F_SUB
    cols_i = ["Equipo", "Tipo", "Horas operadas", "Utilización", "Volumen (m³)", "Caudal medio (m³/h)",
              "Caudal vs nominal", "Corriente prom. (A)", "Corriente vs nominal", "Desbalance prom.",
              "Capacidad específica (m³/h/m)", "Cap. específica vs base", "Fallas correctivas",
              "Horas fuera de servicio", "Disponibilidad", "Horómetro actual (h)",
              "Horómetro último preventivo (h)", "Horas desde preventivo", "% del intervalo de mant.",
              "Alertas", "Estado"]
    for i, t in enumerate(cols_i, start=1):
        c = wi.cell(row=H_EQ, column=i, value=t)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = WRAP; c.border = BORDE
    wi.row_dimensions[H_EQ].height = 60
    anchos = [11, 8, 10, 10, 11, 10, 10, 10, 10, 10, 12, 11, 9, 10, 11, 11, 12, 11, 11, 55, 10]
    for i, w in enumerate(anchos, start=1):
        wi.column_dimensions[get_column_letter(i)].width = w
    wi.column_dimensions["A"].width = 34
    fmts = [None, None, "0.0", "0%", "#,##0", "0.0", "0%", "0.0", "0%", "0.0%", "0.00", "0%", "0", "0.0", "0.0%",
            "#,##0", "#,##0", "#,##0", "0%", None, None]
    for k in range(N_EQ):
        r = H_EQ + 1 + k
        e = R0 + k
        A = f"$A{r}"
        rr = lambda col: f"{REG(col)},{REG('B')},{A},{fechas_reg}"
        f = {
            1: f'=IF(Equipos!A{e}="","",Equipos!A{e})',
            2: f'=IF({A}="","",Equipos!D{e}&"")',
            3: f'=IF({A}="","",SUMIFS({rr("E")}))',
            4: f'=IF({A}="","",C{r}/($C$5*24))',
            5: f'=IF({A}="","",SUMIFS({rr("H")}))',
            6: f'=IF({A}="","",IFERROR(E{r}/C{r},""))',
            7: f'=IF({A}="","",IFERROR(F{r}/Equipos!J{e},""))',
            8: f'=IF({A}="","",IFERROR(AVERAGEIFS({rr("M")}),""))',
            9: f'=IF({A}="","",IFERROR(H{r}/Equipos!H{e},""))',
            10: f'=IF({A}="","",IFERROR(AVERAGEIFS({rr("N")}),""))',
            11: f'=IF({A}="","",IFERROR(AVERAGEIFS({rr("S")}),""))',
            12: f'=IF({A}="","",IFERROR(K{r}/Equipos!M{e}-1,""))',
            13: f'=IF({A}="","",COUNTIFS({MAN("B")},{A},{MAN("C")},"Correctivo",{fechas_man}))',
            14: f'=IF({A}="","",SUMIFS({MAN("H")},{MAN("B")},{A},{fechas_man}))',
            15: f'=IF({A}="","",1-N14/($C$5*24))'.replace("N14", f"N{r}"),
            16: (f'=IF({A}="","",SUMPRODUCT(MAX(({REG("B")}={A})*({REG("A")}<={D2})*'
                 f'{REG("D")})))'),
            17: (f'=IF({A}="","",MAX(N(Equipos!Q{e}),SUMPRODUCT(MAX(({MAN("B")}={A})*({MAN("C")}="Preventivo")*'
                 f'({MAN("A")}<={D2})*{MAN("I")}))))'),
            18: f'=IF(OR({A}="",N(P{r})=0),"",P{r}-Q{r})',
            19: f'=IF({A}="","",IFERROR(R{r}/Equipos!L{e},""))',
            20: (f'=IF({A}="","",TRIM('
                 f'IF(C{r}=0,"Sin registros en el mes. ","")'
                 f'&IF(AND(ISNUMBER(G{r}),G{r}<P_CAUDAL),"Caudal bajo vs nominal. ","")'
                 f'&IF(AND(ISNUMBER(I{r}),I{r}>P_CORR),"Corriente sobre nominal. ","")'
                 f'&IF(AND(ISNUMBER(J{r}),J{r}>P_DESB),"Desbalance de corriente. ","")'
                 f'&IF(AND(ISNUMBER(L{r}),L{r}<-P_CESP),"Caída de capacidad específica del pozo. ","")'
                 f'&IF(AND(ISNUMBER(O{r}),O{r}<P_DISP),"Disponibilidad baja. ","")'
                 f'&IF(AND(ISNUMBER(S{r}),S{r}>=1),"Mantenimiento preventivo VENCIDO. ",'
                 f'IF(AND(ISNUMBER(S{r}),S{r}>=P_MANT),"Mantenimiento preventivo próximo. ",""))))'),
            21: f'=IF({A}="","",IF(T{r}="","OK","REVISAR"))',
        }
        for col, formula in f.items():
            c = wi.cell(row=r, column=col, value=formula)
            c.fill = FILL_CALC
            c.border = BORDE
            if fmts[col - 1]:
                c.number_format = fmts[col - 1]
        wi.cell(row=r, column=20).alignment = Alignment(wrap_text=True, vertical="center")
    eq_rows = f"U{H_EQ + 1}:U{H_EQ + N_EQ}"
    wi.conditional_formatting.add(eq_rows, CellIsRule(operator="equal", formula=['"REVISAR"'], fill=FILL_ROJO))
    wi.conditional_formatting.add(eq_rows, CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_VERDE))

    # Tabla por medidor
    H_ME = H_EQ + N_EQ + 3
    wi.cell(row=H_ME - 1, column=1, value="ENERGÍA POR MEDIDOR").font = F_SUB
    cols_m = ["Medidor", "Energía (kWh)", "Volumen bombeado (m³)", "kWh/m³", "Línea base kWh/m³",
              "Variación vs base", "Costo (Q)", "Q/m³", "Alerta"]
    for i, t in enumerate(cols_m, start=1):
        c = wi.cell(row=H_ME, column=i, value=t)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = WRAP; c.border = BORDE
    wi.row_dimensions[H_ME].height = 45
    fm = [None, "#,##0", "#,##0", "0.000", "0.000", "0%", '"Q"#,##0.00', '"Q"0.000', None]
    for k in range(N_MED):
        r = H_ME + 1 + k
        m = R0 + k
        A = f"$A{r}"
        f = {
            1: f'=IF(Medidores!A{m}="","",Medidores!A{m})',
            2: f'=IF({A}="","",SUMIFS({ENE("D")},{ENE("C")},{A},{fechas_ene}))',
            3: f'=IF({A}="","",SUMIFS({REG("H")},{REG("X")},{A},{fechas_reg}))',
            4: f'=IF(OR({A}="",N(B{r})=0),"",IFERROR(B{r}/C{r},""))',
            5: f'=IF(OR({A}="",Medidores!C{m}=""),"",Medidores!C{m})',
            6: f'=IF(OR(D{r}="",E{r}=""),"",D{r}/E{r}-1)',
            7: f'=IF({A}="","",SUMIFS({ENE("F")},{ENE("C")},{A},{fechas_ene}))',
            8: f'=IF(OR({A}="",N(G{r})=0),"",IFERROR(G{r}/C{r},""))',
            9: (f'=IF({A}="","",IF(AND(ISNUMBER(F{r}),F{r}>P_KWH),"kWh/m³ alto: revisar eficiencia de bombas y motores",'
                f'IF(N(B{r})=0,"Sin dato de energía en el mes","")))'),
        }
        for col, formula in f.items():
            c = wi.cell(row=r, column=col, value=formula)
            c.fill = FILL_CALC
            c.border = BORDE
            if fm[col - 1]:
                c.number_format = fm[col - 1]
    wi.conditional_formatting.add(f"I{H_ME + 1}:I{H_ME + N_MED}",
                                  FormulaRule(formula=[f'LEFT(I{H_ME + 1},3)="kWh"'], fill=FILL_ROJO))
    wi.freeze_panes = "B6"

    # ------------------------------------------------------- Histórico y proyección
    wh = wb.create_sheet("Historico_Proyeccion")
    titulo(wh, "Histórico mensual y proyección")
    wh["A3"] = "Primer mes del histórico:"
    wh["A3"].font = Font(bold=True)
    wh["C3"] = date(2026, 1, 1)
    wh["C3"].fill = FILL_IN
    wh["C3"].number_format = "mmm yyyy"
    wh["E3"] = "Tendencia de producción (m³ por mes):"
    wh["E3"].font = Font(bold=True)
    HH = 6
    cols_h = ["Mes", "No.", "Producción real (m³)", "Energía real (kWh)", "kWh/m³", "Costo energía (Q)",
              "Fallas correctivas", "Horas fuera de servicio", "Proyección producción (m³)",
              "Proyección energía (kWh)", "Proyección costo energía (Q)", "Gráfico: real", "Gráfico: proyección"]
    for i, t in enumerate(cols_h, start=1):
        c = wh.cell(row=HH, column=i, value=t)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = WRAP; c.border = BORDE
    wh.row_dimensions[HH].height = 45
    for i, w in enumerate([10, 5, 13, 12, 9, 13, 9, 10, 14, 13, 14, 12, 12], start=1):
        wh.column_dimensions[get_column_letter(i)].width = w
    h1, h2 = HH + 1, HH + N_MESES
    C_R, D_R, B_R, F_R = f"$C${h1}:$C${h2}", f"$D${h1}:$D${h2}", f"$B${h1}:$B${h2}", f"$F${h1}:$F${h2}"
    wh["I3"] = f'=IF(COUNT({C_R})<3,"Faltan datos (mínimo 3 meses)",SLOPE({C_R},{B_R}))'
    wh["I3"].number_format = "#,##0"
    wh["I3"].fill = FILL_CALC
    for k in range(N_MESES):
        r = h1 + k
        ini = f"$A{r}"
        fin = f"EOMONTH($A{r},0)"
        fr = f'{REG("A")},">="&{ini},{REG("A")},"<="&{fin}'
        fe = f'{ENE("B")},">="&{ini},{ENE("B")},"<="&{fin}'
        fmn = f'{MAN("A")},">="&{ini},{MAN("A")},"<="&{fin}'
        f = {
            1: "=$C$3" if k == 0 else f"=EDATE(A{r - 1},1)",
            2: k + 1,
            3: f'=IF(COUNTIFS({fr},{REG("W")},"Sí")=0,"",SUMIFS({REG("H")},{REG("W")},"Sí",{fr}))',
            4: f'=IF(COUNTIFS({fe})=0,"",SUMIFS({ENE("D")},{fe}))',
            5: f'=IF(OR(C{r}="",D{r}=""),"",IFERROR(D{r}/C{r},""))',
            6: f'=IF(COUNTIFS({fe})=0,"",SUMIFS({ENE("F")},{fe}))',
            7: f'=COUNTIFS({MAN("C")},"Correctivo",{fmn})',
            8: f'=SUMIFS({MAN("H")},{fmn})',
            9: f'=IF(OR(C{r}<>"",COUNT({C_R})<3),"",IF(B{r}<SUMPRODUCT(MAX(ISNUMBER({C_R})*{B_R})),"",MAX(0,FORECAST(B{r},{C_R},{B_R}))))',
            10: f'=IF(OR(D{r}<>"",COUNT({D_R})<3),"",IF(B{r}<SUMPRODUCT(MAX(ISNUMBER({D_R})*{B_R})),"",MAX(0,FORECAST(B{r},{D_R},{B_R}))))',
            11: f'=IF(OR(F{r}<>"",COUNT({F_R})<3),"",IF(B{r}<SUMPRODUCT(MAX(ISNUMBER({F_R})*{B_R})),"",MAX(0,FORECAST(B{r},{F_R},{B_R}))))',
            12: f'=IF(C{r}="",NA(),C{r})',
            13: f'=IF(I{r}="",IF(AND(C{r}<>"",C{r + 1}="",COUNT({C_R})>=3),C{r},NA()),I{r})' if k < N_MESES - 1 else f'=IF(I{r}="",NA(),I{r})',
        }
        for col, formula in f.items():
            c = wh.cell(row=r, column=col, value=formula)
            c.border = BORDE
            c.fill = FILL_CALC
        wh.cell(row=r, column=1).number_format = "mmm yyyy"
        for col in (3, 4, 9, 10, 12, 13):
            wh.cell(row=r, column=col).number_format = "#,##0"
        wh.cell(row=r, column=5).number_format = "0.000"
        for col in (6, 11):
            wh.cell(row=r, column=col).number_format = '"Q"#,##0'
        wh.cell(row=r, column=8).number_format = "0.0"
    # ocultar #N/A de las columnas auxiliares del gráfico
    wh.conditional_formatting.add(f"L{h1}:M{h2}", FormulaRule(formula=[f"ISNA(L{h1})"], font=Font(color="FFFFFF")))
    for col in (9, 10, 11):
        for r in range(h1, h2 + 1):
            wh.cell(row=r, column=col).font = Font(italic=True, color="2F5597")

    ch = LineChart()
    ch.title = "Producción mensual: real y proyectada"
    ch.y_axis.title = "m³"
    ch.height, ch.width = 9, 22
    data = Reference(wh, min_col=12, max_col=13, min_row=HH, max_row=h2)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(Reference(wh, min_col=1, min_row=h1, max_row=h2))
    ch.series[1].graphicalProperties.line.dashStyle = "dash"
    ch.x_axis.number_format = "mmm yy"
    wh.add_chart(ch, "O6")

    ch2 = LineChart()
    ch2.title = "kWh por m³ producido"
    ch2.height, ch2.width = 9, 22
    ch2.add_data(Reference(wh, min_col=5, min_row=HH, max_row=h2), titles_from_data=True)
    ch2.set_categories(Reference(wh, min_col=1, min_row=h1, max_row=h2))
    ch2.x_axis.number_format = "mmm yy"
    wh.add_chart(ch2, "O25")
    wh.freeze_panes = "B7"

    # ------------------------------------------------------- Tendencias por equipo
    wq = wb.create_sheet("Tendencias_Equipos")
    titulo(wq, "Tendencias mensuales por equipo y por medidor",
           "Caudal medio que baja o kWh/m³ que sube de forma sostenida = programar revisión antes de la falla.")
    HT = 5
    n_eq_t, n_med_t = 6, 4
    wq.cell(row=HT - 1, column=2, value="Caudal medio (m³/h)").font = F_SUB
    wq.cell(row=HT - 1, column=2 + n_eq_t, value="Capacidad específica (m³/h/m)").font = F_SUB
    wq.cell(row=HT - 1, column=4 + n_eq_t, value="kWh/m³ por medidor").font = F_SUB
    hdrs = ["Mes"] + [f"=Equipos!A{R0 + i}&\"\"" for i in range(n_eq_t)] + \
           [f"=Equipos!A{R0 + i}&\"\"" for i in range(2)] + [f"=Medidores!A{R0 + i}&\"\"" for i in range(n_med_t)]
    for i, t in enumerate(hdrs, start=1):
        c = wq.cell(row=HT, column=i, value=t)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = WRAP; c.border = BORDE
        wq.column_dimensions[get_column_letter(i)].width = 12
    for k in range(N_MESES):
        r = HT + 1 + k
        hr = h1 + k
        wq.cell(row=r, column=1, value=f"=Historico_Proyeccion!A{hr}").number_format = "mmm yyyy"
        fr = f'{REG("A")},">="&$A{r},{REG("A")},"<="&EOMONTH($A{r},0)'
        fe = f'{ENE("B")},">="&$A{r},{ENE("B")},"<="&EOMONTH($A{r},0)'
        col = 2
        for i in range(n_eq_t):
            h = f"{get_column_letter(col)}${HT}"
            wq.cell(row=r, column=col,
                    value=f'=IFERROR(SUMIFS({REG("H")},{REG("B")},{h},{fr})/SUMIFS({REG("E")},{REG("B")},{h},{fr}),NA())').number_format = "0.0"
            col += 1
        for i in range(2):
            h = f"{get_column_letter(col)}${HT}"
            wq.cell(row=r, column=col,
                    value=f'=IFERROR(AVERAGEIFS({REG("S")},{REG("B")},{h},{fr}),NA())').number_format = "0.00"
            col += 1
        for i in range(n_med_t):
            h = f"{get_column_letter(col)}${HT}"
            wq.cell(row=r, column=col,
                    value=(f'=IFERROR(IF(SUMIFS({ENE("D")},{ENE("C")},{h},{fe})=0,NA(),'
                           f'SUMIFS({ENE("D")},{ENE("C")},{h},{fe})/SUMIFS({REG("H")},{REG("X")},{h},{fr})),NA())')).number_format = "0.000"
            col += 1
        for c in range(1, col):
            wq.cell(row=r, column=c).border = BORDE
    last_c = get_column_letter(1 + n_eq_t + 2 + n_med_t)
    wq.conditional_formatting.add(f"B{HT + 1}:{last_c}{HT + N_MESES}",
                                  FormulaRule(formula=[f"ISNA(B{HT + 1})"], font=Font(color="BFBFBF")))
    t1, t2 = HT + 1, HT + N_MESES
    charts = [
        ("Caudal medio por equipo (m³/h)", 2, 1 + n_eq_t, "O5"),
        ("Capacidad específica de pozos (m³/h/m)", 2 + n_eq_t, 3 + n_eq_t, "O24"),
        ("kWh/m³ por medidor", 4 + n_eq_t, 3 + n_eq_t + n_med_t, "O43"),
    ]
    for tt, c1, c2, anchor in charts:
        cc = LineChart()
        cc.title = tt
        cc.height, cc.width = 9, 22
        cc.add_data(Reference(wq, min_col=c1, max_col=c2, min_row=HT, max_row=t2), titles_from_data=True)
        cc.set_categories(Reference(wq, min_col=1, min_row=t1, max_row=t2))
        cc.x_axis.number_format = "mmm yy"
        wq.add_chart(cc, anchor)
    wq.freeze_panes = "B6"

    # ------------------------------------------------------- Formato de papel
    wf = wb.create_sheet("Formato_Papel")
    wf["A1"] = "BITÁCORA DIARIA DE OPERACIÓN - POZOS Y FOSAS DE BOMBEO"
    wf["A1"].font = F_TIT
    wf["A3"] = "Fecha: ____/____/________      Turno: ☐ Mañana  ☐ Tarde  ☐ Noche      Operador: ______________________________"
    cols_p = ["Equipo", "Horómetro inicial", "Horómetro final", "Macromedidor inicial (m³)",
              "Macromedidor final (m³)", "L1 (A)", "L2 (A)", "L3 (A)", "Voltaje (V)", "Presión (psi)",
              "Nivel estático (m)", "Nivel dinámico (m)", "Cloro (mg/L)", "Hora lectura", "Observaciones"]
    for i, t in enumerate(cols_p, start=1):
        c = wf.cell(row=5, column=i, value=t)
        c.font = Font(bold=True); c.alignment = WRAP; c.border = BORDE; c.fill = PatternFill("solid", fgColor=CELESTE)
    wf.row_dimensions[5].height = 45
    for i, w in enumerate([10, 10, 10, 13, 13, 7, 7, 7, 8, 8, 9, 9, 8, 8, 30], start=1):
        wf.column_dimensions[get_column_letter(i)].width = w
    for k, e in enumerate(equipos):
        r = 6 + k
        wf.cell(row=r, column=1, value=e[0]).font = Font(bold=True)
        wf.row_dimensions[r].height = 26
        for c in range(1, 16):
            wf.cell(row=r, column=c).border = BORDE
        if e[3] == "Fosa":
            for c in (11, 12):
                wf.cell(row=r, column=c).fill = PatternFill("solid", fgColor="D9D9D9")
    r = 6 + len(equipos) + 1
    wf.cell(row=r, column=1, value="NOVEDADES (hora, equipo, qué pasó, qué se hizo, cómo quedó):").font = Font(bold=True)
    for k in range(1, 9):
        for c in range(1, 16):
            wf.cell(row=r + k, column=c).border = Border(bottom=THIN)
        wf.row_dimensions[r + k].height = 20
    r = r + 10
    wf.cell(row=r, column=1, value="Pendientes para el siguiente turno:").font = Font(bold=True)
    for k in range(1, 4):
        for c in range(1, 16):
            wf.cell(row=r + k, column=c).border = Border(bottom=THIN)
        wf.row_dimensions[r + k].height = 20
    r = r + 5
    wf.cell(row=r, column=1, value="Entrega: nombre y firma ________________________   Hora: ______")
    wf.cell(row=r, column=9, value="Recibe: nombre y firma ________________________   Hora: ______")
    wf.cell(row=r + 2, column=1, value="Revisado por supervisor: ________________________   Fecha: ____/____/______")
    wf.cell(row=r + 4, column=1,
            value="Reglas: usar bolígrafo; no borrar ni usar corrector (tachar con una línea, poner iniciales); anotar hora en formato 24 h; anular espacios vacíos.").font = Font(italic=True, size=9)
    wf.page_setup.orientation = "landscape"
    wf.page_setup.paperSize = wf.PAPERSIZE_LETTER
    wf.page_setup.fitToWidth = 1
    wf.page_setup.fitToHeight = 1
    wf.sheet_properties.pageSetUpPr.fitToPage = True
    wf.print_area = f"A1:O{r + 4}"

    # orden de hojas: indicadores al frente después de instrucciones
    orden = ["Instrucciones", "Indicadores_Mes", "Historico_Proyeccion", "Tendencias_Equipos",
             "Registro_Diario", "Energia", "Mantenimiento", "Equipos", "Medidores", "Parametros", "Formato_Papel"]
    wb._sheets = [wb[n] for n in orden]
    wb.active = 1

    if ejemplo:
        cargar_ejemplo(wb, equipos, placa)
    return wb


# ================================================================ datos de ejemplo
def cargar_ejemplo(wb, equipos, placa):
    """Datos ficticios ene-oct 2026 con desgaste deliberado para mostrar alertas."""
    random.seed(7)
    wr, wn, wt = wb["Registro_Diario"], wb["Energia"], wb["Mantenimiento"]
    hor = {e[0]: 10000.0 + 1500 * i for i, e in enumerate(equipos)}
    mac = {e[0]: 500000.0 + 80000 * i for i, e in enumerate(equipos)}
    vol_mes = {}
    hor_dia = {}
    ult_prev = {e[0]: 10000.0 + 1500 * i - 400 for i, e in enumerate(equipos)}
    prev_auto = []
    r = 5
    d = date(2026, 1, 1)
    fin = date(2026, 10, 31)
    while d <= fin:
        mes_idx = d.month - 1
        seca = d.month in (1, 2, 3, 4, 10)
        for e in equipos:
            eid = e[0]
            q_nom = placa[eid][3]
            amp_nom = placa[eid][1]
            if eid.startswith("F1-L2") or eid.startswith("F2-L2"):
                horas = round(random.uniform(8, 12), 1)
            else:
                horas = round(random.uniform(14, 20) + (1.5 if seca else 0), 1)
            q = q_nom * random.uniform(0.92, 0.98)
            if eid == "F1-L2":                    # bomba con desgaste progresivo
                q *= 1 - 0.022 * mes_idx
            vol = round(q * horas)
            h0, m0 = hor[eid], mac[eid]
            hor[eid] += horas
            hor_dia[(eid, d)] = round(hor[eid], 1)
            # preventivos al cumplir el intervalo; el POZO-1 deja de recibirlos en julio (ejemplo de vencido)
            if hor[eid] - ult_prev[eid] >= placa[eid][5] * 0.97 and not (eid == "POZO-1" and d.month >= 7):
                ult_prev[eid] = hor[eid]
                prev_auto.append((d, eid, "Preventivo", "Motor", f"Rutina {placa[eid][5]} h", "",
                                  "Engrase, apriete de bornes, medición de aislamiento, revisión de empaques", 4, None, 900, 600))
            mac[eid] += vol
            amps = [amp_nom * random.uniform(0.85, 0.92) for _ in range(3)]
            if eid == "POZO-1" and d.month >= 9:   # desbalance en el motor del pozo 1
                amps[2] *= 1.10
            row = {1: d, 2: eid, 3: round(h0, 1), 4: round(hor[eid], 1), 6: round(m0), 7: round(mac[eid]),
                   10: round(amps[0], 1), 11: round(amps[1], 1), 12: round(amps[2], 1),
                   15: random.choice([455, 458, 460, 462]), 16: round(random.uniform(55, 65) if e[3] == "Pozo" else random.uniform(40, 48), 1),
                   21: random.choice(["J. López", "M. Pérez", "C. Ramírez"])}
            if e[3] == "Pozo":
                est = 62 + 0.3 * mes_idx + (1.5 if seca else 0) + random.uniform(-0.3, 0.3)
                ce_base = placa[eid][6]
                ce = ce_base * (1 - (0.03 * mes_idx if eid == "POZO-2" else 0.004 * mes_idx)) * random.uniform(0.97, 1.03)
                row[17] = round(est, 2)
                row[18] = round(est + q / ce, 2)
                row[20] = round(random.uniform(0.6, 0.95) if random.random() > 0.03 else 0.35, 2)
            for c, v in row.items():
                wr.cell(row=r, column=c, value=v)
            key = (d.month, placa_medidor(eid))
            vol_mes[key] = vol_mes.get(key, 0) + vol
            r += 1
        d += timedelta(days=1)

    # energía mensual por medidor
    base = {"MED-POZO1": 0.62, "MED-POZO2": 0.58, "MED-FOSA1": 0.22, "MED-FOSA2": 0.20}
    re = 5
    for m in range(1, 11):
        ini = date(2026, m, 1)
        fn = date(2026, m + 1, 1) - timedelta(days=1)
        for med, b in base.items():
            factor = 1.0
            if med == "MED-FOSA1":
                factor = 1 + 0.025 * (m - 1)       # la bomba desgastada gasta más
            if med == "MED-POZO2":
                factor = 1 + 0.02 * (m - 1)        # pozo colmatándose: más altura de bombeo
            kwh = round(vol_mes[(m, med)] * b * factor * random.uniform(0.98, 1.02))
            vals = [ini, fn, med, kwh, round(kwh / (24 * 30) * 1.6, 1), round(kwh * 1.45, 2), "Portal web distribuidora"]
            for c, v in enumerate(vals, start=1):
                wn.cell(row=re, column=c, value=v)
            re += 1

    # mantenimiento
    mant = [
        (date(2026, 3, 18), "F2-L1", "Correctivo", "Arrancador/variador", "Disparo por sobrecarga", "Contactor dañado", "Cambio de contactor", 10, 13200, 1800, 400),
        (date(2026, 5, 22), "POZO-2", "Inspección", "Pozo/rejilla", "Revisión de niveles", "", "Se observa aumento de abatimiento", 0, 13900, 0, 300),
        (date(2026, 8, 14), "F1-L2", "Correctivo", "Bomba", "Ruido y vibración", "Rodamiento desgastado", "Cambio de rodamientos", 30, 13600, 2600, 900),
        (date(2026, 10, 9), "POZO-1", "Correctivo", "Tablero eléctrico", "Calentamiento en fase L3", "Borne flojo / posible daño en devanado", "Apriete de bornes; pendiente prueba de aislamiento", 8, 14700, 300, 500),
        (date(2026, 10, 20), "F1-L2", "Correctivo", "Bomba", "Caudal bajo", "Desgaste de impulsor", "Pendiente cotizar impulsor", 42, 14100, 0, 400),
    ]
    mant = sorted(mant + prev_auto, key=lambda x: x[0])
    for k, row in enumerate(mant, start=5):
        for c, v in enumerate(row[:9], start=1):
            wt.cell(row=k, column=c, value=v)
        wt.cell(row=k, column=9, value=hor_dia[(row[1], row[0])])
        wt.cell(row=k, column=10, value=row[9])
        wt.cell(row=k, column=11, value=row[10])
        wt.cell(row=k, column=13, value="Taller interno")


def placa_medidor(eid):
    return {"POZO-1": "MED-POZO1", "POZO-2": "MED-POZO2", "F1-L1": "MED-FOSA1",
            "F1-L2": "MED-FOSA1", "F2-L1": "MED-FOSA2", "F2-L2": "MED-FOSA2"}[eid]


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    construir(False).save(f"{out}/Registro_Indicadores_Bombeo.xlsx")
    construir(True).save(f"{out}/Registro_Indicadores_Bombeo_EJEMPLO.xlsx")
    print("ok")
