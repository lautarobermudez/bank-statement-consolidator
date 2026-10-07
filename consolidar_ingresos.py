#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
consolidar_ingresos.py — arma el consolidado mensual de ingresos por
transferencia de todas las cuentas bancarias, a partir de los extractos que
estan en "Extractos para facturacion".

QUE HACE (una vez por mes)
  1. Lee el extracto de cada cuenta (Brubank x2, Supervielle, Mercado Pago x2).
  2. Se queda solo con las TRANSFERENCIAS RECIBIDAS del mes (sin debitos,
     impuestos, cheques, rechazos, multas, cobros QR/tarjetas, rendimientos).
  3. Deja los extractos .xlsx mostrando solo esas filas (autofiltro + filas
     ocultas; no se borra nada, se puede desocultar). El PDF de Mercado Pago
     Unit B no se puede filtrar: se genera al lado un xlsx con sus
     transferencias recibidas (el PDF queda intacto).
  4. Arma "Consolidado ingresos bancarios - <Mes> <Año>.xlsx" (hoja Resumen +
     hoja Detalle) marcando cada fila como "Tercero" o "Cuenta propia".

USO (desde WSL; necesita openpyxl y pdfplumber, que estan en este venv)
    ~/.venvs/extractos-ingresos/bin/python consolidar_ingresos.py            # simula
    ~/.venvs/extractos-ingresos/bin/python consolidar_ingresos.py --apply    # escribe

    --mes 2026-09        mes a consolidar (default: el mes anterior a hoy)
    --carpeta RUTA       carpeta "Extractos para facturacion" (default: la de este script)
    --propio TEXTO       suma un remitente como "Cuenta propia" (repetible),
                         ej: --propio "ACME" --propio "JOHN SMITH"
    --no-filtrar         no toca los extractos, solo arma el consolidado
    --forzar             aplica aunque falte un extracto o haya mas de uno por cuenta

Sin --apply NO se escribe nada: solo muestra que incluiria y que dejaria afuera.
Los extractos se clasifican por CONTENIDO (no por filas ocultas), asi que da lo
mismo si estan crudos o ya filtrados.

Antes de modificar un extracto se guarda una copia en _originales/<AAAA-MM>/.
"""
import argparse
import calendar
import collections
import datetime
import glob
import os
import re
import shutil
import sys
import unicodedata
import zipfile

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit('Falta openpyxl. Correr con el venv de este script:\n'
                     '  ~/.venvs/extractos-ingresos/bin/python consolidar_ingresos.py')

# ---------------------------------------------------------------------------
# Configuracion: una entrada por cuenta. Si se suma una cuenta, agregarla aca.
# "patron" es relativo a la carpeta de extractos; "tipo" elige el lector.
# ---------------------------------------------------------------------------
CUENTAS = [
    dict(banco='Brubank', titular='ACME S.A.', cuenta='Brubank 000-000-000',
         patron='Acme/Brubank/*.xlsx', tipo='brubank'),
    dict(banco='Brubank', titular='ACME UNIT_C S.A.', cuenta='Brubank 000-000-000',
         patron='Unit C/*.xlsx', tipo='brubank'),
    dict(banco='Supervielle', titular='ACME UNIT_B S.A.', cuenta='Supervielle CC $',
         patron='Unit B/*Supervielle*.xlsx', tipo='supervielle'),
    dict(banco='Mercado Pago', titular='ACME UNIT_B S.A.', cuenta='Mercado Pago brand.example',
         patron='Unit B/*.pdf', tipo='mp_pdf'),
    dict(banco='Mercado Pago', titular='JANE DOE', cuenta='Mercado Pago janedoe.mp',
         patron='Jane Doe/*.xlsx', tipo='mp_xlsx'),
]

# Remitentes que son del propio grupo (no son ventas): "Cuenta propia".
# ACME en cualquier variante, y DOE junto a JANE o JANE.
PROPIOS = [re.compile(r'ACME'), re.compile(r'(?=.*DOE)(?=.*(?:JANE|JOHN))')]

# Brubank: creditos que NO son transferencias (se excluyen). Lo demas es una
# transferencia: la descripcion es el nombre del remitente.
BRUBANK_NO_TRANSF = re.compile(r'^(devol|rechazo|reversa|reintegro|ajuste|interes|rendim|anulac)', re.I)
# Creditos que se incluyen pero que huelen a movimiento del banco: se avisa.
BRUBANK_SOSPECHOSA = re.compile(r'cheque|cobro|impuesto|comisi|multa|acredit|deposito|efectivo|plazo fijo|inversi', re.I)

MESES_ES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
            'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()
    return ' '.join(s.lower().split())


def _monto(v):
    """Numero o texto ('$ 1.300.000,00', '-24.872,59', '-') -> float, o None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace('$', '').replace(' ', '').strip()
    if s in ('', '-'):
        return None
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def _fecha(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v or '').strip()
    for f in ('%d-%m-%y', '%d-%m-%Y', '%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d'):
        try:
            return datetime.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def pesos(x):
    return '$ ' + f'{x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def es_propio(remitente, extras):
    r = (remitente or '').upper()
    return any(p.search(r) for p in PROPIOS) or any(e in r for e in extras)


def _cabecera(ws, requeridas, path):
    """Busca la fila de encabezados. requeridas: {clave: texto normalizado}.
    Devuelve (fila, {clave: columna 1-based})."""
    for r in range(1, min(ws.max_row, 15) + 1):
        hdr = {_norm(ws.cell(r, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value}
        if all(v in hdr for v in requeridas.values()):
            return r, {k: hdr[v] for k, v in requeridas.items()}
    raise SystemExit(f'ERROR: {os.path.basename(path)}: no encuentro las columnas '
                     f'{list(requeridas.values())}. ¿Cambió el formato del extracto?')


def mov(row, fecha, texto, remitente, importe, incluida, categoria=None, sospechosa=False):
    return dict(row=row, fecha=fecha, texto=texto, remitente=remitente, importe=importe,
                incluida=incluida, categoria=categoria, sospechosa=sospechosa)


# ---------------------------------------------------------------------------
# Lectores. Cada uno devuelve dict(path, movs, hdr, filtros) donde "movs" son
# los CREDITOS del extracto (incluidos o no) y "filtros" describe el autofiltro
# a dejar en el xlsx: [(columna, {valores de celda que se muestran})].
# ---------------------------------------------------------------------------

def leer_brubank(path):
    ws = openpyxl.load_workbook(path).active
    h, c = _cabecera(ws, dict(fecha='fecha', desc='descripcion', cred='creditos'), path)
    movs, descs, creds = [], set(), set()
    for r in range(h + 1, ws.max_row + 1):
        f, imp = _fecha(ws.cell(r, c['fecha']).value), _monto(ws.cell(r, c['cred']).value)
        if f is None or not imp or imp <= 0:
            continue
        d = str(ws.cell(r, c['desc']).value or '').strip()
        inc = not BRUBANK_NO_TRANSF.match(d)
        movs.append(mov(r, f, d, d, imp, inc, None if inc else d, inc and bool(BRUBANK_SOSPECHOSA.search(d))))
        if inc:
            descs.add(ws.cell(r, c['desc']).value)
            creds.add(ws.cell(r, c['cred']).value)
    # se filtra por descripcion Y por importe: asi un DEBITO que se llame igual
    # que un remitente (ej. el mismo nombre en una transferencia enviada) no reaparece
    return dict(path=path, movs=movs, hdr=h, filtros=[(c['desc'], descs), (c['cred'], creds)])


def leer_supervielle(path):
    ws = openpyxl.load_workbook(path).active
    h, c = _cabecera(ws, dict(fecha='fecha', concepto='concepto', detalle='detalle', cred='credito'), path)
    movs, conceptos = [], set()
    for r in range(h + 1, ws.max_row + 1):
        f, imp = _fecha(ws.cell(r, c['fecha']).value), _monto(ws.cell(r, c['cred']).value)
        if f is None or not imp or imp <= 0:
            continue
        conc = str(ws.cell(r, c['concepto']).value or '').strip()
        det = str(ws.cell(r, c['detalle']).value or '')
        inc = _norm(conc).startswith('credito por transferencia')
        m = re.search(r'NOMBRE:\s*(.*?)\s+DOCUMENTO:', det)
        movs.append(mov(r, f, conc, (m.group(1) if m else det).strip(), imp, inc, None if inc else conc))
        if inc:
            conceptos.add(ws.cell(r, c['concepto']).value)
    return dict(path=path, movs=movs, hdr=h, filtros=[(c['concepto'], conceptos)])


def leer_mp_xlsx(path):
    ws = openpyxl.load_workbook(path).active
    h, c = _cabecera(ws, dict(fecha='release_date', tipo='transaction_type', monto='transaction_net_amount'), path)
    movs, tipos = [], set()
    for r in range(h + 1, ws.max_row + 1):
        f, imp = _fecha(ws.cell(r, c['fecha']).value), _monto(ws.cell(r, c['monto']).value)
        if f is None or not imp or imp <= 0:
            continue
        t = str(ws.cell(r, c['tipo']).value or '').strip()
        inc = t.startswith('Transferencia recibida')
        movs.append(mov(r, f, t, t.replace('Transferencia recibida', '', 1).strip(), imp, inc,
                        None if inc else ' '.join(t.split()[:3])))
        if inc:
            tipos.add(ws.cell(r, c['tipo']).value)
    return dict(path=path, movs=movs, hdr=h, filtros=[(c['tipo'], tipos)])


_FECHA_PDF = re.compile(r'^\d{2}-\d{2}-\d{4}$')
_MONTO_PDF = re.compile(r'^-?[\d.]+,\d\d$')


def leer_mp_pdf(path):
    """Resumen de cuenta de Mercado Pago en PDF. Se valida contra el propio
    resumen (entradas, salidas, saldo final y continuidad del saldo): si no
    cuadra al centavo se corta, para no consolidar numeros mal leidos."""
    try:
        import pdfplumber
    except ImportError:
        raise SystemExit('Falta pdfplumber para leer el PDF de Mercado Pago. Correr con el venv:\n'
                         '  ~/.venvs/extractos-ingresos/bin/python consolidar_ingresos.py')
    rows = []
    with pdfplumber.open(path) as pdf:
        head = pdf.pages[0].extract_text() or ''
        for pn, p in enumerate(pdf.pages, 1):
            words = p.extract_words()
            top = max((w['top'] for w in words if w['text'] == 'operación'), default=0)
            end = min((w['top'] for w in words if w['text'] == 'generación:'), default=1e9) - 8
            body = [w for w in words if top + 3 < w['top'] < end]
            anchors = []
            for w in body:
                if w['x0'] < 80 and _FECHA_PDF.match(w['text']):
                    line = [x for x in body if abs(x['top'] - w['top']) < 3]
                    ids = [x['text'] for x in line if 185 < x['x0'] < 250 and x['text'].isdigit()]
                    am = [x for x in line if _MONTO_PDF.match(x['text'])]
                    val = [x for x in am if x['x1'] < 370]
                    sal = [x for x in am if x['x1'] >= 370]
                    anchors.append(dict(top=w['top'], fecha=w['text'], id=ids[0] if ids else None,
                                        valor=_monto(val[0]['text']) if val else None,
                                        saldo=_monto(sal[0]['text']) if sal else None, desc=[]))
            for w in body:
                if 85 <= w['x0'] < 190 and anchors:
                    min(anchors, key=lambda a: abs(a['top'] - w['top']))['desc'].append((w['top'], w['x0'], w['text']))
            for a in anchors:
                a['desc'] = ' '.join(t for _, _, t in sorted(a['desc']))
                a['pagina'] = pn
                rows.append(a)

    nombre = os.path.basename(path)
    if not rows or any(a['valor'] is None or a['saldo'] is None or a['id'] is None or not a['desc'] for a in rows):
        raise SystemExit(f'ERROR: {nombre}: no pude leer todos los movimientos del PDF. No se consolida.')

    def resumen(etq):
        m = re.search(etq + r':\s*\$\s*(-?[\d.]+,\d\d)', head)
        return _monto(m.group(1)) if m else None
    esperado = dict(entradas=resumen('Entradas'), salidas=resumen('Salidas'), saldo=resumen('Saldo final'))
    real = dict(entradas=round(sum(a['valor'] for a in rows if a['valor'] > 0), 2),
                salidas=round(sum(a['valor'] for a in rows if a['valor'] < 0), 2), saldo=rows[-1]['saldo'])
    for k in real:
        if esperado[k] is None or abs(esperado[k] - real[k]) > 0.01:
            raise SystemExit(f'ERROR: {nombre}: {k} leído {pesos(real[k])} != resumen del PDF {pesos(esperado[k]) if esperado[k] is not None else 'no encontrado'}. No se consolida.')
    prev = 0.0
    for a in rows:
        if abs(prev + a['valor'] - a['saldo']) > 0.01:
            raise SystemExit(f'ERROR: {nombre}: el saldo no cierra en la operación {a["id"]} (pág. {a["pagina"]}). No se consolida.')
        prev = a['saldo']

    movs = []
    for i, a in enumerate(rows):
        if a['valor'] <= 0:
            continue
        d, m, y = map(int, a['fecha'].split('-'))
        inc = 'Transferencia recibida' in a['desc']
        rem = re.sub(r'\s+', ' ', a['desc'].replace('Transferencia recibida', '')).strip()
        x = mov(i, datetime.date(y, m, d), a['desc'], rem, a['valor'], inc, None if inc else ' '.join(a['desc'].split()[:3]))
        x['id'] = a['id']
        movs.append(x)
    return dict(path=path, movs=movs, hdr=None, filtros=None, pdf=True)


LECTORES = dict(brubank=leer_brubank, supervielle=leer_supervielle, mp_xlsx=leer_mp_xlsx, mp_pdf=leer_mp_pdf)


# ---------------------------------------------------------------------------
# Escritura
# ---------------------------------------------------------------------------
NAVY = '1F3864'
_thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
HFONT, HFILL, BAND = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor=NAVY), PatternFill('solid', fgColor='EEF2F8')
MONEY = '"$" #,##0.00'


def _encabezado(ws, fila, textos):
    for i, t in enumerate(textos, 1):
        c = ws.cell(fila, i, t)
        c.font, c.fill, c.border = HFONT, HFILL, BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[fila].height = 22


def guardar(wb, destino, cache):
    """Guarda el libro e inserta el valor calculado de cada formula, para que
    se vea bien tambien en vistas previas que no recalculan (mail, celular)."""
    tmp = destino + '.tmp'
    wb.calculation.fullCalcOnLoad = True
    wb.save(tmp)
    archivo = {t: f'xl/worksheets/sheet{i}.xml' for i, t in enumerate(wb.sheetnames, 1)}
    zin = zipfile.ZipFile(tmp)
    with zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            for hoja, fn in archivo.items():
                if item.filename == fn:
                    s = data.decode('utf-8')

                    def rellenar(m, hoja=hoja):
                        v = cache[(hoja, m.group(1))]
                        return re.sub(r'<v\s*/>|<v></v>', f'<v>{v!r}</v>', m.group(0))
                    s, n = re.subn(r'<c r="([A-Z]+\d+)"[^>]*><f>.*?</f>(?:<v></v>|<v\s*/>)</c>', rellenar, s, flags=re.S)
                    esperadas = sum(1 for k in cache if k[0] == hoja)
                    if n != esperadas:
                        raise SystemExit(f'ERROR interno: {n} formulas con valor en {hoja}, esperaba {esperadas}.')
                    data = s.encode('utf-8')
            zout.writestr(item, data)
    zin.close()
    os.remove(tmp)


def escribir_pdf_xlsx(destino, cuenta, mes_txt, movs):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Transferencias recibidas'
    ws['A1'] = f"{cuenta['titular']} - {cuenta['cuenta']}"
    ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = f'{mes_txt} - solo transferencias recibidas (extraído del resumen de cuenta en PDF)'
    ws['A2'].font = Font(italic=True, color='595959')
    _encabezado(ws, 4, ['Fecha', 'Remitente', 'ID de la operación', 'Valor'])
    n0 = 5
    for i, m in enumerate(movs):
        r = n0 + i
        for c, v in enumerate([m['fecha'], m['remitente'], m['id'], m['importe']], 1):
            ws.cell(r, c, v).border = BORDER
        ws.cell(r, 1).number_format = 'dd-mm-yyyy'
        ws.cell(r, 3).number_format = '@'
        ws.cell(r, 4).number_format = MONEY
    n1 = n0 + len(movs) - 1
    ws.cell(n1 + 2, 3, 'TOTAL (según filtros)').font = Font(bold=True)
    t = ws.cell(n1 + 2, 4, f'=SUBTOTAL(109,D{n0}:D{n1})')
    t.number_format, t.font = MONEY, Font(bold=True)
    for col, w in zip('ABCD', [13, 40, 22, 20]):
        ws.column_dimensions[col].width = w
    ws.auto_filter.ref = f'A4:D{n1}'
    ws.freeze_panes = 'A5'
    guardar(wb, destino, {('Transferencias recibidas', f'D{n1 + 2}'): round(sum(m['importe'] for m in movs), 2)})


def filtrar_en_sitio(lec):
    """Deja el xlsx mostrando solo las filas incluidas: autofiltro real + filas
    ocultas. No borra nada."""
    wb = openpyxl.load_workbook(lec['path'])
    ws = wb.active
    h, last = lec['hdr'], ws.max_row
    ws.auto_filter.ref = f'A{h}:{get_column_letter(ws.max_column)}{last}'
    ws.auto_filter.filterColumn = []
    for col, valores in lec['filtros']:
        ws.auto_filter.add_filter_column(col - 1, sorted(valores, key=str))
    visibles = {m['row'] for m in lec['movs'] if m['incluida']}
    for r in range(h + 1, last + 1):
        ws.row_dimensions[r].hidden = r not in visibles
    wb.save(lec['path'])


def escribir_consolidado(destino, filas, mes_txt, cuentas, periodo_txt):
    total = round(sum(f['importe'] for f in filas), 2)
    wb = openpyxl.Workbook()
    rs, dt = wb.active, wb.create_sheet('Detalle')
    rs.title = 'Resumen'
    cache = {}

    # --- Detalle
    dt['A1'] = f'Detalle de ingresos por transferencia - {mes_txt}'
    dt['A1'].font = Font(bold=True, size=13)
    dt['A2'] = 'Una fila por transferencia recibida. Filtrá por columna; el total de abajo respeta los filtros.'
    dt['A2'].font = Font(italic=True, color='595959')
    _encabezado(dt, 4, ['Fecha', 'Banco', 'Titular de la cuenta', 'Cuenta', 'Remitente', 'Importe', 'Tipo'])
    n0 = 5
    for i, f in enumerate(filas):
        r = n0 + i
        for c, v in enumerate([f['fecha'], f['banco'], f['titular'], f['cuenta'], f['remitente'], f['importe'], f['tipo']], 1):
            cell = dt.cell(r, c, v)
            cell.border = BORDER
            cell.alignment = Alignment(vertical='center')
            if i % 2 == 0:
                cell.fill = BAND
        dt.cell(r, 1).number_format = 'dd/mm/yyyy'
        dt.cell(r, 1).alignment = Alignment(horizontal='center')
        dt.cell(r, 6).number_format = MONEY
    n1 = max(n0 + len(filas) - 1, n0)
    tr = n1 + 2
    dt.cell(tr, 5, 'TOTAL (según filtros)').font = Font(bold=True)
    dt.cell(tr, 5).alignment = Alignment(horizontal='right')
    t = dt.cell(tr, 6, f'=SUBTOTAL(109,F{n0}:F{n1})')
    t.number_format, t.font, t.border = MONEY, Font(bold=True), Border(top=Side(style='medium'))
    cache[('Detalle', f'F{tr}')] = total
    for col, w in zip('ABCDEFG', [12, 14, 26, 26, 42, 20, 15]):
        dt.column_dimensions[col].width = w
    dt.auto_filter.ref = f'A4:G{n1}'
    dt.freeze_panes = 'A5'
    dt.page_setup.orientation, dt.page_setup.fitToWidth, dt.page_setup.fitToHeight = 'landscape', 1, 0
    dt.sheet_properties.pageSetUpPr.fitToPage = True
    dt.print_title_rows = '4:4'

    # --- Resumen
    def R(col):
        return f'Detalle!${col}${n0}:${col}${n1}'
    rs['A1'] = 'Ingresos bancarios por transferencia'
    rs['A1'].font = Font(bold=True, size=15, color=NAVY)
    rs['A2'] = periodo_txt
    rs['A2'].font = Font(bold=True, size=11)
    rs['A3'] = ('Alcance: transferencias recibidas en las cuentas de Acme S.A., Acme Unit C S.A., Acme Unit B S.A. '
                'y Jane Doe. No incluye cobros con QR ni liquidaciones de tarjetas, rendimientos ni ingresos de '
                'dinero sin remitente. "Cuenta propia" son transferencias entre cuentas del propio grupo (Acme / Doe), '
                'que no son ventas.')
    rs.merge_cells('A3:F3')
    rs['A3'].alignment = Alignment(wrap_text=True, vertical='top')
    rs['A3'].font = Font(italic=True, color='595959')
    rs.row_dimensions[3].height = 62
    rs['A5'] = 'TOTAL INGRESOS POR TRANSFERENCIA'
    rs['A5'].font = Font(bold=True, size=12)
    rs['D5'] = f'=COUNT({R("F")})'
    rs['D5'].font, rs['D5'].alignment = Font(bold=True, size=12), Alignment(horizontal='center')
    rs['E5'] = f'=SUM({R("F")})'
    rs['E5'].number_format, rs['E5'].font = MONEY, Font(bold=True, size=12, color=NAVY)
    cache[('Resumen', 'D5')], cache[('Resumen', 'E5')] = len(filas), total
    for c in 'ABCDEF':
        rs[f'{c}5'].border = Border(top=Side(style='medium', color=NAVY), bottom=Side(style='medium', color=NAVY))

    campo = {'D': 'cuenta', 'C': 'titular', 'G': 'tipo'}

    def tabla(row, titulo, heads, items, col):
        rs.cell(row, 1, titulo).font = Font(bold=True, size=12, color=NAVY)
        _encabezado(rs, row + 1, heads)
        r = first = row + 2
        for textos, clave in items:
            for c, v in enumerate(textos, 1):
                rs.cell(r, c, v)
            sel = [f for f in filas if f[campo[col]] == clave]
            rs.cell(r, 4, f'=COUNTIFS({R(col)},"{clave}")')
            rs.cell(r, 5, f'=SUMIFS({R("F")},{R(col)},"{clave}")')
            rs.cell(r, 6, f'=IF($E$5=0,0,E{r}/$E$5)')
            cache[('Resumen', f'D{r}')] = len(sel)
            cache[('Resumen', f'E{r}')] = round(sum(f['importe'] for f in sel), 2)
            cache[('Resumen', f'F{r}')] = (cache[('Resumen', f'E{r}')] / total) if total else 0
            r += 1
        rs.cell(r, 1, 'Total')
        rs.cell(r, 4, f'=SUM(D{first}:D{r - 1})')
        rs.cell(r, 5, f'=SUM(E{first}:E{r - 1})')
        rs.cell(r, 6, f'=IF($E$5=0,0,E{r}/$E$5)')
        cache[('Resumen', f'D{r}')] = sum(cache[('Resumen', f'D{k}')] for k in range(first, r))
        cache[('Resumen', f'E{r}')] = round(sum(cache[('Resumen', f'E{k}')] for k in range(first, r)), 2)
        cache[('Resumen', f'F{r}')] = (cache[('Resumen', f'E{r}')] / total) if total else 0
        for rr in range(first, r + 1):
            for c in range(1, 7):
                cell = rs.cell(rr, c)
                cell.border = BORDER
                if rr == r:
                    cell.font, cell.fill = Font(bold=True), BAND
            rs.cell(rr, 4).alignment = Alignment(horizontal='center')
            rs.cell(rr, 5).number_format = MONEY
            rs.cell(rr, 6).number_format = '0.0%'
        return r + 2

    row = 7
    row = tabla(row, 'Composición', ['Concepto', '', '', 'Cantidad', 'Importe', '% del total'],
                [(['Ingresos de terceros'], 'Tercero'), (['Transferencias entre cuentas propias'], 'Cuenta propia')], 'G')
    row = tabla(row, 'Por cuenta', ['Banco', 'Titular', 'Cuenta', 'Cantidad', 'Importe', '% del total'],
                [([c['banco'], c['titular'], c['cuenta']], c['cuenta']) for c in cuentas], 'D')
    titulares = list(dict.fromkeys(c['titular'] for c in cuentas))
    tabla(row, 'Por titular', ['Titular', '', '', 'Cantidad', 'Importe', '% del total'],
          [([t], t) for t in titulares], 'C')
    for col, w in zip('ABCDEF', [16, 28, 28, 11, 22, 12]):
        rs.column_dimensions[col].width = w
    rs.sheet_view.showGridLines = False
    rs.page_setup.orientation, rs.page_setup.fitToWidth, rs.page_setup.fitToHeight = 'portrait', 1, 1
    rs.sheet_properties.pageSetUpPr.fitToPage = True
    guardar(wb, destino, cache)


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--mes', default=None, help='AAAA-MM (default: el mes anterior a hoy)')
    ap.add_argument('--carpeta', default=None, help='carpeta "Extractos para facturacion"')
    ap.add_argument('--propio', action='append', default=[], help='remitente a tratar como "Cuenta propia" (repetible)')
    ap.add_argument('--apply', action='store_true', help='escribe los archivos (sin esto solo simula)')
    ap.add_argument('--no-filtrar', action='store_true', help='no modifica los extractos')
    ap.add_argument('--forzar', action='store_true', help='aplica aunque falte un extracto o haya mas de uno por cuenta')
    args = ap.parse_args()

    carpeta = os.path.abspath(args.carpeta or os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    if args.mes:
        m = re.fullmatch(r'(\d{4})-(\d{2})', args.mes)
        if not m or not 1 <= int(m.group(2)) <= 12:
            raise SystemExit('--mes tiene que ser AAAA-MM, ej: 2026-09')
        anio, mes = int(m.group(1)), int(m.group(2))
    else:
        ult = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
        anio, mes = ult.year, ult.month
    mes_txt = f'{MESES_ES[mes - 1]} {anio}'
    extras = [e.upper() for e in args.propio]
    salida = os.path.join(carpeta, f'Consolidado ingresos bancarios - {mes_txt}.xlsx')

    print(f'Carpeta: {carpeta}\nMes: {mes_txt}\n')
    filas, problemas, lecturas = [], [], []
    for cta in CUENTAS:
        archivos = sorted(f for f in glob.glob(os.path.join(carpeta, cta['patron']))
                          if not os.path.basename(f).startswith('~$'))
        print(f'== {cta["cuenta"]}  ({cta["titular"]})')
        if not archivos:
            problemas.append(f'falta el extracto de {cta["cuenta"]} ({cta["patron"]})')
            print(f'   ¡NO ENCONTRÉ extracto! (busqué {cta["patron"]})\n')
            continue
        con_datos = 0
        for path in archivos:
            lec = LECTORES[cta['tipo']](path)
            lec['cuenta'] = cta
            enmes = [m for m in lec['movs'] if (m['fecha'].year, m['fecha'].month) == (anio, mes)]
            fuera = len(lec['movs']) - len(enmes)
            inc = [m for m in enmes if m['incluida']]
            con_datos += bool(enmes)
            print(f'   {os.path.relpath(path, carpeta)}')
            print(f'   {len(inc)} transferencias recibidas  {pesos(sum(m["importe"] for m in inc))}'
                  + (f'   ({fuera} créditos de otros meses ignorados)' if fuera else ''))
            excl = collections.defaultdict(lambda: [0, 0.0])
            for m in enmes:
                if not m['incluida']:
                    excl[m['categoria']][0] += 1
                    excl[m['categoria']][1] += m['importe']
            for k, (n, imp) in sorted(excl.items(), key=lambda kv: -kv[1][1]):
                print(f'   · NO incluido: {k}: {n} movimientos  {pesos(imp)}')
            for m in inc:
                if m['sospechosa']:
                    print(f'   ¡REVISAR! incluido pero parece un movimiento del banco: "{m["texto"]}" {pesos(m["importe"])} ({m["fecha"]:%d/%m})')
            for m in inc:
                filas.append(dict(fecha=m['fecha'], banco=cta['banco'], titular=cta['titular'], cuenta=cta['cuenta'],
                                  remitente=m['remitente'], importe=m['importe'],
                                  tipo='Cuenta propia' if es_propio(m['remitente'], extras) else 'Tercero'))
            lec['inc_mes'] = inc
            lecturas.append(lec)
        if con_datos == 0:
            problemas.append(f'el extracto de {cta["cuenta"]} no tiene movimientos de {mes_txt}')
        elif con_datos > 1:
            problemas.append(f'{cta["cuenta"]}: más de un extracto con movimientos de {mes_txt} (¿se solapan? revisar doble conteo)')
        print()

    orden = {c['cuenta']: i for i, c in enumerate(CUENTAS)}
    filas.sort(key=lambda f: (f['fecha'], orden[f['cuenta']], -f['importe']))
    total = sum(f['importe'] for f in filas)
    propias = [f for f in filas if f['tipo'] == 'Cuenta propia']
    print('=' * 70)
    print(f'TOTAL {mes_txt}: {len(filas)} transferencias  {pesos(total)}')
    print(f'   de terceros:        {len(filas) - len(propias):>4}  {pesos(total - sum(f["importe"] for f in propias))}')
    print(f'   entre cuentas propias: {len(propias):>4}  {pesos(sum(f["importe"] for f in propias))}')
    if problemas:
        print('\nATENCIÓN:')
        for p in problemas:
            print(f'  - {p}')

    if not args.apply:
        print('\nDRY-RUN: no se escribió nada. Agregá --apply cuando los números estén bien.')
        return
    if problemas and not args.forzar:
        raise SystemExit('\nNo se escribió nada por lo de arriba. Revisalo, o agregá --forzar si es a propósito.')

    if not args.no_filtrar:
        for lec in lecturas:
            if lec.get('pdf'):
                continue
            rel = os.path.relpath(lec['path'], carpeta)
            bak = os.path.join(carpeta, '_originales', f'{anio}-{mes:02d}', rel)
            if not os.path.exists(bak):
                os.makedirs(os.path.dirname(bak), exist_ok=True)
                shutil.copy2(lec['path'], bak)
            filtrar_en_sitio(lec)
            print(f'Filtrado: {rel}')
    for lec in lecturas:
        if lec.get('pdf'):
            base = os.path.splitext(lec['path'])[0]
            destino = f'{base} - solo transferencias recibidas.xlsx'
            escribir_pdf_xlsx(destino, lec['cuenta'], mes_txt, sorted(lec['inc_mes'], key=lambda m: m['row']))
            print(f'Generado: {os.path.relpath(destino, carpeta)}')
    escribir_consolidado(salida, filas, mes_txt, CUENTAS,
                         f'{mes_txt} (1 al {calendar.monthrange(anio, mes)[1]})')
    print(f'Generado: {os.path.relpath(salida, carpeta)}')


if __name__ == '__main__':
    main()
