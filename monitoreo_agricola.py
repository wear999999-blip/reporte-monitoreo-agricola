"""Reporte de cultivos/productores desde ArcGIS Survey123."""
import re
import unicodedata
from datetime import datetime
from pathlib import Path
import requests
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

URL = ('https://services.arcgis.com/53Yac7UQ7RmeEyAd/arcgis/rest/services/'
       'service_84e4fdd81cec4ddbbb996834a7349c21/FeatureServer/0')
OUT = Path(__file__).resolve().parent / 'Reportes'
OUT.mkdir(exist_ok=True)
DESTINO = OUT / 'Reporte_Arroz_en_campo.xlsx'
SESSION = requests.Session()


def get(url, **params):
    r = SESSION.get(url, params={'f': 'json', **params}, timeout=90)
    r.raise_for_status()
    d = r.json()
    if 'error' in d:
        raise RuntimeError(f"Error ArcGIS: {d['error']}")
    return d


def normalizar(x):
    s = unicodedata.normalize('NFD', str(x or '').strip().lower())
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')


def obtener_campo(campos, nombres=(), prefijo=None, obligatorio=True):
    for n in nombres:
        for c in campos:
            if normalizar(c['name']) == normalizar(n):
                return c['name']
    if prefijo:
        patron = re.compile(r'^' + re.escape(normalizar(prefijo).strip()) + r'(?=\s|$)')
        encontrados = [c['name'] for c in campos if patron.search(normalizar(c.get('alias', '')))]
        if len(encontrados) == 1:
            return encontrados[0]
    if obligatorio:
        raise ValueError(f'No se pudo identificar campo {nombres} / {prefijo}. '\
                         'Revisar los nombres impresos al inicio.')
    return None


def valor(fila, campo):
    return fila.get(campo) if campo else None


def coordenadas(registro, C):
    lat = numero(valor(registro, C['latitud']))
    lon = numero(valor(registro, C['longitud']))
    alt = numero(valor(registro, C['altura']))
    geo = registro.get('_geometry') or {}
    gx, gy = numero(geo.get('x')), numero(geo.get('y'))
    if lat is None:
        lat = gy
    if lon is None:
        lon = gx
    if alt is None:
        alt = numero(geo.get('z'))
    if lat is not None and lon is not None and abs(lat) > 20 and abs(lon) <= 20:
        print(f'AVISO: coordenadas aparentemente invertidas en registro {registro.get("_oid", "")}. Se intercambian.')
        lat, lon = lon, lat
    return lat, lon, alt


def numero(x):
    try:
        if x is None or str(x).strip() == '':
            return None
        return float(str(x).strip().replace(',', '.'))
    except (ValueError, TypeError):
        return None


def decodificar(x, campo, metadata):
    if x is None:
        return ''
    s = str(x).strip()
    if not campo:
        return s
    meta = next((c for c in metadata if c['name'] == campo), {})
    domain = (meta.get('domain') or {}).get('codedValues') or []
    mapa = {str(o['code']): str(o['name']) for o in domain}
    if s in mapa:
        return mapa[s]
    tokens = re.split(r'[,;\s]+', s)
    if tokens and all(t in mapa for t in tokens):
        return ', '.join(mapa[t] for t in tokens)
    return s


def fecha_ecuador(x):
    if x is None or str(x).strip() == '':
        return pd.NaT
    if isinstance(x, (int, float)):
        return pd.to_datetime(x, unit='ms', utc=True, errors='coerce').tz_convert('America/Guayaquil').tz_localize(None)
    t = pd.to_datetime(x, errors='coerce')
    if pd.isna(t):
        return pd.NaT
    if t.tzinfo is not None:
        return t.tz_convert('America/Guayaquil').tz_localize(None)
    return t


def ejecutar():
    print('Conectando con ArcGIS Survey123...')
    meta = get(URL)
    campos = meta['fields']
    print('Campos disponibles (nombre | alias):')
    for c in campos:
        print(' ', c['name'], '|', c.get('alias', ''))

    # Mapeo actualizado según los campos reales de tu servicio
    C = {
        'provincia': obtener_campo(campos, ['provincia']),
        'canton': obtener_campo(campos, ['canton']),
        'parroquia': obtener_campo(campos, ['parroquia']),
        'latitud': obtener_campo(campos, ['Y', 'latitud', 'latitude'], obligatorio=False),
        'longitud': obtener_campo(campos, ['X', 'longitud', 'longitude'], obligatorio=False),
        'altura': obtener_campo(campos, ['Z', 'altura', 'altitud', 'elevation'], obligatorio=False),
        'productor': obtener_campo(campos, ['nombre_productor'], prefijo='B1.'),
        'cedula_prod': obtener_campo(campos, ['cedula_productor'], prefijo='B2.'),
        'superficie': obtener_campo(campos, ['superficie_sembrada'], prefijo='C1.'),
        'etapa': obtener_campo(campos, ['etapa_fenologica'], prefijo='C2.'),
        'fecha_siembra': obtener_campo(campos, ['fecha_siembra'], prefijo='C3.', obligatorio=False),
        'fecha_cosecha': obtener_campo(campos, ['fecha_cosecha'], prefijo='C4.', obligatorio=False),
        'produccion': obtener_campo(campos, ['produccion_sacas'], prefijo='C5.'),
        'otros_sacas': obtener_campo(campos, ['otros_sacas'], prefijo='C5.1', obligatorio=False),
        'inundaciones': obtener_campo(campos, ['inundaciones'], prefijo='C6.', obligatorio=False),
        'afectacion': obtener_campo(campos, ['afectacion'], prefijo='C6.1', obligatorio=False),
        'variedad': obtener_campo(campos, ['variedad'], prefijo='C7.', obligatorio=False),
        'tecnico': obtener_campo(campos, ['nombre_tecnico'], obligatorio=False),
        'cedula_tec': obtener_campo(campos, ['cedula_tecnico'], obligatorio=False),
        'fecha': obtener_campo(campos, ['fecha', 'fecha_levantamiento', 'end'], obligatorio=False),
        'observaciones': obtener_campo(campos, ['observaciones'], obligatorio=False),
    }

    for k, v in C.items():
        print(f'{k:15}: {v}')

    oid = meta.get('objectIdField') or next(c['name'] for c in campos if c['type'] == 'esriFieldTypeOID')
    ids = get(URL + '/query', where='1=1', returnIdsOnly='true').get('objectIds') or []
    registros = []
    for i in range(0, len(ids), 200):
        batch = ids[i:i+200]
        d = get(URL + '/query', objectIds=','.join(map(str, batch)), outFields='*', returnGeometry='true', outSR='4326')
        for f in d.get('features', []):
            a = dict(f['attributes'])
            a['_geometry'] = f.get('geometry') or {}
            a['_oid'] = a.get(oid)
            registros.append(a)
        print(f'Descargados {len(registros)} de {len(ids)}')

    if len(registros) != len(ids):
        raise RuntimeError('La descarga está incompleta.')

    filas = []
    for r in registros:
        lat, lon, alt = coordenadas(r, C)
        filas.append({
            '_ObjectID': r[oid],
            'Provincia': valor(r, C['provincia']),
            'Cantón': valor(r, C['canton']),
            'Parroquia': valor(r, C['parroquia']),
            'Latitud': lat,
            'Longitud': lon,
            'Altura': alt,
            'Nombre del Productor': valor(r, C['productor']),
            'Cédula del Productor': str(valor(r, C['cedula_prod']) or '').strip(),
            'Superficie Sembrada (ha)': numero(valor(r, C['superficie'])),
            'Etapa Fenológica': decodificar(valor(r, C['etapa']), C['etapa'], campos),
            'Fecha de Siembra': fecha_ecuador(valor(r, C['fecha_siembra'])),
            'Fecha Aproximada Cosecha': fecha_ecuador(valor(r, C['fecha_cosecha'])),
            'Producción Esperada (sacas)': numero(valor(r, C['produccion'])),
            'Otros Sacas': numero(valor(r, C['otros_sacas'])),
            'Variedad': decodificar(valor(r, C['variedad']), C['variedad'], campos),
            'Inundaciones': decodificar(valor(r, C['inundaciones']), C['inundaciones'], campos),
            'Nivel de Afectación': decodificar(valor(r, C['afectacion']), C['afectacion'], campos),
            'Nombre Técnico': valor(r, C['tecnico']),
            'Cédula del Técnico': str(valor(r, C['cedula_tec']) or '').strip(),
            'Fecha de Registro': fecha_ecuador(valor(r, C['fecha'])),
            'Observaciones': valor(r, C['observaciones']),
        })

    df = pd.DataFrame(filas)
    
    columnas = [
        'Provincia', 'Cantón', 'Parroquia', 'Latitud', 'Longitud', 'Altura',
        'Nombre del Productor', 'Cédula del Productor', 'Superficie Sembrada (ha)',
        'Etapa Fenológica', 'Fecha de Siembra', 'Fecha Aproximada Cosecha',
        'Producción Esperada (sacas)', 'Otros Sacas', 'Variedad',
        'Inundaciones', 'Nivel de Afectación', 'Nombre Técnico',
        'Cédula del Técnico', 'Fecha de Registro', 'Observaciones'
    ]

    with pd.ExcelWriter(DESTINO, engine='openpyxl') as writer:
        df[columnas].to_excel(writer, sheet_name='REPORTE', index=False)
        ws = writer.sheets['REPORTE']
        ws.freeze_panes = 'D2'
        ws.auto_filter.ref = ws.dimensions
        for celda in ws[1]:
            celda.fill = PatternFill('solid', fgColor='17365D')
            celda.font = Font(bold=True, color='FFFFFF')
            celda.alignment = Alignment(horizontal='center', wrap_text=True)
        for j, nombre in enumerate(columnas, 1):
            ws.column_dimensions[get_column_letter(j)].width = min(max(len(nombre) + 3, 17), 36)
            if 'Fecha' in nombre:
                for i in range(2, ws.max_row + 1):
                    ws.cell(i, j).number_format = 'dd/mm/yyyy'
            if nombre in ('Cédula del Productor', 'Cédula del Técnico'):
                for i in range(2, ws.max_row + 1):
                    ws.cell(i, j).number_format = '@'

    print('\nREPORTE GENERADO:', DESTINO)
    print('Total de registros:', len(df))

if __name__ == '__main__':
    ejecutar()