"""Objetivos aportados por BUS; colores aplicables solo al total por centro."""
from datetime import datetime
from zoneinfo import ZoneInfo

# Denominaciones de Forms; correspondencias explícitas con la tabla facilitada.
TARGETS = {
    'Arquitectura y Edificación': (12, 18, 24),
    'Bellas Artes': (6, 9, 12),
    'Centros de la Salud': (10, 15, 20),
    'Comunicación': (8, 12, 16),
    'CRAI Antonio de Ulloa': (24, 36, 48),
    'Derecho y Ciencias del Trabajo': (14, 21, 28),
    'Económicas': (8, 12, 16),
    'Educación': (10, 15, 20),
    'Filosofía y Psicología': (10, 15, 20),
    'Humanidades': (22, 33, 44),
    'Informática': (4, 6, 8),
    'Ingeniería Agronómica': (4, 6, 8),
    'Ingeniería': (10, 15, 20),
    'Matemáticas': (4, 6, 8),
    'Politécnica': (4, 6, 8),
    'Turismo y Finanzas': (6, 9, 12),
}
COLORS = {
    'Por debajo del primer objetivo': '#F5E1E5',
    'Cumplido (con condiciones)': '#FFF2C9',
    'Cumplido suficiente': '#E0EFD2',
    'Cumplido satisfactorio': '#BBDFAE',
}


def level(centre, total):
    limits = TARGETS.get(centre)
    if limits is None:
        return None
    if total >= limits[2]:
        return 'Cumplido satisfactorio'
    if total >= limits[1]:
        return 'Cumplido suficiente'
    if total >= limits[0]:
        return 'Cumplido (con condiciones)'
    return 'Por debajo del primer objetivo'


def total_style(row):
    styles = [''] * len(row)
    state = level(row['Centro'], row['Total'])
    if state:
        styles[list(row.index).index('Total')] = 'background-color: ' + COLORS[state] + '; color: #333333; font-weight: 600'
    return styles


def updated_label(value):
    stamp = datetime.fromisoformat(value.replace('Z', '+00:00'))
    if stamp.tzinfo is None:
        raise ValueError('Fecha sin zona horaria')
    local = stamp.astimezone(ZoneInfo('Europe/Madrid'))
    return local.strftime('%d/%m/%Y a las %H:%M') + ' (hora de Sevilla)'


def display_rows(records, codes):
    """Incluye centros del listado BUS sin observaciones en el paquete vigente."""
    indexed = {row['centre']: row['counts'] for row in records}
    centres = sorted(set(TARGETS) | set(indexed))
    result = []
    for centre in centres:
        counts = indexed.get(centre, {})
        values = [counts.get(code, 0) for code in codes]
        result.append([centre, *values, sum(values)])
    return result
