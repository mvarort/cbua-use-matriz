"""Matriz pública: exclusivamente recuentos agregados por centro y cuaderno."""
from io import BytesIO
import json
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from matrix_style import COLORS, TARGETS, level, total_style, updated_label, display_rows

CODES=('UX_1_1','UX_2_1','UX_2_4','UX_2_5')


def validate(data):
    if set(data)!={'kind','updated_utc','from_inclusive','rows','totals','total'} or data['kind']!='cbua_use_public_matrix_v1' or data['from_inclusive']!='2026-10-05':
        raise ValueError('Contrato incompatible')
    if set(data['totals'])!=set(CODES):
        raise ValueError('Cuadernos incompatibles')
    seen=set()
    for row in data['rows']:
        if set(row)!={'state','centre','counts'} or row['state']!='identificado' or set(row['counts'])!=set(CODES):
            raise ValueError('Fila incompatible')
        if row['state']=='identificado' and (not isinstance(row['centre'],str) or not row['centre'].strip()):
            raise ValueError('Centro incompatible')
        if row['state']!='identificado' and row['centre'] is not None:
            raise ValueError('Centro incompatible')
        key=(row['state'],row['centre'])
        if key in seen or any(type(n) is not int or n<0 for n in row['counts'].values()):
            raise ValueError('Recuento incompatible')
        seen.add(key)
    if any(type(data['totals'][c]) is not int or data['totals'][c]!=sum(r['counts'][c] for r in data['rows']) for c in CODES):
        raise ValueError('Totales contradictorios')
    if type(data['total']) is not int or data['total']!=sum(data['totals'].values()):
        raise ValueError('Total contradictorio')


def excel(columns,rows):
    book=Workbook();sheet=book.active;sheet.title='Matriz USE'
    for i,row in enumerate([columns,*rows],1):
        for j,value in enumerate(row,1):
            cell=sheet.cell(i,j,value)
            if isinstance(value,str):cell.data_type='s'
            cell.alignment=Alignment(vertical='top',wrap_text=True)
            if i==1 or i==len(rows)+1:
                cell.font=Font(bold=True,color='FFFFFF' if i==1 else '6E2332')
                if i==1:cell.fill=PatternFill('solid',fgColor='9E354B')
    for number, row in enumerate(rows, 2):
        state=level(row[0], row[-1])
        if state:
            sheet.cell(number, len(columns)).fill=PatternFill('solid',fgColor=COLORS[state].lstrip('#'))
    legend=book.create_sheet('Leyenda')
    legend.append(['Aplicación','Total del centro sumando los cuatro cuadernos'])
    for state,color in COLORS.items():
        legend.append([state]);legend.cell(legend.max_row,1).fill=PatternFill('solid',fgColor=color.lstrip('#'))
    legend.append(['Sin relleno','Centro sin objetivo definido o total general'])
    legend.append(['Centro','Con condiciones','Suficiente','Satisfactorio'])
    for centre,limits in TARGETS.items():legend.append([centre,*limits])
    legend.column_dimensions['A'].width=44
    legend.column_dimensions['B'].width=55
    sheet.freeze_panes='B2'
    sheet.auto_filter.ref=f'A1:F{len(rows)}'
    for col in 'BCDEF':sheet.column_dimensions[col].width=16
    sheet.column_dimensions['A'].width=48
    result=BytesIO();book.save(result);book.close()
    return result.getvalue()


def main():
    st.set_page_config(page_title='Matriz USE | CBUA',page_icon='📚',layout='wide')
    st.markdown('<style>[data-testid="stAppViewContainer"]{background:#FCFAFA;}[data-testid="stSidebar"]{background:#F6EDEF;}h1,h2,h3{color:#8C293F;}[data-testid="stSidebar"]{border-right:1px solid #EAD9DE;}button[kind="secondary"]{border-color:#D7B8C0;}.block-container{max-width:1250px;padding-top:3rem;}</style>',unsafe_allow_html=True)
    st.sidebar.markdown('### Biblioteca de la Universidad de Sevilla')
    st.title('Matriz de observaciones USE')
    st.caption('Resultados provisionales — estudios UX en curso')
    try:
        data=json.loads((Path(__file__).parent/'data/matrix.json').read_text(encoding='utf-8'))
        validate(data)
    except Exception:
        st.error('La matriz no está disponible o sus totales son incompatibles.')
        st.stop()
    st.markdown('**Última actualización de los datos: '+updated_label(data['updated_utc'])+'**')
    st.caption('Observaciones desde el 05/10/2026 · Actualización diaria programada a las 09:00 (Sevilla)')
    source=display_rows(data['rows'], CODES)
    st.caption('Se incluyen todos los centros del listado BUS. Un 0 indica que no hay observaciones asociadas en esta instantánea.')
    selected=st.sidebar.multiselect('Centros',[r[0] for r in source])
    rows=[r for r in source if not selected or r[0] in selected]
    rows.append(['TOTAL',*[sum(r[i] for r in rows) for i in range(1,6)]])
    columns=['Centro',*CODES,'Total']
    st.dataframe(pd.DataFrame(rows,columns=columns).style.apply(total_style,axis=1),hide_index=True,width='stretch',height=min(900,max(180,40+35*len(rows))))
    st.markdown('**Cumplimiento · solo columna Total**')
    st.markdown('<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px">'+''.join(
        '<span style="background:'+color+';color:#333333;padding:6px 10px;border-radius:6px;font-size:0.85rem">'+state+'</span>'
        for state,color in COLORS.items())+'</div>',unsafe_allow_html=True)
    st.caption('Objetivos por centro sobre la suma de los cuatro cuadernos. Sin relleno: sin objetivo definido o total general. Rector Machado permanece sin relleno.')
    st.download_button('Descargar matriz Excel (.xlsx)',excel(columns,rows),file_name='Matriz_USE.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with st.expander('Objetivos por centro'):
        st.dataframe(pd.DataFrame([[name,*limits] for name,limits in TARGETS.items()],columns=['Centro','Con condiciones','Suficiente','Satisfactorio']),hide_index=True,width='stretch')
    with st.expander('Criterios del recuento'):
        st.caption('Cada respuesta de Forms cuenta como una observación de su cuaderno. No se analizan tareas. '
                   'Se excluyen las respuestas anteriores al corte y las pruebas 65/66 en UX 1_1. '
                   'Solo se incluyen observaciones con centro informado e identificado. Los centros conservan su denominación original. Los totales suman observaciones, no personas únicas. '
                   'Actualización diaria desde el ordenador de origen, si está disponible y sincronizado; no es tiempo real.')


if __name__=='__main__':
    main()
