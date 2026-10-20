"""Matriz pública: exclusivamente recuentos agregados por centro y cuaderno."""
from io import BytesIO
import json
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

CODES=('UX_1_1','UX_2_1','UX_2_4','UX_2_5')


def validate(data):
    if set(data)!={'kind','updated_utc','from_inclusive','rows','totals','total'} or data['kind']!='cbua_use_public_matrix_v1' or data['from_inclusive']!='2026-10-05':
        raise ValueError('Contrato incompatible')
    if set(data['totals'])!=set(CODES):
        raise ValueError('Cuadernos incompatibles')
    seen=set()
    for row in data['rows']:
        if set(row)!={'state','centre','counts'} or row['state'] not in {'identificado','sin_centro','conflicto'} or set(row['counts'])!=set(CODES):
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
                cell.font=Font(bold=True,color='FFFFFF' if i==1 else '173B37')
                cell.fill=PatternFill('solid',fgColor='16817A' if i==1 else 'EAF5F2')
    sheet.freeze_panes='B2'
    sheet.auto_filter.ref=f'A1:F{len(rows)}'
    for col in 'BCDEF':sheet.column_dimensions[col].width=16
    sheet.column_dimensions['A'].width=48
    result=BytesIO();book.save(result);book.close()
    return result.getvalue()


def main():
    st.set_page_config(page_title='Matriz USE | CBUA',page_icon='📚',layout='wide')
    st.markdown('<style>[data-testid="stAppViewContainer"]{background:#f8fafb;}[data-testid="stSidebar"]{background:#f0f5f4;}.block-container{max-width:1250px;padding-top:3rem;}</style>',unsafe_allow_html=True)
    st.sidebar.markdown('### Universidad de Sevilla')
    st.title('Matriz de observaciones USE')
    st.caption('Resultados provisionales — estudios UX en curso')
    try:
        data=json.loads((Path(__file__).parent/'data/matrix.json').read_text(encoding='utf-8'))
        validate(data)
    except Exception:
        st.error('La matriz no está disponible o sus totales son incompatibles.')
        st.stop()
    st.caption('Desde el 05/10/2026 · Actualización manual UTC: '+data['updated_utc'])
    source=[]
    for row in data['rows']:
        label=row['centre'] if row['state']=='identificado' else ('Sin centro informado' if row['state']=='sin_centro' else 'Centro pendiente de identificación')
        values=[row['counts'][c] for c in CODES]
        source.append([label,*values,sum(values)])
    selected=st.sidebar.multiselect('Centros',[r[0] for r in source])
    rows=[r for r in source if not selected or r[0] in selected]
    rows.append(['TOTAL',*[sum(r[i] for r in rows) for i in range(1,6)]])
    columns=['Centro',*CODES,'Total']
    st.dataframe(pd.DataFrame(rows,columns=columns),hide_index=True,use_container_width=True,height=min(900,max(180,40+35*len(rows))))
    st.download_button('Descargar matriz Excel (.xlsx)',excel(columns,rows),file_name='Matriz_USE.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with st.expander('Criterios del recuento'):
        st.caption('Cada respuesta de Forms cuenta como una observación de su cuaderno. No se analizan tareas. '
                   'Se excluyen las respuestas anteriores al corte y las pruebas 65/66 en UX 1_1. '
                   'Los centros conservan su denominación original. Los totales suman observaciones, no personas únicas. '
                   'La actualización es manual, no en tiempo real.')


if __name__=='__main__':
    main()
