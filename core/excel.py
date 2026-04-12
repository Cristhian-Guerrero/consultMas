"""
Generación y formato de reportes Excel.
"""

from datetime import datetime


def apply_formatting(workbook, worksheet, df_out, tipo="COMPLETO", tipo_consulta="basica"):
    """Aplica formato profesional al Excel generado."""

    # ─── Título ───
    if tipo == "PARCIAL":
        if tipo_consulta == "basica":
            title_fmt = workbook.add_format({
                'bold': True, 'font_size': 16, 'font_color': '#D63384',
                'align': 'center', 'valign': 'vcenter', 'bg_color': '#FDF2F8', 'font_name': 'Calibri'
            })
            titulo = f'[PARCIAL - Express] Consulta DIAN Profesional - A.S. Contadores & Asesores SAS - {datetime.now().strftime("%d/%m/%Y %H:%M")}'
        else:
            title_fmt = workbook.add_format({
                'bold': True, 'font_size': 16, 'font_color': '#1565C0',
                'align': 'center', 'valign': 'vcenter', 'bg_color': '#E3F2FD', 'font_name': 'Calibri'
            })
            titulo = f'[PARCIAL - RUT DETALLADO] Consulta DIAN Profesional - A.S. Contadores & Asesores SAS - {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    else:
        if tipo_consulta == "basica":
            title_fmt = workbook.add_format({
                'bold': True, 'font_size': 16, 'font_color': '#1B5E20',
                'align': 'center', 'valign': 'vcenter', 'bg_color': '#E8F5E9', 'font_name': 'Calibri'
            })
            titulo = f'Consulta Gestión Masiva DIAN Express - A.S. Contadores & Asesores SAS - {datetime.now().strftime("%d/%m/%Y %H:%M")}'
        else:
            title_fmt = workbook.add_format({
                'bold': True, 'font_size': 16, 'font_color': '#0D47A1',
                'align': 'center', 'valign': 'vcenter', 'bg_color': '#E1F5FE', 'font_name': 'Calibri'
            })
            titulo = f'Consulta Gestión Masiva DIAN RUT DETALLADO - A.S. Contadores & Asesores SAS - {datetime.now().strftime("%d/%m/%Y %H:%M")}'

    # ─── Headers ───
    if tipo_consulta == "basica":
        header_fmt = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#1B5E20', 'font_color': 'white', 'border': 2,
            'border_color': '#2E7D32', 'font_size': 11, 'font_name': 'Calibri'
        })
    else:
        header_fmt = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#0D47A1', 'font_color': 'white', 'border': 2,
            'border_color': '#1976D2', 'font_size': 11, 'font_name': 'Calibri'
        })

    # ─── Formatos de datos ───
    base = {'valign': 'vcenter', 'border': 1, 'border_color': '#9E9E9E', 'font_size': 10, 'font_name': 'Calibri'}
    num_fmt      = workbook.add_format({**base, 'num_format': '0', 'align': 'center', 'bg_color': '#FFFFFF'})
    text_fmt     = workbook.add_format({**base, 'bg_color': '#FFFFFF'})
    text_fmt_par = workbook.add_format({**base, 'bg_color': '#EAF4EA'})
    num_fmt_par  = workbook.add_format({**base, 'num_format': '0', 'align': 'center', 'bg_color': '#EAF4EA'})
    success_fmt  = workbook.add_format({**base, 'align': 'center', 'bg_color': '#1B5E20', 'font_color': '#FFFFFF', 'bold': True, 'border_color': '#145214'})
    no_ins_fmt   = workbook.add_format({**base, 'align': 'center', 'bg_color': '#E65100', 'font_color': '#FFFFFF', 'bold': True, 'border_color': '#BF360C'})

    if tipo_consulta == "rut_detallado":
        estado_activo_fmt   = workbook.add_format({**base, 'align': 'center', 'bg_color': '#B3E5FC', 'font_color': '#0D47A1', 'bold': True, 'border_color': '#E1F5FE'})
        estado_inactivo_fmt = workbook.add_format({**base, 'align': 'center', 'bg_color': '#FFECB3', 'font_color': '#E65100', 'bold': True, 'border_color': '#E1F5FE'})

    # ─── Título y anchos ───
    worksheet.merge_range(0, 0, 0, len(df_out.columns) - 1, titulo, title_fmt)
    worksheet.set_row(0, 35)

    if tipo_consulta == "basica":
        col_widths = {
            'NIT': 15, 'DV': 5, 'Primer Apellido': 20, 'Segundo Apellido': 20,
            'Primer Nombre': 20, 'Otros Nombres': 16, 'Razón Social': 38,
            'Fecha Consulta': 19, 'Estado Consulta': 14, 'Tipo de Consulta': 10, 'Observaciones': 22
        }
    else:
        col_widths = {
            'NIT': 15, 'DV': 5, 'Primer Apellido': 20, 'Segundo Apellido': 20,
            'Primer Nombre': 20, 'Otros Nombres': 16, 'Razón Social': 32,
            'Estado del Registro': 22, 'Fecha Consulta': 19, 'Estado Consulta': 14,
            'Tipo de Consulta': 10, 'Observaciones': 22
        }

    for i, col in enumerate(df_out.columns):
        worksheet.set_column(i, i, col_widths.get(col, 15))

    # ─── Headers ───
    for col_num, col_name in enumerate(df_out.columns):
        worksheet.write(2, col_num, col_name, header_fmt)

    # ─── Datos ───
    for row_num in range(len(df_out)):
        es_par = row_num % 2 == 0
        tf = text_fmt_par if es_par else text_fmt
        nf = num_fmt_par  if es_par else num_fmt

        for col_num, col_name in enumerate(df_out.columns):
            cell_value = df_out.iloc[row_num, col_num]
            actual_row = row_num + 3

            if col_name in ['NIT', 'DV']:
                cell_fmt = nf
            elif col_name == 'Estado Consulta':
                cell_fmt = success_fmt if str(cell_value).lower() == 'exitoso' else no_ins_fmt
            elif col_name == 'Estado del Registro' and tipo_consulta == "rut_detallado":
                v = str(cell_value).upper()
                if 'ACTIVO' in v:
                    cell_fmt = estado_activo_fmt
                elif 'SUSPENDIDO' in v or 'ERROR' in v or 'SIN' in v:
                    cell_fmt = estado_inactivo_fmt
                else:
                    cell_fmt = tf
            else:
                cell_fmt = tf

            worksheet.write(actual_row, col_num, cell_value, cell_fmt)

    worksheet.freeze_panes(3, 0)
    worksheet.autofilter(2, 0, len(df_out) + 2, len(df_out.columns) - 1)
    worksheet.set_default_row(18)
    worksheet.set_row(2, 32)
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)
