linhas = []
    grande_total_valor = 0.0
    grande_total_devido_valor = 0.0
    total_blocks = len(blocks)
    primeira_empresa_info = None

    try:
        start_date = datetime.datetime.strptime(start_date_str, "%d/%m/%Y") if start_date_str else None
        end_date = datetime.datetime.strptime(end_date_str, "%d/%m/%Y") if end_date_str else None
    except ValueError:
        raise ValueError("Formato de data inválido. Use DD/MM/AAAA.")

    for i, block in enumerate(blocks):
        enterprise = re.search(r"Cliente:\s*(.+?)(?=\s+CPF\b)", block)
        enterprise_code = re.search(r"\d+", enterprise.group(1)).group() if enterprise else None
        nome_empresa = codigo_empresas.get(enterprise_code, "CÓDIGO DESCONHECIDO")

        if i == 0 and enterprise_code:
            primeira_empresa_info = f"{enterprise_code} - {nome_empresa}"

        process = re.search(r"^\s*(\d+)", block)
        cnpj = re.search(r"CPF/CNPJ\s*:\s*([\d./-]+)", block)
        name_raw = re.search(r"Nome:\s*(.+?)(?:Endereço:|CPF/CNPJ:|Telefone:|Cobrador:|Situação:)", block)
        name = name_raw.group(1).strip() if name_raw else (enterprise.group(1).strip() if enterprise else "")
        cobrador_raw = re.search(r"Cobrador\s*:\s*(.+?)\s+Telefone:", block)
        cobrador = " ".join(cobrador_raw.group(1).split()[:2]).upper() if cobrador_raw else ""
        situation = re.search(r"Situação:\s*(.+?)\s+(?:Cliente:|CPF/CNPJ)", block)
        titles = re.findall(r"([Dd][Pp]|[Pp][Rr])\s+([\d/]+).*?(\d{2}/\d{2}/\d{4}).*?([\d.,]+)", block)
        tipos_lista = ", ".join(sorted({t[0].upper() for t in titles}))
        title_list = ", ".join([t[1] for t in titles])
        title_date = ", ".join([t[2] for t in titles])
        value_list = ", ".join([str(clean_and_convert_to_float(t[3])) for t in titles])

        total_matches = re.findall(r"Total do Devedor:.*?([\d.,]+)", block)
        total_value_str = total_matches[-1] if total_matches else ""
        total_devido_float = clean_and_convert_to_float(total_value_str)
        grande_total_devido_valor += total_devido_float

        base_data = {
            "USUÁRIO": enterprise_code, "EMPRESA COB.": nome_empresa,
            "PROCESSO": process.group(1) if process else "",
            "NOME CLIENTE": name.upper(), "CPF/CNPJ": cnpj.group(1) if cnpj else "",
            "ESPÉCIE": tipos_lista, "TÍTULOS COB.": title_list,
            "DATA VENC.": title_date, "VALORES TIT.": None,
            "TOTAL DEVIDO": total_devido_float,
            "SITUAÇÃO": situation.group(1).strip().upper() if situation else "",
            "COBRADOR": cobrador,
        }

        pc_data = []
        pc_header_pattern = r"Prestação de Contas\s+Nº da PC\s+Dt\. Pgto\."
        pc_section_search = re.search(pc_header_pattern, block, re.IGNORECASE)

        if pc_section_search:
            pc_text_block = block[pc_section_search.start():]
            pc_row_pattern = re.compile(
                r"(\d+)\s+(\d{2}/\d{2}/\d{4})\s+(R\$\s*[\d.,]+)\s+(R\$\s*[\d.,]+)\s+(R\$\s*[\d.,]+)\s+(R\$\s*[\d.,]+)")
            pc_data_raw = pc_row_pattern.findall(pc_text_block)

            if start_date or end_date:
                for entry in pc_data_raw:
                    try:
                        payment_date = datetime.datetime.strptime(entry[1], "%d/%m/%Y")
                        if (not start_date or payment_date >= start_date) and \
                                (not end_date or payment_date <= end_date):
                            pc_data.append(entry)
                    except ValueError:
                        continue
            else:
                pc_data = pc_data_raw

        total_depositado_valor = 0.0
        if pc_data:
            for pc_entry in pc_data:
                total_depositado_valor += clean_and_convert_to_float(pc_entry[5])

        grande_total_valor += total_depositado_valor

        if pc_data:
            first_pc_entry = pc_data[0]
            first_row = base_data.copy()
            first_row.update({
                "Nº DA PC": first_pc_entry[0].strip(),
                "DATA PGTO": first_pc_entry[1].strip(),
                "VALOR CRÉDITO": clean_and_convert_to_float(first_pc_entry[2]),
                "TAXA": clean_and_convert_to_float(first_pc_entry[3]),
                "VALORES DEP.": clean_and_convert_to_float(first_pc_entry[5]),
                "TOTAL CLIENTE": total_depositado_valor,
            })
            linhas.append(first_row)

            for subsequent_pc_entry in pc_data[1:]:
                empty_row = {key: "" for key in ALL_COLUMNS}
                empty_row.update({
                    "Nº DA PC": subsequent_pc_entry[0].strip(),
                    "DATA PGTO": subsequent_pc_entry[1].strip(),
                    "VALOR CRÉDITO": clean_and_convert_to_float(subsequent_pc_entry[2]),
                    "TAXA": clean_and_convert_to_float(subsequent_pc_entry[3]),
                    "VALORES DEP.": clean_and_convert_to_float(subsequent_pc_entry[5]),
                })
                linhas.append(empty_row)
        else:
            if not start_date and not end_date:
                row_data = base_data.copy()
                row_data.update(
                    {"Nº DA PC": "", "DATA PGTO": "", "VALOR CRÉDITO": None, "TAXA": None, "VALORES DEP.": None,
                     "TOTAL CLIENTE": None})
                linhas.append(row_data)

        if gui_instance:
            time.sleep(0.01)
            progress_percent = int(((i + 1) / total_blocks) * 95)
            gui_instance.update_progress(progress_percent)

    df = pd.DataFrame(linhas, columns=ALL_COLUMNS)
    df = df[selected_columns]

    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Relatório', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Relatório']

    currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})
    bold_format = workbook.add_format({'bold': True})

    header = df.columns.values.tolist()
    currency_cols = ["TOTAL DEVIDO", "VALOR CRÉDITO", "TAXA", "VALORES DEP.", "TOTAL CLIENTE"]

    for i, col_name in enumerate(header):
        column_len = max(df[col_name].astype(str).map(len).max(), len(col_name))
        adjusted_width = column_len + 2

        if col_name in currency_cols:
            worksheet.set_column(i, i, adjusted_width, currency_format)
        else:
            worksheet.set_column(i, i, adjusted_width)

    if (grande_total_valor > 0 or grande_total_devido_valor > 0):
        last_row = len(df) + 2
        summary_col_label = max(0, len(selected_columns) - 2)
        summary_col_value = max(1, len(selected_columns) - 1)

        worksheet.write(last_row, summary_col_label, "EMPRESA DE COBRANÇA:", bold_format)
        worksheet.write(last_row, summary_col_value, primeira_empresa_info)
        last_row += 1

        worksheet.write(last_row, summary_col_label, "DATA DE EXTRAÇÃO:", bold_format)
        worksheet.write(last_row, summary_col_value, f"{datetime.datetime.now():%d/%m/%Y %H:%M}")
        last_row += 1

        if start_date_str or end_date_str:
            periodo_str = f"{start_date_str} a {end_date_str}" if start_date_str and end_date_str else f"A partir de {start_date_str}" if start_date_str else f"Até {end_date_str}"
            worksheet.write(last_row, summary_col_label, "PERÍODO DO FILTRO:", bold_format)
            worksheet.write(last_row, summary_col_value, periodo_str)
            last_row += 1

        worksheet.write(last_row, summary_col_label, "TOTAL ARRECADADO:", bold_format)
        worksheet.write(last_row, summary_col_value, grande_total_valor, currency_format)
        last_row += 1

        worksheet.write(last_row, summary_col_label, "TOTAL EM COBRANÇA:", bold_format)
        worksheet.write(last_row, summary_col_value, grande_total_devido_valor, currency_format)

    writer.close()
