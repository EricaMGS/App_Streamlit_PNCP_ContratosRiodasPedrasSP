# 2. GERAÇÃO DO RELATÓRIO ESTILO WORD (.docx)
                doc = Document()
                
                # Cabeçalho / Título
                p_title = doc.add_paragraph()
                run_title = p_title.add_run(f"Relatório Executivo: {tipo_consulta}")
                run_title.bold = True
                run_title.font.size = Pt(18)
                run_title.font.color.rgb = RGBColor(0, 51, 102)

                doc.add_paragraph(f"Município: Prefeitura Municipal de Rio das Pedras / SP")
                doc.add_paragraph(f"Período consultado: {data_inicio.strftime('%d/%m/%Y')} até {data_fim.strftime('%d/%m/%Y')}")
                doc.add_paragraph(f"Total de registros encontrados: {len(df)}")
                if col_valor:
                    doc.add_paragraph(f"Valor Consolidado: R$ {total_valor:,.2f}")

                doc.add_heading("Detalhes dos Registros", level=2)

                # Criação da tabela no Word com todas as colunas do dataframe
                # Usamos 'df.columns' para criar o cabeçalho dinamicamente
                table = doc.add_table(rows=1, cols=len(df.columns))
                table.style = 'Table Grid'
                hdr_cells = table.rows[0].cells
                
                for i, column_name in enumerate(df.columns):
                    hdr_cells[i].text = str(column_name)

                # Preenchimento das linhas (limitado a 50 linhas para não travar o Word)
                for index, row in df.head(50).iterrows():
                    row_cells = table.add_row().cells
                    for i, column_name in enumerate(df.columns):
                        row_cells[i].text = str(row[column_name])

                doc.add_paragraph("\nRelatório gerado automaticamente.")

                buffer_word = io.BytesIO()
                doc.save(buffer_word)
                buffer_word.seek(0)

                with col_dl2:
                    st.download_button(
                        label="📝 Baixar Relatório em Word (.docx)",
                        data=buffer_word.getvalue(),
                        file_name=f"Relatorio_{tipo_consulta.replace(' ', '_')}_Rio_Das_Pedras.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
