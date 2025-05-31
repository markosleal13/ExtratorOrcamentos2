import csv
import os
import platform
from io import BytesIO, StringIO
from datetime import datetime

from flask import Flask, abort, make_response, request, render_template
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Side, Border
from openpyxl.utils import get_column_letter


app = Flask(__name__)

@app.route("/seplan_or_download/", methods=["GET"])
@app.route("/seplan_or_download", methods=["GET"])
def seplan_or_download():
    # Obter parâmetros de filtro da URL
    ano = request.args.get("ano", default="%")
    mes = request.args.get("mes", default="%")
    descricaoacao = request.args.get("descricaoacao", default="%")
    descricaoorcamentaria = request.args.get("descricaoorcamentaria", default="%")
    gnd = request.args.get("gnd", default="%")
    programaticaorcamentaria = request.args.get("programaticaorcamentaria", default="%")
    descricaoprograma = request.args.get("descricaoprograma", default="%")
    formato_csv = request.args.get("csv", default=None)

    try:
        excel_file_path = os.path.join(app.root_path, 'templates', 'DadosOrcamentoConsolidado (2).xlsx')
        
        if not os.path.exists(excel_file_path):
            abort(404, description=f"Arquivo de base de dados '{os.path.basename(excel_file_path)}' não encontrado na pasta 'templates'.")

        wb = load_workbook(excel_file_path)
        ws = wb.active

        # --- 1. Modificação para a Célula B4 (Data de Referência) ---
        today_date = datetime.now().strftime("%d/%m/%Y")
        ws['B4'] = f"Data de referência: {today_date}"

        # --- Definir a Estrutura do Excel para a Lógica de Filtragem e Formatação ---
        header_row_index = 9  # Assumindo que os cabeçalhos da tabela estão na linha 9
        data_start_row = 10   # Assumindo que os dados começam na linha 10
        data_start_col = 2    # Assumindo que os dados e cabeçalhos começam na Coluna B (índice 2)

        excel_header_map = {}
        max_excel_col = ws.max_column
        for col_idx in range(data_start_col, max_excel_col + 1):
            header_cell_value = ws.cell(row=header_row_index, column=col_idx).value
            if header_cell_value is not None:
                excel_header_map[str(header_cell_value)] = col_idx

        # Mapeamento dos parâmetros da URL para os nomes EXATOS das colunas no Excel
        # ATENÇÃO: Corrigido para "Ano" e "Mês" conforme sua última instrução.
        FILTER_MAP = {
            "ano": "Ano", # *** CORRIGIDO: "Ano" (sem underscore) ***
            "mes": "Mês", # *** CORRIGIDO: "Mês" (com acento e sem underscore) ***
            "descricaoacao": "Ação e Subtítulo", 
            "descricaoorcamentaria": "Descrição", 
            "gnd": "GND",
            "programaticaorcamentaria": "Programática (Programa, Ação e Subtítulo)", 
            "descricaoprograma": "Programa", 
            "codigo": "Código", 
            "codigo_fonte": "Código Fonte", 
            "descricao_fonte": "Descrição Fonte",
            "funcao_subfuncao": "Função e Subfunção",
            "esfera": "Esfera",
            "dotacao_inicial": "Dotação Inicial",
            "acrescimos": "Acréscimos",
            "decrescimos": "Decréscimos",
            "dotacao_atualizada": "Dotação Atualizada",
            "contingenciado": "Contingenciado",
            "provisao": "Provisão",
            "destaque": "Destaque",
            "dotacao_liquida": "Dotação Líquida",
            "empenhado": "Empenhado",
            "empenhado_porcento": "% Empenhado",
            "liquidado": "Liquidado",
            "liquidado_porcento": "% Liquidado",
            "pago": "Pago",
            "pago_porcento": "% Pago"
        }
        
        # --- Implementar a Lógica de Filtragem de Linhas no Excel ---
        filter_params = {
            "ano": ano.strip('%').lower() if ano != '%' else None,
            "mes": mes.strip('%').lower() if mes != '%' else None,
            "descricaoacao": descricaoacao.strip('%').lower() if descricaoacao != '%' else None,
            "descricaoorcamentaria": descricaoorcamentaria.strip('%').lower() if descricaoorcamentaria != '%' else None,
            "gnd": gnd.strip('%').lower() if gnd != '%' else None,
            "programaticaorcamentaria": programaticaorcamentaria.strip('%').lower() if programaticaorcamentaria != '%' else None,
            "descricaoprograma": descricaoprograma.strip('%').lower() if descricaoprograma != '%' else None,
        }

        # Definir estilos para aplicação de bordas e fonte nas células de dados
        thin_border = Border(left=Side(style='thin'), 
                             right=Side(style='thin'), 
                             top=Side(style='thin'), 
                             bottom=Side(style='thin'))
        data_font = Font(name="Arial", size=8, color="000000") # Fonte padrão para os dados

        for r_idx in range(data_start_row, ws.max_row + 1):
            row_hidden = False
            
            for param_name, excel_header_name in FILTER_MAP.items():
                param_value = filter_params.get(param_name)

                if param_value is not None:
                    col_idx = excel_header_map.get(excel_header_name)
                    
                    if col_idx is None:
                        continue 

                    excel_cell_value = ws.cell(row=r_idx, column=col_idx).value
                    cell_val_str = str(excel_cell_value).strip().lower() if excel_cell_value is not None else ""

                    is_match = False
                    # --- Lógica de filtro para Ano e Mês (revertida para string exata) ---
                    if param_name in ["ano", "mes"]: 
                        if param_value == cell_val_str: # Comparação de string exata (após strip().lower())
                            is_match = True
                    else: # Para outros campos de texto: lógica de "contém"
                        if param_value in cell_val_str:
                            is_match = True
                    
                    if not is_match:
                        row_hidden = True
                        break 

            ws.row_dimensions[r_idx].hidden = row_hidden
            
            # --- APLICAR FORMATAÇÃO BÁSICA (Bordas e Fonte) para Linhas VISÍVEIS ---
            # Isso garante que as grades e a fonte sejam aplicadas consistentemente.
            if not row_hidden:
                for c_idx in range(data_start_col, max_excel_col + 1): # Itera por todas as colunas de dados
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.border = thin_border # Aplica as bordas
                    cell.font = data_font    # Aplica a fonte

                    # Opcional: ajustar alinhamento para números se necessário
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")


        # --- Aplicar Filtros (Triângulos) na Tabela de Dados ---
        filter_start_col_letter = get_column_letter(data_start_col)
        filter_start_row = header_row_index

        filter_end_col_letter = get_column_letter(max_excel_col) 
        filter_end_row = ws.max_row

        filter_range = f"{filter_start_col_letter}{filter_start_row}:{filter_end_col_letter}{filter_end_row}"
        ws.auto_filter.ref = filter_range

        # --- Congelar Painéis ---
        ws.freeze_panes = f"{get_column_letter(data_start_col)}{header_row_index + 1}"


        xlsx_data = BytesIO()
        wb.save(xlsx_data)
        xlsx_data.seek(0)

        response = make_response(xlsx_data.getvalue())
        response.headers["Content-Disposition"] = "attachment; filename=DadosOrcamentoConsolidado (2).xlsx"
        response.headers["Content-type"] = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        return response

    except Exception as e:
        abort(500, description=f"Erro ao processar o arquivo Excel: {str(e)}")

    finally:
        pass


@app.route("/seplan_or/", methods=["GET"])
@app.route("/seplan_or", methods=["GET"])
def seplan_or():
    return render_template("seplan_or.html")


@app.route("/", methods=["GET"])
def index():
    return "Extrator DGT - Versão Pública - Gerador de Relatórios Orçamentários."
