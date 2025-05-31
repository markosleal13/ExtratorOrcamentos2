import csv
import os
import platform
from io import BytesIO, StringIO
from datetime import datetime

from flask import Flask, abort, make_response, request, render_template # Adicionado render_template
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Side, Border
from openpyxl.utils import get_column_letter


# App Flask
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
        # --- CARREGAR O ARQUIVO EXCEL DadosOrcamentoConsolidado.xlsx COMO BASE DE DADOS ---
        # O arquivo está na pasta 'templates' que está no mesmo nível de 'app.py'
        excel_file_path = os.path.join(app.root_path, 'templates', 'DadosOrcamentoConsolidado.xlsx')
        
        # Verifica se o arquivo existe na pasta 'templates'
        if not os.path.exists(excel_file_path):
            abort(404, description=f"Arquivo de base de dados '{os.path.basename(excel_file_path)}' não encontrado na pasta 'templates'.")

        wb = load_workbook(excel_file_path)
        ws = wb.active # Pega a planilha ativa do arquivo Excel

        # --- 1. Modificação para a Célula B4 (Data de Referência) ---
        today_date = datetime.now().strftime("%d/%m/%Y")
        ws['B4'] = f"Data de referência: {today_date}"

        # --- Definir a Estrutura do Excel para a Lógica de Filtragem e Formatação ---
        header_row_index = 9  # Assumindo que os cabeçalhos da tabela estão na linha 9
        data_start_row = 10   # Assumindo que os dados começam na linha 10
        data_start_col = 2    # Assumindo que os dados e cabeçalhos começam na Coluna B (índice 2)

        # Mapear os cabeçalhos do Excel para seus índices de coluna
        excel_header_map = {}
        max_excel_col = ws.max_column
        for col_idx in range(data_start_col, max_excel_col + 1):
            header_cell_value = ws.cell(row=header_row_index, column=col_idx).value
            if header_cell_value is not None:
                excel_header_map[str(header_cell_value)] = col_idx

        # Mapeamento dos parâmetros da URL para os nomes EXATOS das colunas no Excel
        # ATENÇÃO: VERIFIQUE E AJUSTE ESTES NOMES PARA QUE CORRESPONDAM EXATAMENTE
        # AOS CABEÇALHOS NO SEU ARQUIVO DadosOrcamentoConsolidado.xlsx
        FILTER_MAP = {
            "ano": "ANO_REFERENCIA",
            "mes": "MES_REFERENCIA",
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

        for r_idx in range(data_start_row, ws.max_row + 1):
            row_hidden = False
            
            for param_name, excel_header_name in FILTER_MAP.items():
                param_value = filter_params.get(param_name)

                if param_value is not None:
                    col_idx = excel_header_map.get(excel_header_name)
                    
                    if col_idx is None:
                        continue 

                    excel_cell_value = ws.cell(row=r_idx, column=col_idx).value
                    cell_val_str = str(excel_cell_value).lower() if excel_cell_value is not None else ""

                    if param_value not in cell_val_str:
                        row_hidden = True
                        break 

            ws.row_dimensions[r_idx].hidden = row_hidden

        # --- 2. Ocultar Colunas de Mês e Ano ---
        col_mes_excel_idx = excel_header_map.get("MES_REFERENCIA")
        col_ano_excel_idx = excel_header_map.get("ANO_REFERENCIA")

        if col_mes_excel_idx:
            ws.column_dimensions[get_column_letter(col_mes_excel_idx)].hidden = True
        if col_ano_excel_idx:
            ws.column_dimensions[get_column_letter(col_ano_excel_idx)].hidden = True

        # --- 3. Aplicar Filtros (Triângulos) na Tabela de Dados ---
        filter_start_col_letter = get_column_letter(data_start_col)
        filter_start_row = header_row_index

        filter_end_col_letter = get_column_letter(max_excel_col) 
        filter_end_row = ws.max_row

        filter_range = f"{filter_start_col_letter}{filter_start_row}:{filter_end_col_letter}{filter_end_row}"
        ws.auto_filter.ref = filter_range

        # --- 4. Congelar Painéis ---
        ws.freeze_panes = f"{get_column_letter(data_start_col)}{header_row_index + 1}"


        xlsx_data = BytesIO()
        wb.save(xlsx_data)
        xlsx_data.seek(0)

        response = make_response(xlsx_data.getvalue())
        response.headers["Content-Disposition"] = "attachment; filename=DadosOrcamentoConsolidado.xlsx"
        response.headers["Content-type"] = (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        return response

    except Exception as e:
        abort(500, description=f"Erro ao processar o arquivo Excel: {str(e)}")

    finally:
        pass


# --- Rotas Adicionais ---
# Esta rota renderiza o arquivo HTML que você forneceu
@app.route("/seplan_or/", methods=["GET"])
@app.route("/seplan_or", methods=["GET"])
def seplan_or():
    return render_template("seplan_or.html")


# Rota raiz do aplicativo
@app.route("/", methods=["GET"])
def index():
    return "Extrator DGT - Versão Pública - Gerador de Relatórios Orçamentários."
