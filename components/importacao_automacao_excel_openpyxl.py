"""IMPORTAÇÕES PARA AUTOMAÇÃO DE PLANILHAS COM EXCEL"""
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, NamedStyle
from openpyxl.utils.exceptions import InvalidFileException

def carrega_excel(arquivoExcel):
    try:
        workbook = load_workbook(arquivoExcel)
        style_moeda = NamedStyle(name="estilo_moeda", number_format='_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-')
        if "estilo_moeda" not in workbook.named_styles:
            workbook.add_named_style(style_moeda)
        sheet = workbook.active
        return workbook, sheet, style_moeda
    except FileNotFoundError:
        print(f"Erro: O arquivo '{arquivoExcel}' não foi encontrado.")
        return None, None, None
    except InvalidFileException:
        print(f"Erro: O arquivo '{arquivoExcel}' está corrompido ou não é um arquivo Excel válido.")
        return None, None, None
    except Exception as e:
        print(f"Erro inesperado: {e}")
        return None, None, None