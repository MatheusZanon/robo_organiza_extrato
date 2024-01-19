"""IMPORTAÇÕES PARA AUTOMAÇÃO DE PLANILHAS COM EXCEL"""
from openpyxl import load_workbook

def auto_excel(arquivoExcel):
    wb = load_workbook(filename=arquivoExcel, read_only=True)
    ws = wb['2023']
    return ws