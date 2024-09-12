import win32com.client as win32
import win32

def save_excel_as_pdf(excel_path, pdf_path, row_start, row_end, col_start, col_end):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Worksheets[0]
    
    # 行と列を数値で動的に印刷範囲を設定
    start_cell = ws.Cells(row_start, col_start).Address  # 開始セル
    end_cell = ws.Cells(row_end, col_end).Address  # 終了セル
    ws.PageSetup.PrintArea = f"{start_cell}:{end_cell}"  # 印刷範囲を設定
    
    # PDFとして保存
    ws.ExportAsFixedFormat(0, pdf_path)
    wb.Close(SaveChanges=False)
    excel.Quit()

# 使用例
save_excel_as_pdf("目的_旧バージョン.xlsx", "output.pdf", 1, 10, 1, 25)  # A1:G100

