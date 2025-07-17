import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
from datetime import datetime
import locale

# 日本語フォントの設定（Windows環境）
try:
    # Windows標準の日本語フォントを使用
    pdfmetrics.registerFont(TTFont('MSGothic', 'C:/Windows/Fonts/msgothic.ttc'))
    FONT_NAME = 'MSGothic'
except:
    # フォントが見つからない場合はデフォルトを使用
    FONT_NAME = 'Helvetica'

class SalaryPayslipGenerator:
    def __init__(self, excel_file_path):
        """
        給与明細PDF生成器の初期化
        """
        self.excel_file_path = excel_file_path
        self.workbook_values = None
        self.workbook_formulas = None
        self.load_excel_data()
        
    def load_excel_data(self):
        """
        エクセルファイルから給与データを読み込む（数式と計算結果の両方を取得）
        """
        try:
            # 計算結果を取得
            self.workbook_values = openpyxl.load_workbook(self.excel_file_path, data_only=True)
            # 数式を取得
            self.workbook_formulas = openpyxl.load_workbook(self.excel_file_path, data_only=False)
            print(f"エクセルファイル '{self.excel_file_path}' を読み込みました（数式と計算結果を取得）。")
        except Exception as e:
            print(f"エクセルファイルの読み込みエラー: {e}")
            raise
    
    def get_cell_value_with_formula(self, sheet_values, sheet_formulas, row, col):
        """
        セルの値を取得（数式の場合は数式を表示）
        """
        value = sheet_values.cell(row=row, column=col).value
        formula = sheet_formulas.cell(row=row, column=col).value
        
        # 値がNoneで数式がある場合は数式を返す
        if value is None and formula and str(formula).startswith('='):
            return f"[数式: {formula}]"
        elif value is None:
            return ""
        else:
            return value
    
    def get_salary_data(self, employee_name, target_month):
        """
        指定された従業員と月の給与データを取得
        """
        salary_sheet_values = self.workbook_values.active  # アクティブシートを使用
        salary_sheet_formulas = self.workbook_formulas.active  # アクティブシートを使用
        
        # ヘッダー行を取得
        headers = []
        for col in range(1, salary_sheet_values.max_column + 1):
            cell_value = salary_sheet_values.cell(row=1, column=col).value
            headers.append(str(cell_value) if cell_value else "")
        
        # 指定された従業員と月のデータを検索
        employee_data = None
        for row in range(2, salary_sheet_values.max_row + 1):
            name_cell = salary_sheet_values.cell(row=row, column=3).value  # 氏名列（3列目）
            date_cell = salary_sheet_values.cell(row=row, column=1).value  # 支給日列（1列目）
            
            if name_cell == employee_name and date_cell:
                # 日付が指定された月と一致するかチェック
                if hasattr(date_cell, 'month') and date_cell.month == target_month:
                    employee_data = {}
                    for col in range(1, salary_sheet_values.max_column + 1):
                        cell_value = self.get_cell_value_with_formula(
                            salary_sheet_values, salary_sheet_formulas, row, col
                        )
                        header = headers[col-1] if col-1 < len(headers) else f"列{col}"
                        employee_data[header] = cell_value
                    break
        
        return employee_data, headers
    
    def get_employee_info(self, employee_name):
        """
        従業員の基本情報を取得
        """
        personal_sheet = self.workbook_values.active  # アクティブシートを使用
        
        employee_info = {}
        for row in range(2, personal_sheet.max_row + 1):
            name_cell = personal_sheet.cell(row=row, column=3).value  # 氏名列（3列目）
            if name_cell == employee_name:
                employee_info['社員番号'] = personal_sheet.cell(row=row, column=2).value  # 社員番号（2列目）
                employee_info['氏名'] = name_cell
                employee_info['生年月日'] = personal_sheet.cell(row=row, column=1).value  # 支給日（1列目）
                break
        
        return employee_info
    
    def get_available_employees(self):
        """
        利用可能な従業員のリストを取得
        """
        personal_sheet = self.workbook_values.active  # アクティブシートを使用
        employees = []
        
        for row in range(2, personal_sheet.max_row + 1):
            name_cell = personal_sheet.cell(row=row, column=3).value  # 氏名列（3列目）
            if name_cell and name_cell.strip():  # 空でない名前のみ
                employees.append(name_cell)
        
        return employees
    
    def get_employee_salary_months(self, employee_name):
        """
        指定された従業員の給与データがある月を取得
        """
        salary_sheet = self.workbook_values.active  # アクティブシートを使用
        months = set()
        
        for row in range(2, salary_sheet.max_row + 1):
            name_cell = salary_sheet.cell(row=row, column=3).value  # 氏名列（3列目）
            date_cell = salary_sheet.cell(row=row, column=1).value  # 支給日列（1列目）
            
            if name_cell == employee_name and date_cell and hasattr(date_cell, 'month'):
                months.add(date_cell.month)
        
        return sorted(months)
    
    def format_currency_value(self, value):
        """
        通貨値を適切にフォーマットする
        """
        if value is None:
            return ""
        elif isinstance(value, str) and value.startswith('[数式:'):
            return value  # 数式の場合はそのまま表示
        elif isinstance(value, (int, float)):
            return f"{value:,}"
        else:
            return str(value)
    
    def create_payslip_pdf(self, employee_name, target_month, output_path):
        """
        給与明細PDFを作成
        """
        # 給与データを取得
        salary_data, headers = self.get_salary_data(employee_name, target_month)
        if not salary_data:
            print(f"従業員 '{employee_name}' の {target_month}月の給与データが見つかりません。")
            return False
        
        # 従業員基本情報を取得
        employee_info = self.get_employee_info(employee_name)
        
        # PDFドキュメントを作成
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            rightMargin=20*mm,
            leftMargin=20*mm,
            topMargin=20*mm,
            bottomMargin=20*mm
        )
        
        # スタイルを設定
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            fontName=FONT_NAME,
            alignment=1,  # 中央揃え
            spaceAfter=20
        )
        
        normal_style = ParagraphStyle(
            'Normal',
            parent=styles['Normal'],
            fontName=FONT_NAME,
            fontSize=10
        )
        
        # ストーリー（PDFの内容）を作成
        story = []
        
        # タイトル
        title = Paragraph(f"給与明細書", title_style)
        story.append(title)
        story.append(Spacer(1, 10))
        
        # 基本情報テーブル
        basic_info_data = [
            ['項目', '内容'],
            ['氏名', employee_info.get('氏名', '')],
            ['社員番号', str(employee_info.get('社員番号', ''))],
            ['生年月日', str(employee_info.get('生年月日', '')) if employee_info.get('生年月日') else ''],
            ['支給年月', f"{salary_data.get('支給日', '').year}年{target_month}月"],
            ['支給日', str(salary_data.get('支給日', ''))[:10]]
        ]
        
        basic_table = Table(basic_info_data, colWidths=[60*mm, 100*mm])
        basic_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(basic_table)
        story.append(Spacer(1, 20))
        
        # 給与詳細テーブル（縦書き、2列表示）
        salary_details_data = [
            ['項目', '金額']
        ]
        
        # 主要な給与項目を抽出（実際のExcelファイルの列名に合わせる）
        key_items = [
            '総支給額',
            '標準報酬\n月額', 
            '健康保険',
            '厚生年金',
            '社会保険料\n控除後',
            '源泉所得税',
            '差引支給額',
            '振込金額',
        ]
        
        for item in key_items:
            value = salary_data.get(item, '')
            salary_details_data.append([item, self.format_currency_value(value)])
        
        salary_table = Table(salary_details_data, colWidths=[80*mm, 60*mm])
        salary_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(salary_table)
        story.append(Spacer(1, 20))
        
        # 控除内訳テーブル（縦書き、2列表示）
        deduction_data = [
            ['控除項目', '金額']
        ]
        
        deduction_items = [
            '健康保険料（従業員）',
            '厚生年金（従業員）',
            '社会保険料\n控除額',
            '扶養親族\n等の数',
        ]
        
        for item in deduction_items:
            value = salary_data.get(item, '')
            deduction_data.append([item, self.format_currency_value(value)])
        
        deduction_table = Table(deduction_data, colWidths=[80*mm, 60*mm])
        deduction_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), FONT_NAME),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(deduction_table)
        story.append(Spacer(1, 30))
        
        # 備考欄
        note_style = ParagraphStyle(
            'Note',
            parent=styles['Normal'],
            fontName=FONT_NAME,
            fontSize=9,
            leftIndent=20
        )
        note = Paragraph("※ この明細書は給与計算システムにより自動生成されています。", note_style)
        story.append(note)
        
        # PDFを生成（既存ファイルは上書き）
        try:
            doc.build(story)
            print(f"給与明細PDFを '{output_path}' に作成しました。")
            return True
        except Exception as e:
            print(f"PDF作成エラー: {e}")
            return False
    
    def close(self):
        """
        ワークブックを閉じる
        """
        if self.workbook_values:
            self.workbook_values.close()
        if self.workbook_formulas:
            self.workbook_formulas.close()

def main():
    """
    メイン関数
    """
    excel_file = "給与支給一覧令和7年.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"エラー: ファイル '{excel_file}' が見つかりません。")
        return
    
    # PDF生成器を作成
    generator = SalaryPayslipGenerator(excel_file)
    
    try:
        # 利用可能な従業員を確認
        employees = generator.get_available_employees()
        print(f"利用可能な従業員: {employees}")
        
        # 各従業員の給与データがある月を確認
        print("\n各従業員の給与データ月:")
        for employee in employees:
            months = generator.get_employee_salary_months(employee)
            print(f"{employee}: {months}月")
        
        # サンプルとして渡邉俊行さんの1月の給与明細を作成
        employee_name = "渡邉俊行"
        target_month = 1
        
        output_filename = f"給与明細_{employee_name}_{target_month}月.pdf"
        
        success = generator.create_payslip_pdf(employee_name, target_month, output_filename)
        
        if success:
            print(f"\n給与明細PDFが正常に作成されました: {output_filename}")
        else:
            print("PDFの作成に失敗しました。")
            
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        generator.close()

if __name__ == "__main__":
    main() 