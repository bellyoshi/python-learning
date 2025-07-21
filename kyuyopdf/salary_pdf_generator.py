#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
給与明細PDF生成器
Excelファイルから給与データを読み込み、給与明細PDFを生成
"""

import openpyxl
import os
from datetime import datetime
from decimal import Decimal
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from personal_salary_data import PersonalSalaryData, PersonalSalaryDataManager

class SalaryPayslipGenerator:
    """給与明細PDF生成器"""
    
    def __init__(self, excel_file_path: str):
        """
        初期化
        
        Args:
            excel_file_path: Excelファイルのパス
        """
        self.excel_file_path = excel_file_path
        self.manager = PersonalSalaryDataManager()
        self.load_data()
    
    def load_data(self):
        """エクセルファイルからデータを読み込み"""
        try:
            self.manager.load_from_excel_file(self.excel_file_path)
            print(f"データ読み込み完了: {len(self.manager.get_all_personal_data())} 件")
        except Exception as e:
            print(f"データ読み込みエラー: {e}")
    
    def get_available_employees(self) -> list:
        """利用可能な従業員のリストを取得"""
        return self.manager.get_available_employees()
    
    def get_employee_salary_months(self, employee_name: str) -> list:
        """指定された従業員の給与データがある月を取得"""
        employee_data = self.manager.get_personal_data_by_employee(employee_name)
        months = []
        for data in employee_data:
            if data.支給日:
                months.append(data.支給日.month)
        return sorted(list(set(months)))
    
    def get_salary_data(self, employee_name: str, target_month: int) -> PersonalSalaryData:
        """指定された従業員と月の給与データを取得"""
        employee_data = self.manager.get_personal_data_by_employee_and_month(employee_name, target_month)
        if employee_data:
            return employee_data[0]
        return None
    
    def create_payslip_pdf(self, employee_name: str, target_month: int, output_path: str) -> bool:
        """
        給与明細PDFを作成
        
        Args:
            employee_name: 従業員名
            target_month: 対象月
            output_path: 出力ファイルパス
            
        Returns:
            成功した場合はTrue、失敗した場合はFalse
        """
        try:
            # 給与データを取得
            salary_data = self.get_salary_data(employee_name, target_month)
            if not salary_data:
                print(f"従業員 '{employee_name}' の {target_month}月の給与データが見つかりません。")
                return False
            
            # PDFを作成
            self._create_pdf(salary_data, target_month, output_path)
            return True
            
        except Exception as e:
            print(f"PDF作成エラー: {e}")
            return False
    
    def _create_pdf(self, salary_data: PersonalSalaryData, target_month: int, output_path: str):
        """PDFファイルを作成"""
        # PDFキャンバスを作成
        c = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        
        # フォント設定
        try:
            # Windows標準フォントを使用
            pdfmetrics.registerFont(TTFont('MSGothic', 'C:/Windows/Fonts/msgothic.ttc'))
            font_name = 'MSGothic'
        except:
            # フォールバック
            font_name = 'Helvetica'
        
        # タイトル
        c.setFont(font_name, 16)
        c.drawString(50*mm, 270*mm, f"給与明細書")
        
        # 基本情報
        c.setFont(font_name, 12)
        y_position = 250*mm
        
        # 従業員名
        c.drawString(50*mm, y_position, f"氏名: {salary_data.氏名 or 'N/A'}")
        y_position -= 15*mm
        
        # 支給日
        if salary_data.支給日:
            payment_date = salary_data.支給日.strftime("%Y年%m月%d日")
        else:
            payment_date = "N/A"
        c.drawString(50*mm, y_position, f"支給日: {payment_date}")
        y_position -= 15*mm
        
        # 給与項目
        c.setFont(font_name, 10)
        
        # 総支給額
        if salary_data.総支給額 is not None:
            total_salary = f"{salary_data.総支給額:,.0f}円"
        else:
            total_salary = "N/A"
        c.drawString(50*mm, y_position, f"総支給額: {total_salary}")
        y_position -= 12*mm
        
        # 標準報酬月額
        if salary_data.標準報酬月額 is not None:
            standard_reward = f"{salary_data.標準報酬月額:,.0f}円"
        else:
            standard_reward = "N/A"
        c.drawString(50*mm, y_position, f"標準報酬月額: {standard_reward}")
        y_position -= 12*mm
        
        # 健康保険（従業員負担分）
        if salary_data.健康保険料_従業員 is not None:
            health_insurance = f"{salary_data.健康保険料_従業員:,.0f}円"
        else:
            health_insurance = "N/A"
        c.drawString(50*mm, y_position, f"健康保険料: {health_insurance}")
        y_position -= 12*mm
        
        # 厚生年金（従業員負担分）
        if salary_data.厚生年金_従業員 is not None:
            pension = f"{salary_data.厚生年金_従業員:,.0f}円"
        else:
            pension = "N/A"
        c.drawString(50*mm, y_position, f"厚生年金: {pension}")
        y_position -= 12*mm
        
        # 社会保険料控除後
        if salary_data.社会保険料控除後 is not None:
            after_social = f"{salary_data.社会保険料控除後:,.0f}円"
        else:
            after_social = "N/A"
        c.drawString(50*mm, y_position, f"社会保険料控除後: {after_social}")
        y_position -= 12*mm
        
        # 源泉所得税
        if salary_data.源泉所得税 is not None:
            income_tax = f"{salary_data.源泉所得税:,.0f}円"
        else:
            income_tax = "N/A"
        c.drawString(50*mm, y_position, f"源泉所得税: {income_tax}")
        y_position -= 12*mm
        
        # 賃料控除（金額がある場合のみ表示）
        if salary_data.賃料控除 is not None and salary_data.賃料控除 > 0:
            rent_deduction = f"{salary_data.賃料控除:,.0f}円"
            c.drawString(50*mm, y_position, f"賃料控除: {rent_deduction}")
            y_position -= 12*mm
        
        # 駐車場控除（金額がある場合のみ表示）
        if salary_data.駐車場控除 is not None and salary_data.駐車場控除 > 0:
            parking_deduction = f"{salary_data.駐車場控除:,.0f}円"
            c.drawString(50*mm, y_position, f"駐車場控除: {parking_deduction}")
            y_position -= 12*mm
        
        # 振込金額
        if salary_data.振込金額 is not None:
            transfer_amount = f"{salary_data.振込金額:,.0f}円"
        else:
            transfer_amount = "N/A"
        c.drawString(50*mm, y_position, f"振込金額: {transfer_amount}")
        y_position -= 12*mm
        
        # 扶養親族等の数
        if salary_data.扶養親族等の数 is not None:
            dependents = f"{salary_data.扶養親族等の数}人"
        else:
            dependents = "N/A"
        c.drawString(50*mm, y_position, f"扶養親族等の数: {dependents}")
        
        # PDFを保存
        c.save()
        print(f"PDF作成完了: {output_path}")
    
    def close(self):
        """リソースを解放"""
        # 特に必要な処理はないが、将来的な拡張のために残す
        pass

def main():
    """メイン関数"""
    excel_file = "給与支給一覧令和7年.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"エラー: ファイル '{excel_file}' が見つかりません。")
        return
    
    # 利用可能な従業員を確認
    generator = SalaryPayslipGenerator(excel_file)
    employees = generator.get_available_employees()
    
    print(f"利用可能な従業員: {employees}")
    
    # 各従業員の給与データがある月を確認
    print("\n各従業員の給与データ月:")
    for employee in employees:
        months = generator.get_employee_salary_months(employee)
        print(f"{employee}: {months}月")
    
    # 特定の従業員と月の給与明細を作成
    if employees:
        test_employee = employees[0]
        test_months = generator.get_employee_salary_months(test_employee)
        
        if test_months:
            test_month = test_months[0]
            output_path = f"給与明細_{test_employee}_{test_month}月.pdf"
            
            success = generator.create_payslip_pdf(test_employee, test_month, output_path)
            if success:
                print(f"\nテストPDF作成成功: {output_path}")
            else:
                print(f"\nテストPDF作成失敗")
    
    generator.close()

if __name__ == "__main__":
    main()
    




