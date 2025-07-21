#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
個人給与データクラスのテストファイル
"""

import os
from salary_pdf_generator import SalaryPayslipGenerator
from personal_salary_data import PersonalSalaryData, PersonalSalaryDataFactory, PersonalSalaryDataManager
from decimal import Decimal

def test_personal_salary_data():
    """個人給与データクラスのテスト"""
    
    excel_file = "給与支給一覧令和7年.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"エラー: ファイル '{excel_file}' が見つかりません。")
        return
    
    # PDF生成器を作成
    generator = SalaryPayslipGenerator(excel_file)
    
    try:
        print("=== 個人給与データクラスのテスト ===")
        
        # テスト用の従業員と月
        test_employee = "渡邉俊行"
        test_month = 1
        
        # 既存の方法でデータを取得
        salary_data, headers = generator.get_salary_data(test_employee, test_month)
        employee_info = generator.get_employee_info(test_employee)
        
        print(f"従業員: {test_employee}")
        print(f"月: {test_month}")
        
        if salary_data and employee_info:
            # 個人給与データを作成
            personal_data = PersonalSalaryDataFactory.create_from_salary_data(salary_data, employee_info)
            
            print(f"\n=== 個人給与データの内容 ===")
            print(f"支給日: {personal_data.支給日}")
            print(f"氏名: {personal_data.氏名}")
            print(f"社員番号: {personal_data.社員番号}")
            print(f"総支給額: {personal_data.総支給額}")
            print(f"標準報酬月額: {personal_data.標準報酬月額}")
            print(f"健康保険: {personal_data.健康保険}")
            print(f"厚生年金: {personal_data.厚生年金}")
            print(f"社会保険料控除後: {personal_data.社会保険料控除後}")
            print(f"源泉所得税: {personal_data.源泉所得税}")
            print(f"差引支給額: {personal_data.差引支給額}")
            print(f"振込金額: {personal_data.振込金額}")
            print(f"健康保険料_従業員: {personal_data.健康保険料_従業員}")
            print(f"厚生年金_従業員: {personal_data.厚生年金_従業員}")
            print(f"社会保険料控除額: {personal_data.社会保険料控除額}")
            print(f"扶養親族等の数: {personal_data.扶養親族等の数}")
            print(f"生年月日: {personal_data.生年月日}")
            
            # 個人給与データマネージャーのテスト
            print(f"\n=== 個人給与データマネージャーのテスト ===")
            manager = PersonalSalaryDataManager()
            
            # 複数月のデータを追加
            for month in range(1, 4):  # 1月から3月
                month_salary_data, _ = generator.get_salary_data(test_employee, month)
                if month_salary_data:
                    month_personal_data = PersonalSalaryDataFactory.create_from_salary_data(
                        month_salary_data, employee_info
                    )
                    manager.add_personal_data(month_personal_data)
                    print(f"{month}月のデータを追加: 総支給額 {month_personal_data.総支給額}")
            
            # 従業員別データ取得
            employee_data_list = manager.get_by_employee_name(test_employee)
            print(f"\n従業員 '{test_employee}' のデータ数: {len(employee_data_list)}")
            
            # 月別データ取得
            for month in range(1, 4):
                month_data = manager.get_by_employee_and_month(test_employee, month)
                if month_data:
                    print(f"{month}月のデータ: 総支給額 {month_data.総支給額}")
            
            # 総支給額の合計
            total_salary = manager.get_total_salary_by_employee(test_employee)
            print(f"\n従業員 '{test_employee}' の総支給額合計: {total_salary}")
            
            # 利用可能な従業員と月
            available_employees = manager.get_available_employees()
            available_months = manager.get_available_months()
            print(f"利用可能な従業員: {available_employees}")
            print(f"利用可能な月: {available_months}")
            
            # 辞書形式でのエクスポート
            export_data = manager.export_to_dict()
            print(f"\nエクスポートデータ数: {len(export_data)}")
            if export_data:
                print(f"最初のデータ: {export_data[0]}")
        
        else:
            print("データが見つかりませんでした。")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        generator.close()

def test_empty_personal_data():
    """空の個人給与データのテスト"""
    print(f"\n=== 空の個人給与データのテスト ===")
    
    # 空のデータを作成
    empty_data = PersonalSalaryDataFactory.create_empty()
    print(f"空のデータ: {empty_data}")
    
    # 手動でデータを設定
    empty_data.氏名 = "テスト太郎"
    empty_data.社員番号 = "001"
    empty_data.総支給額 = Decimal("300000")
    empty_data.健康保険料_従業員 = Decimal("15000")
    empty_data.厚生年金_従業員 = Decimal("25000")
    
    print(f"設定後のデータ:")
    print(f"  氏名: {empty_data.氏名}")
    print(f"  社員番号: {empty_data.社員番号}")
    print(f"  総支給額: {empty_data.総支給額}")
    print(f"  健康保険料_従業員: {empty_data.健康保険料_従業員}")
    print(f"  厚生年金_従業員: {empty_data.厚生年金_従業員}")

if __name__ == "__main__":
    print("個人給与データクラスのテストを開始します...")
    test_personal_salary_data()
    test_empty_personal_data()
    print("\nテストが完了しました。") 