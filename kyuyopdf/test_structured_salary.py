#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
給与データ構造体のテストファイル
"""

import os
from salary_pdf_generator import SalaryPayslipGenerator
from salary_data_structures import SalaryDataFactory, EmployeeInfo, SalaryData, SalaryRecord, SalarySummary

def test_structured_salary_data():
    """構造体を使用した給与データの取得テスト"""
    
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
        
        # 各従業員の給与データ月を確認
        for employee in employees:
            months = generator.get_employee_salary_months(employee)
            print(f"{employee}: {months}月")
        
        # 特定の従業員の給与データを構造体で取得
        test_employee = "渡邉俊行"
        test_month = 1
        
        print(f"\n=== {test_employee}の{test_month}月の給与データ（構造体形式）===")
        
        # 給与データを構造体で取得
        salary_data, headers = generator.get_salary_data(test_employee, test_month)
        
        if salary_data:
            print(f"給与データ構造体: {type(salary_data)}")
            print(f"支給日: {salary_data.支給日}")
            print(f"氏名: {salary_data.氏名}")
            print(f"総支給額: {salary_data.総支給額}")
            print(f"標準報酬月額: {salary_data.標準報酬月額}")
            print(f"健康保険: {salary_data.健康保険}")
            print(f"厚生年金: {salary_data.厚生年金}")
            print(f"社会保険料控除後: {salary_data.社会保険料控除後}")
            print(f"源泉所得税: {salary_data.源泉所得税}")
            print(f"差引支給額: {salary_data.差引支給額}")
            print(f"振込金額: {salary_data.振込金額}")
            print(f"健康保険料_従業員: {salary_data.健康保険料_従業員}")
            print(f"厚生年金_従業員: {salary_data.厚生年金_従業員}")
            print(f"社会保険料控除額: {salary_data.社会保険料控除額}")
            print(f"扶養親族等の数: {salary_data.扶養親族等の数}")
            
            # その他の項目を表示
            if salary_data.その他:
                print(f"\nその他の項目:")
                for key, value in salary_data.その他.items():
                    print(f"  {key}: {value}")
        else:
            print(f"給与データが見つかりませんでした。")
        
        # 従業員基本情報を構造体で取得
        print(f"\n=== {test_employee}の基本情報（構造体形式）===")
        employee_info = generator.get_employee_info(test_employee)
        
        if employee_info:
            print(f"従業員情報構造体: {type(employee_info)}")
            print(f"社員番号: {employee_info.社員番号}")
            print(f"氏名: {employee_info.氏名}")
            print(f"生年月日: {employee_info.生年月日}")
        else:
            print(f"従業員情報が見つかりませんでした。")
        
        # 給与記録構造体を作成
        if salary_data and employee_info:
            print(f"\n=== 給与記録構造体の作成 ===")
            salary_record = SalaryDataFactory.create_salary_record(employee_info, salary_data)
            print(f"給与記録構造体: {type(salary_record)}")
            print(f"月: {salary_record.月}")
            print(f"従業員情報: {salary_record.従業員情報.氏名}")
            print(f"給与データ: {salary_record.給与データ.総支給額}")
        
        # 複数の給与記録からサマリーを作成
        print(f"\n=== 給与サマリー構造体の作成 ===")
        salary_records = []
        
        # 複数月のデータを取得してサマリーを作成
        for month in range(1, 4):  # 1月から3月
            month_salary_data, _ = generator.get_salary_data(test_employee, month)
            if month_salary_data:
                month_record = SalaryDataFactory.create_salary_record(employee_info, month_salary_data)
                salary_records.append(month_record)
                print(f"{month}月の給与: {month_salary_data.総支給額}")
        
        if salary_records:
            salary_summary = SalaryDataFactory.create_salary_summary(test_employee, salary_records)
            print(f"給与サマリー構造体: {type(salary_summary)}")
            print(f"従業員名: {salary_summary.従業員名}")
            print(f"利用可能な月: {salary_summary.get_available_months()}")
            
            # 各月の総支給額を取得
            for month in salary_summary.get_available_months():
                total_salary = salary_summary.get_total_salary(month)
                print(f"{month}月の総支給額: {total_salary}")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        generator.close()

def test_data_validation():
    """データ検証のテスト"""
    print(f"\n=== データ検証テスト ===")
    
    # 従業員情報の作成と検証
    employee_info = SalaryDataFactory.create_employee_info(
        社員番号="001",
        氏名="  テスト太郎  ",
        生年月日=None
    )
    print(f"従業員情報: {employee_info}")
    print(f"氏名（正規化後）: '{employee_info.氏名}'")
    
    # 給与データの作成と検証
    salary_data = SalaryDataFactory.create_salary_data_from_dict({
        '支給日': '2024-01-15',
        '氏名': 'テスト太郎',
        '総支給額': '300,000',
        '標準報酬月額': '350,000',
        '健康保険': '15,000',
        '厚生年金': '25,000',
        '社会保険料控除後': '260,000',
        '源泉所得税': '10,000',
        '差引支給額': '250,000',
        '振込金額': '250,000',
        '健康保険料（従業員）': '15,000',
        '厚生年金（従業員）': '25,000',
        '社会保険料控除額': '40,000',
        '扶養親族等の数': '2',
        'その他の項目': 'テストデータ'
    })
    
    print(f"給与データ: {salary_data}")
    print(f"総支給額（Decimal）: {salary_data.総支給額}")
    print(f"扶養親族等の数（int）: {salary_data.扶養親族等の数}")
    print(f"その他の項目: {salary_data.その他}")

if __name__ == "__main__":
    print("給与データ構造体のテストを開始します...")
    test_structured_salary_data()
    test_data_validation()
    print("\nテストが完了しました。") 