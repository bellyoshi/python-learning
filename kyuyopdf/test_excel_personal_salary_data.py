#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
エクセルファイルからPersonalSalaryDataを取得するメソッドのテスト
"""

import os
from decimal import Decimal
from datetime import datetime
from personal_salary_data import PersonalSalaryData, PersonalSalaryDataManager

def test_excel_personal_salary_data():
    """エクセルファイルからPersonalSalaryDataを取得するテスト"""
    
    print("=== エクセルファイルからPersonalSalaryDataを取得するテスト ===")
    
    # テスト用のエクセルファイルパス
    excel_file_path = "給与データ.xlsx"
    
    if not os.path.exists(excel_file_path):
        print(f"エクセルファイル '{excel_file_path}' が見つかりません。")
        print("テスト用のエクセルファイルを用意してください。")
        return
    
    # 1. PersonalSalaryData.from_excel_file() のテスト
    print("\n1. 全データの読み込みテスト")
    try:
        personal_data_list = PersonalSalaryData.from_excel_file(excel_file_path)
        print(f"読み込まれたデータ数: {len(personal_data_list)}")
        
        if personal_data_list:
            # 最初のデータを表示
            first_data = personal_data_list[0]
            print(f"最初のデータ:")
            print(f"  氏名: {first_data.氏名}")
            print(f"  支給日: {first_data.支給日}")
            print(f"  総支給額: {first_data.総支給額}")
            print(f"  標準報酬月額: {first_data.標準報酬月額}")
            print(f"  健康保険: {first_data.健康保険}")
            print(f"  厚生年金: {first_data.厚生年金}")
            print(f"  社会保険料控除後: {first_data.社会保険料控除後}")
            print(f"  源泉所得税: {first_data.源泉所得税}")
            print(f"  振込金額: {first_data.振込金額}")
            print(f"  健康保険料_従業員: {first_data.健康保険料_従業員}")
            print(f"  厚生年金_従業員: {first_data.厚生年金_従業員}")
            print(f"  社会保険料控除額: {first_data.社会保険料控除額}")
            print(f"  扶養親族等の数: {first_data.扶養親族等の数}")
            
            # 辞書形式での出力テスト
            print(f"\n辞書形式での出力:")
            dict_data = first_data.to_dict()
            for key, value in dict_data.items():
                print(f"  {key}: {value}")
        
    except Exception as e:
        print(f"エラー: {e}")
    
    # 2. 特定の従業員のデータ読み込みテスト
    print("\n2. 特定の従業員のデータ読み込みテスト")
    try:
        # 利用可能な従業員を確認
        all_data = PersonalSalaryData.from_excel_file(excel_file_path)
        available_employees = list(set(data.氏名 for data in all_data if data.氏名))
        print(f"利用可能な従業員: {available_employees}")
        
        if available_employees:
            test_employee = available_employees[0]
            employee_data = PersonalSalaryData.from_excel_file(
                excel_file_path, employee_name=test_employee
            )
            print(f"従業員 '{test_employee}' のデータ数: {len(employee_data)}")
            
            if employee_data:
                print(f"従業員 '{test_employee}' の最初のデータ:")
                first_employee_data = employee_data[0]
                print(f"  支給日: {first_employee_data.支給日}")
                print(f"  総支給額: {first_employee_data.総支給額}")
                print(f"  振込金額: {first_employee_data.振込金額}")
        
    except Exception as e:
        print(f"エラー: {e}")
    
    # 3. 特定の行のデータ読み込みテスト
    print("\n3. 特定の行のデータ読み込みテスト")
    try:
        row_data = PersonalSalaryData.from_excel_file(excel_file_path, row_number=2)
        print(f"行2のデータ数: {len(row_data)}")
        
        if row_data:
            print(f"行2のデータ:")
            row2_data = row_data[0]
            print(f"  氏名: {row2_data.氏名}")
            print(f"  支給日: {row2_data.支給日}")
            print(f"  総支給額: {row2_data.総支給額}")
        
    except Exception as e:
        print(f"エラー: {e}")
    
    # 4. PersonalSalaryDataManagerのテスト
    print("\n4. PersonalSalaryDataManagerのテスト")
    try:
        manager = PersonalSalaryDataManager()
        
        # エクセルファイルからデータを読み込み
        manager.load_from_excel_file(excel_file_path)
        
        # 統計情報を取得
        stats = manager.get_statistics()
        print(f"統計情報:")
        for key, value in stats.items():
            print(f"  {key}: {value}")
        
        # 利用可能な従業員を取得
        available_employees = manager.get_available_employees()
        print(f"利用可能な従業員: {available_employees}")
        
        # 利用可能な月を取得
        available_months = manager.get_available_months()
        print(f"利用可能な月: {available_months}")
        
        if available_employees:
            test_employee = available_employees[0]
            
            # 特定の従業員のデータを取得
            employee_data = manager.get_personal_data_by_employee(test_employee)
            print(f"従業員 '{test_employee}' のデータ数: {len(employee_data)}")
            
            # 従業員の総支給額合計を取得
            total_salary = manager.get_total_salary_by_employee(test_employee)
            print(f"従業員 '{test_employee}' の総支給額合計: {total_salary}")
            
            if available_months:
                test_month = available_months[0]
                
                # 特定の従業員と月のデータを取得
                month_data = manager.get_personal_data_by_employee_and_month(test_employee, test_month)
                print(f"従業員 '{test_employee}' の {test_month}月のデータ数: {len(month_data)}")
        
        # 辞書形式でのエクスポートテスト
        dict_list = manager.export_to_dict_list()
        print(f"辞書形式でのエクスポート件数: {len(dict_list)}")
        
    except Exception as e:
        print(f"エラー: {e}")
    
    # 5. 特定の従業員のデータ読み込みテスト（Manager）
    print("\n5. 特定の従業員のデータ読み込みテスト（Manager）")
    try:
        manager = PersonalSalaryDataManager()
        
        if available_employees:
            test_employee = available_employees[0]
            manager.load_from_excel_file_by_employee(excel_file_path, test_employee)
            
            employee_data = manager.get_all_personal_data()
            print(f"Manager経由で読み込まれた従業員 '{test_employee}' のデータ数: {len(employee_data)}")
        
    except Exception as e:
        print(f"エラー: {e}")
    
    # 6. 特定の行のデータ読み込みテスト（Manager）
    print("\n6. 特定の行のデータ読み込みテスト（Manager）")
    try:
        manager = PersonalSalaryDataManager()
        manager.load_from_excel_file_by_row(excel_file_path, 2)
        
        row_data = manager.get_all_personal_data()
        print(f"Manager経由で読み込まれた行2のデータ数: {len(row_data)}")
        
        if row_data:
            print(f"行2のデータ:")
            row2_data = row_data[0]
            print(f"  氏名: {row2_data.氏名}")
            print(f"  支給日: {row2_data.支給日}")
            print(f"  総支給額: {row2_data.総支給額}")
        
    except Exception as e:
        print(f"エラー: {e}")
    
    print("\n=== テスト完了 ===")

if __name__ == "__main__":
    test_excel_personal_salary_data() 