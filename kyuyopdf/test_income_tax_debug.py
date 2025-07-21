#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
源泉所得税のデータ読み込みデバッグテスト
"""

from personal_salary_data import PersonalSalaryData
from excel_columns import ExcelColumns

def test_income_tax_debug():
    """源泉所得税のデータ読み込みをデバッグ"""
    
    print("=== 源泉所得税データ読み込みデバッグ ===")
    
    # テスト用の行データ（実際のエクセルデータ）
    test_row_data = [
        "2025-01-25 00:00:00",  # 1: 支給日
        1,                       # 2: 社員番号
        "渡邉俊行",              # 3: 氏名
        92000,                  # 4: 総支給額
        88000,                  # 5: 標準報酬月額
        8975.999999999998,      # 6: 健康保険
        16104,                  # 7: 厚生年金
        4488,                   # 8: 健康保険料（従業員）
        8052,                   # 9: 厚生年金（従業員）
        12540,                  # 10: 社会保険料控除額
        79460,                  # 11: 社会保険料控除後
        6,                      # 12: 扶養親族等の数
        0,                      # 13: 源泉所得税 ← これが問題
        79460,                  # 14: 差引支給額
        False,                  # 15: 介護保険対象者
        0,                      # 16: 算出源泉所得税
        4487.999999999998,      # 17: 健康保険会社負担
        8052,                   # 18: 厚生年金会社負担
        0,                      # 19: 賃料控除
        0,                      # 20: 駐車場控除
        79460                   # 21: 振込金額
    ]
    
    print(f"テスト行データ:")
    for i, value in enumerate(test_row_data, 1):
        print(f"  {i}: {value} (型: {type(value)})")
    
    print(f"\n源泉所得税の列番号: {ExcelColumns.INCOME_TAX}")
    print(f"源泉所得税の値: {test_row_data[ExcelColumns.INCOME_TAX - 1]}")
    
    # PersonalSalaryDataインスタンスを作成
    personal_data = PersonalSalaryData.from_excel_row(test_row_data)
    
    print(f"\nPersonalSalaryDataの源泉所得税: {personal_data.源泉所得税}")
    print(f"源泉所得税の型: {type(personal_data.源泉所得税)}")
    
    # 他の給与項目も確認
    print(f"\n他の給与項目:")
    print(f"  総支給額: {personal_data.総支給額}")
    print(f"  健康保険料_従業員: {personal_data.健康保険料_従業員}")
    print(f"  厚生年金_従業員: {personal_data.厚生年金_従業員}")
    print(f"  社会保険料控除額: {personal_data.社会保険料控除額}")
    print(f"  振込金額: {personal_data.振込金額}")
    
    # 辞書形式での出力
    dict_data = personal_data.to_dict()
    print(f"\n辞書形式での源泉所得税: {dict_data['源泉所得税']}")
    
    print("\n=== デバッグ完了 ===")

if __name__ == "__main__":
    test_income_tax_debug() 