#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
給与項目の詳細デバッグテスト
"""

from personal_salary_data import PersonalSalaryData, PersonalSalaryDataManager
from excel_columns import ExcelColumns

def debug_salary_items():
    """給与項目の詳細デバッグ"""
    
    print("=== 給与項目詳細デバッグ ===")
    
    # PersonalSalaryDataManagerを使用してデータを読み込み
    manager = PersonalSalaryDataManager()
    manager.load_from_excel_file('給与支給一覧令和7年.xlsx')
    
    # 最初のデータを取得
    all_data = manager.get_all_personal_data()
    if all_data:
        first_data = all_data[0]
        
        print(f"最初のデータ詳細:")
        print(f"  氏名: {first_data.氏名}")
        print(f"  支給日: {first_data.支給日}")
        print(f"  総支給額: {first_data.総支給額}")
        print(f"  標準報酬月額: {first_data.標準報酬月額}")
        print(f"  健康保険: {first_data.健康保険}")
        print(f"  厚生年金: {first_data.厚生年金}")
        print(f"  社会保険料控除後: {first_data.社会保険料控除後}")
        print(f"  源泉所得税: {first_data.源泉所得税}")
        print(f"  健康保険料_従業員: {first_data.健康保険料_従業員}")
        print(f"  厚生年金_従業員: {first_data.厚生年金_従業員}")
        print(f"  社会保険料控除額: {first_data.社会保険料控除額}")
        print(f"  健康保険料_会社負担: {first_data.健康保険料_会社負担}")
        print(f"  厚生年金_会社負担: {first_data.厚生年金_会社負担}")
        print(f"  賃料控除: {first_data.賃料控除}")
        print(f"  駐車場控除: {first_data.駐車場控除}")
        print(f"  振込金額: {first_data.振込金額}")
        
        # 辞書形式での出力
        dict_data = first_data.to_dict()
        print(f"\n辞書形式でのデータ:")
        print(f"  標準報酬月額: {dict_data['標準報酬月額']}")
        print(f"  健康保険: {dict_data['健康保険']}")
        print(f"  厚生年金: {dict_data['厚生年金']}")
        print(f"  社会保険料控除後: {dict_data['社会保険料控除後']}")
        print(f"  源泉所得税: {dict_data['源泉所得税']}")
        print(f"  健康保険料_会社負担: {dict_data.get('健康保険料_会社負担', 'N/A')}")
        print(f"  厚生年金_会社負担: {dict_data.get('厚生年金_会社負担', 'N/A')}")
        print(f"  賃料控除: {dict_data.get('賃料控除', 'N/A')}")
        print(f"  駐車場控除: {dict_data.get('駐車場控除', 'N/A')}")
        
        # データ型の確認
        print(f"\nデータ型の確認:")
        print(f"  標準報酬月額の型: {type(first_data.標準報酬月額)}")
        print(f"  健康保険の型: {type(first_data.健康保険)}")
        print(f"  厚生年金の型: {type(first_data.厚生年金)}")
        print(f"  社会保険料控除後の型: {type(first_data.社会保険料控除後)}")
        print(f"  源泉所得税の型: {type(first_data.源泉所得税)}")
        
    # 統計情報も確認
    stats = manager.get_statistics()
    print(f"\n統計情報:")
    print(f"  総レコード数: {stats['total_records']}")
    print(f"  平均総支給額: {stats['average_total_salary']}")
    print(f"  最大総支給額: {stats['max_total_salary']}")
    print(f"  最小総支給額: {stats['min_total_salary']}")
    
    print("\n=== デバッグ完了 ===")

if __name__ == "__main__":
    debug_salary_items() 