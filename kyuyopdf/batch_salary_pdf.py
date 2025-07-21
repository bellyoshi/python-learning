import openpyxl
import os
from datetime import datetime
from salary_pdf_generator import SalaryPayslipGenerator
from personal_salary_data import PersonalSalaryData, PersonalSalaryDataManager

def get_available_employees_and_months(excel_file):
    """
    利用可能な従業員と月の組み合わせを取得
    """
    try:
        # PersonalSalaryDataManagerを使用してデータを読み込み
        manager = PersonalSalaryDataManager()
        manager.load_from_excel_file(excel_file)
        
        # 利用可能な従業員を取得
        employees = manager.get_available_employees()
        
        # 各従業員の給与データがある月を取得
        employee_months = {}
        for employee in employees:
            months = manager.get_available_months()
            if months:  # 給与データがある従業員のみ
                employee_months[employee] = months
        
        return employee_months
        
    except Exception as e:
        print(f"データ読み込みエラー: {e}")
        return {}

def create_all_payslips(excel_file, output_dir="給与明細PDF"):
    """
    全従業員の給与明細PDFを作成
    """
    # 出力ディレクトリを作成
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"出力ディレクトリ '{output_dir}' を作成しました。")
    
    # 利用可能な従業員と月を取得
    employee_months = get_available_employees_and_months(excel_file)
    
    if not employee_months:
        print("給与データが見つかりません。")
        return
    
    print("=== 利用可能な従業員と月 ===")
    for employee, months in employee_months.items():
        print(f"{employee}: {months}月")
    print()
    
    # PDF生成器を作成
    generator = SalaryPayslipGenerator(excel_file)
    
    success_count = 0
    total_count = 0
    
    try:
        for employee, months in employee_months.items():
            print(f"処理中: {employee}")
            
            for month in months:
                total_count += 1
                
                # ファイル名を生成
                safe_employee_name = employee.replace('/', '_').replace('\\', '_')
                output_filename = f"給与明細_{safe_employee_name}_{month}月.pdf"
                output_path = os.path.join(output_dir, output_filename)
                
                # 既存ファイルがある場合は削除（上書きのため）
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                        print(f"  既存ファイルを削除: {output_filename}")
                    except Exception as e:
                        print(f"  既存ファイル削除エラー: {e}")
                
                print(f"  {month}月の給与明細を作成中...")
                
                success = generator.create_payslip_pdf(employee, month, output_path)
                
                if success:
                    success_count += 1
                    print(f"  ✓ 作成完了: {output_filename}")
                else:
                    print(f"  ✗ 作成失敗: {month}月のデータが見つかりません")
                
                print()
        
        print(f"=== 処理完了 ===")
        print(f"成功: {success_count}/{total_count} 件")
        print(f"出力先: {output_dir}")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        generator.close()

def create_specific_payslips(excel_file, employee_names=None, months=None, output_dir="給与明細PDF"):
    """
    指定された従業員と月の給与明細PDFを作成
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"出力ディレクトリ '{output_dir}' を作成しました。")
    
    # 利用可能な従業員と月を取得
    employee_months = get_available_employees_and_months(excel_file)
    
    if not employee_months:
        print("給与データが見つかりません。")
        return
    
    # 指定された従業員をフィルタリング
    if employee_names:
        filtered_employees = {k: v for k, v in employee_months.items() if k in employee_names}
    else:
        filtered_employees = employee_months
    
    if not filtered_employees:
        print("指定された従業員のデータが見つかりません。")
        return
    
    # PDF生成器を作成
    generator = SalaryPayslipGenerator(excel_file)
    
    success_count = 0
    total_count = 0
    
    try:
        for employee, available_months in filtered_employees.items():
            print(f"処理中: {employee}")
            
            # 指定された月をフィルタリング
            target_months = available_months
            if months:
                target_months = [m for m in available_months if m in months]
            
            for month in target_months:
                total_count += 1
                
                # ファイル名を生成
                safe_employee_name = employee.replace('/', '_').replace('\\', '_')
                output_filename = f"給与明細_{safe_employee_name}_{month}月.pdf"
                output_path = os.path.join(output_dir, output_filename)
                
                # 既存ファイルがある場合は削除（上書きのため）
                if os.path.exists(output_path):
                    try:
                        os.remove(output_path)
                        print(f"  既存ファイルを削除: {output_filename}")
                    except Exception as e:
                        print(f"  既存ファイル削除エラー: {e}")
                
                print(f"  {month}月の給与明細を作成中...")
                
                success = generator.create_payslip_pdf(employee, month, output_path)
                
                if success:
                    success_count += 1
                    print(f"  ✓ 作成完了: {output_filename}")
                else:
                    print(f"  ✗ 作成失敗: {month}月のデータが見つかりません")
                
                print()
        
        print(f"=== 処理完了 ===")
        print(f"成功: {success_count}/{total_count} 件")
        print(f"出力先: {output_dir}")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        generator.close()

def test_excel_file_loading():
    """
    エクセルファイルの読み込みテスト
    """
    excel_file = "給与支給一覧令和7年.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"エラー: ファイル '{excel_file}' が見つかりません。")
        return False
    
    try:
        print("=== エクセルファイル読み込みテスト ===")
        
        # PersonalSalaryDataManagerを使用してデータを読み込み
        manager = PersonalSalaryDataManager()
        manager.load_from_excel_file(excel_file)
        
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
            employee_data = manager.get_personal_data_by_employee(test_employee)
            print(f"従業員 '{test_employee}' のデータ数: {len(employee_data)}")
            
            if employee_data:
                first_data = employee_data[0]
                print(f"最初のデータ:")
                print(f"  氏名: {first_data.氏名}")
                print(f"  支給日: {first_data.支給日}")
                print(f"  総支給額: {first_data.総支給額}")
                print(f"  振込金額: {first_data.振込金額}")
        
        print("=== テスト完了 ===")
        return True
        
    except Exception as e:
        print(f"テストエラー: {e}")
        return False

def main():
    """
    メイン関数
    """
    excel_file = "給与支給一覧令和7年.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"エラー: ファイル '{excel_file}' が見つかりません。")
        return
    
    # まずエクセルファイルの読み込みテストを実行
    if not test_excel_file_loading():
        print("エクセルファイルの読み込みに失敗しました。")
        return
    
    print("\n=== 給与明細PDF一括作成ツール ===")
    print("選択肢:")
    print("1. 全従業員の給与明細を作成 - 全ての従業員の全月分の給与明細を一括作成")
    print("2. 特定の従業員の給与明細を作成 - 指定した従業員の全月分の給与明細を作成")
    print("3. 特定の月の給与明細を作成 - 指定した月の全従業員の給与明細を作成")
    
    choice = input("\n選択してください (1-3): ").strip()
    
    if choice == "1":
        # 全従業員の給与明細を作成
        create_all_payslips(excel_file)
        
    elif choice == "2":
        # 特定の従業員の給与明細を作成
        employee_months = get_available_employees_and_months(excel_file)
        print("\n利用可能な従業員:")
        for i, employee in enumerate(employee_months.keys(), 1):
            print(f"{i}. {employee}")
        
        try:
            employee_choice = int(input("\n従業員番号を選択してください: ")) - 1
            employee_names = list(employee_months.keys())
            if 0 <= employee_choice < len(employee_names):
                selected_employee = employee_names[employee_choice]
                create_specific_payslips(excel_file, [selected_employee])
            else:
                print("無効な選択です。")
        except ValueError:
            print("無効な入力です。")
            
    elif choice == "3":
        # 特定の月の給与明細を作成
        month = input("月を入力してください (1-12): ").strip()
        try:
            month_num = int(month)
            if 1 <= month_num <= 12:
                create_specific_payslips(excel_file, months=[month_num])
            else:
                print("無効な月です。")
        except ValueError:
            print("無効な入力です。")
            
    else:
        print("無効な選択です。")

if __name__ == "__main__":
    main() 