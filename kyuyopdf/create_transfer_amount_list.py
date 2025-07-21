#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
年月を指定して振込金額一覧をPDFで作成するバッチスクリプト
"""

import os
import sys
from datetime import datetime
from decimal import Decimal
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from personal_salary_data import PersonalSalaryDataManager

class TransferAmountListGenerator:
    """振込金額一覧PDF生成クラス"""
    
    def __init__(self, excel_file_path: str):
        """初期化"""
        self.excel_file_path = excel_file_path
        self.manager = PersonalSalaryDataManager()
        
    def load_data(self):
        """データを読み込み"""
        try:
            self.manager.load_from_excel_file(self.excel_file_path)
            print(f"データ読み込み完了: {len(self.manager.get_all_personal_data())} 件")
            return True
        except Exception as e:
            print(f"データ読み込みエラー: {e}")
            return False
    
    def get_transfer_amounts_by_month(self, year: int, month: int):
        """指定年月の振込金額データを取得"""
        all_data = self.manager.get_all_personal_data()
        month_data = []
        
        for data in all_data:
            if data.支給日 and data.支給日.year == year and data.支給日.month == month:
                month_data.append(data)
        
        return month_data
    
    def create_transfer_amount_list_pdf(self, year: int, month: int, output_path: str):
        """振込金額一覧PDFを作成"""
        try:
            # 指定年月のデータを取得
            month_data = self.get_transfer_amounts_by_month(year, month)
            
            if not month_data:
                print(f"{year}年{month}月のデータが見つかりません")
                return False
            
            # 出力ディレクトリを作成
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # PDFを作成
            self._create_pdf(month_data, year, month, output_path)
            return True
            
        except Exception as e:
            print(f"PDF作成エラー: {e}")
            return False
    
    def _create_pdf(self, month_data, year: int, month: int, output_path: str):
        """PDFを生成"""
        c = canvas.Canvas(output_path, pagesize=landscape(A4))
        
        # 日本語フォント設定
        try:
            # Windows標準の日本語フォントを試行
            font_paths = [
                'C:/Windows/Fonts/msgothic.ttc',
                'C:/Windows/Fonts/msmincho.ttc',
                'C:/Windows/Fonts/yu Gothic.ttc',
                'C:/Windows/Fonts/meiryo.ttc'
            ]
            
            font_name = 'JapaneseFont'
            font_loaded = False
            
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        font_loaded = True
                        break
                    except:
                        continue
            
            if not font_loaded:
                # フォールバック: 標準フォント
                font_name = 'Helvetica'
                
        except Exception as e:
            print(f"フォント設定エラー: {e}")
            font_name = 'Helvetica'
        
        c.setFont(font_name, 12)
        
        # タイトル
        c.setFont(font_name, 16)
        c.drawString(50*mm, 190*mm, f"振込金額一覧")
        c.setFont(font_name, 12)
        c.drawString(50*mm, 175*mm, f"{year}年{month}月")
        
        # テーブルデータを作成
        table_data = []
        
        # ヘッダー
        table_data.append([
            "No.",
            "氏名",
            "振込金額",
            "健康保険料_会社負担",
            "厚生年金_会社負担",
            "源泉所得税"
        ])
        
        # データ行
        total_transfer = Decimal('0')
        total_health_company = Decimal('0')
        total_pension_company = Decimal('0')
        total_tax = Decimal('0')
        
        for i, data in enumerate(month_data, 1):
            transfer_amount = data.振込金額 or Decimal('0')
            health_company_amount = data.健康保険料_会社負担 or Decimal('0')
            pension_company_amount = data.厚生年金_会社負担 or Decimal('0')
            tax_amount = data.源泉所得税 or Decimal('0')
            
            table_data.append([
                str(i),
                data.氏名 or "N/A",
                f"{transfer_amount:,.0f}円",
                f"{health_company_amount:,.2f}円",
                f"{pension_company_amount:,.2f}円",
                f"{tax_amount:,.0f}円"
            ])
            
            total_transfer += transfer_amount
            total_health_company += health_company_amount
            total_pension_company += pension_company_amount
            total_tax += tax_amount
        
        # 合計行
        table_data.append([
            "合計",
            f"{len(month_data)}名",
            f"{total_transfer:,.0f}円",
            f"{total_health_company:,.2f}円",
            f"{total_pension_company:,.2f}円",
            f"{total_tax:,.0f}円"
        ])
        
        # テーブルを作成（横書き用に列幅を調整）
        table = Table(table_data, colWidths=[25*mm, 50*mm, 40*mm, 40*mm, 40*mm, 40*mm])
        
        # テーブルスタイル
        style = TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), font_name),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (1, 1), (1, -2), 'LEFT'),  # 氏名は左寄せ
            ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),  # 金額は右寄せ
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -2), font_name),
            ('FONTSIZE', (0, 1), (-1, -2), 9),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), font_name),
            ('FONTSIZE', (0, -1), (-1, -1), 10),
            ('FONTNAME', (1, -1), (1, -1), font_name),
        ])
        
        table.setStyle(style)
        
        # テーブルを配置（横書き用に調整、位置をさらに下げる）
        table.wrapOn(c, 270*mm, 150*mm)
        table.drawOn(c, 10*mm, 110*mm)
        
        # 作成日時
        c.setFont(font_name, 8)
        c.drawString(10*mm, 20*mm, f"作成日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        
        c.save()
        print(f"振込金額一覧PDF作成完了: {output_path}")

def main():
    """メイン関数"""
    excel_file = "給与支給一覧令和7年.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"エラー: {excel_file} が見つかりません")
        return
    
    # 年月の入力
    try:
        year = int(input("年を入力してください (例: 2025): "))
        month = int(input("月を入力してください (1-12): "))
        
        if month < 1 or month > 12:
            print("エラー: 月は1-12の範囲で入力してください")
            return
            
    except ValueError:
        print("エラー: 正しい数値を入力してください")
        return
    
    # 生成器を作成
    generator = TransferAmountListGenerator(excel_file)
    
    # データを読み込み
    if not generator.load_data():
        return
    
    # 出力パス
    output_dir = "振込金額一覧PDF"
    output_path = os.path.join(output_dir, f"振込金額一覧_{year}年{month}月.pdf")
    
    # PDFを作成
    if generator.create_transfer_amount_list_pdf(year, month, output_path):
        print(f"\n✓ 振込金額一覧PDFが正常に作成されました: {output_path}")
    else:
        print("\n✗ PDF作成に失敗しました")

if __name__ == "__main__":
    main() 