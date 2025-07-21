#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excelファイルの列番号をシンボル化するクラス
"""

class ExcelColumns:
    """Excelファイルの列番号をシンボル化"""
    # 給与データシートの列番号
    PAYMENT_DATE = 1      # 支給日
    EMPLOYEE_ID = 2       # 社員番号
    EMPLOYEE_NAME = 3     # 氏名
    TOTAL_SALARY = 4      # 総支給額
    STANDARD_REWARD = 5   # 標準報酬月額
    HEALTH_INSURANCE = 6  # 健康保険
    KOUSEI_INSURANCE = 7  # 厚生年金
    SOCIAL_INSURANCE_AFTER = 8  # 社会保険料控除後
    INCOME_TAX = 13        # 源泉所得税
    TRANSFER_AMOUNT = 21  # 振込金額
    HEALTH_INSURANCE_EMPLOYEE = 8  # 健康保険料（従業員）
    PENSION_EMPLOYEE = 9 # 厚生年金（従業員）
    SOCIAL_INSURANCE_DEDUCTION = 10 # 社会保険料控除額
    HEALTH_INSURANCE_COMPANY = 17  # 健康保険料（会社負担）
    PENSION_COMPANY = 18  # 厚生年金（会社負担）
    RENT_DEDUCTION = 19  # 賃料控除
    PARKING_DEDUCTION = 20  # 駐車場控除
    DEPENDENTS_COUNT = 12 # 扶養親族等の数 