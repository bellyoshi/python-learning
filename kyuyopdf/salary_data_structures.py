from dataclasses import dataclass
from datetime import datetime
from typing import Optional, List, Dict, Any
from decimal import Decimal

@dataclass
class EmployeeInfo:
    """従業員基本情報の構造体"""
    社員番号: Optional[str] = None
    氏名: Optional[str] = None
    生年月日: Optional[datetime] = None
    
    def __post_init__(self):
        """データの検証と正規化"""
        if self.氏名:
            self.氏名 = str(self.氏名).strip()
        if self.社員番号:
            self.社員番号 = str(self.社員番号).strip()

@dataclass
class SalaryData:
    """給与データの構造体"""
    # 基本情報
    支給日: Optional[datetime] = None
    氏名: Optional[str] = None
    総支給額: Optional[Decimal] = None
    
    # 標準報酬月額
    標準報酬月額: Optional[Decimal] = None
    
    # 社会保険料
    健康保険: Optional[Decimal] = None
    厚生年金: Optional[Decimal] = None
    社会保険料控除後: Optional[Decimal] = None
    
    # 所得税
    源泉所得税: Optional[Decimal] = None
    
    # 最終金額
    差引支給額: Optional[Decimal] = None
    振込金額: Optional[Decimal] = None
    
    # 控除内訳
    健康保険料_従業員: Optional[Decimal] = None
    厚生年金_従業員: Optional[Decimal] = None
    社会保険料控除額: Optional[Decimal] = None
    扶養親族等の数: Optional[int] = None
    
    # その他の項目（動的に追加される可能性がある項目）
    その他: Dict[str, Any] = None
    
    def __post_init__(self):
        """データの検証と正規化"""
        if self.その他 is None:
            self.その他 = {}
        
        if self.氏名:
            self.氏名 = str(self.氏名).strip()
        
        # 数値項目の正規化
        numeric_fields = [
            '総支給額', '標準報酬月額', '健康保険', '厚生年金', 
            '社会保険料控除後', '源泉所得税', '差引支給額', '振込金額',
            '健康保険料_従業員', '厚生年金_従業員', '社会保険料控除額'
        ]
        
        for field in numeric_fields:
            value = getattr(self, field)
            if value is not None:
                if isinstance(value, str):
                    # 数式の場合はそのまま保持
                    if value.startswith('[数式:'):
                        continue
                    # 数値文字列をDecimalに変換
                    try:
                        # カンマを除去してから変換
                        clean_value = str(value).replace(',', '').replace('¥', '').strip()
                        if clean_value:
                            setattr(self, field, Decimal(clean_value))
                    except (ValueError, TypeError):
                        # 変換できない場合はそのまま保持
                        pass
                elif isinstance(value, (int, float)):
                    setattr(self, field, Decimal(str(value)))
        
        # 扶養親族等の数の正規化
        if self.扶養親族等の数 is not None:
            try:
                self.扶養親族等の数 = int(self.扶養親族等の数)
            except (ValueError, TypeError):
                self.扶養親族等の数 = None

@dataclass
class SalaryRecord:
    """給与記録の構造体（従業員情報と給与データを組み合わせたもの）"""
    従業員情報: EmployeeInfo
    給与データ: SalaryData
    月: int
    
    def __post_init__(self):
        """月の自動設定"""
        if self.給与データ.支給日 and hasattr(self.給与データ.支給日, 'month'):
            self.月 = self.給与データ.支給日.month

@dataclass
class SalarySummary:
    """給与サマリーの構造体（複数の給与記録をまとめたもの）"""
    従業員名: str
    給与記録: List[SalaryRecord]
    
    def get_monthly_data(self, month: int) -> Optional[SalaryRecord]:
        """指定された月の給与記録を取得"""
        for record in self.給与記録:
            if record.月 == month:
                return record
        return None
    
    def get_available_months(self) -> List[int]:
        """利用可能な月のリストを取得"""
        return sorted([record.月 for record in self.給与記録])
    
    def get_total_salary(self, month: int) -> Optional[Decimal]:
        """指定された月の総支給額を取得"""
        record = self.get_monthly_data(month)
        return record.給与データ.総支給額 if record else None

class SalaryDataFactory:
    """給与データ構造体のファクトリークラス"""
    
    @staticmethod
    def create_employee_info(社員番号: str = None, 氏名: str = None, 生年月日: datetime = None) -> EmployeeInfo:
        """従業員情報の構造体を作成"""
        return EmployeeInfo(
            社員番号=社員番号,
            氏名=氏名,
            生年月日=生年月日
        )
    
    @staticmethod
    def create_salary_data_from_dict(data_dict: Dict[str, Any]) -> SalaryData:
        """辞書から給与データの構造体を作成"""
        salary_data = SalaryData()
        
        # 基本フィールドのマッピング
        field_mapping = {
            '支給日': '支給日',
            '氏名': '氏名',
            '総支給額': '総支給額',
            '標準報酬月額': '標準報酬月額',
            '健康保険': '健康保険',
            '厚生年金': '厚生年金',
            '社会保険料控除後': '社会保険料控除後',
            '源泉所得税': '源泉所得税',
            '差引支給額': '差引支給額',
            '振込金額': '振込金額',
            '健康保険料（従業員）': '健康保険料_従業員',
            '厚生年金（従業員）': '厚生年金_従業員',
            '社会保険料控除額': '社会保険料控除額',
            '扶養親族等の数': '扶養親族等の数'
        }
        
        # マッピングに従ってデータを設定
        for excel_field, struct_field in field_mapping.items():
            if excel_field in data_dict:
                setattr(salary_data, struct_field, data_dict[excel_field])
        
        # その他の項目を設定
        for key, value in data_dict.items():
            if key not in field_mapping:
                salary_data.その他[key] = value
        
        return salary_data
    
    @staticmethod
    def create_salary_record(employee_info: EmployeeInfo, salary_data: SalaryData) -> SalaryRecord:
        """給与記録の構造体を作成"""
        return SalaryRecord(
            従業員情報=employee_info,
            給与データ=salary_data,
            月=salary_data.支給日.month if salary_data.支給日 else 0
        )
    
    @staticmethod
    def create_salary_summary(employee_name: str, salary_records: List[SalaryRecord]) -> SalarySummary:
        """給与サマリーの構造体を作成"""
        return SalarySummary(
            従業員名=employee_name,
            給与記録=salary_records
        ) 