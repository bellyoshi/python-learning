# 給与明細PDF作成プログラム

このプログラムは、Excelファイルから給与データを読み取り、給与明細をPDF形式で作成するPythonアプリケーションです。

## 機能

- Excelファイルから給与データを読み込み
- 従業員別・月別の給与明細PDFを生成
- 日本語フォント対応
- 一括処理機能
- 美しいレイアウトのPDF出力
- **構造体を使用したデータ管理**
- **個人給与データクラスによるデータ管理**

## 必要なライブラリ

```bash
pip install -r requirements.txt
```

### 依存ライブラリ
- `openpyxl`: Excelファイル読み込み
- `reportlab`: PDF生成
- `Pillow`: 画像処理

## ファイル構成

```
kyuyopdf/
├── excel_reader.py          # Excelファイル内容確認ツール
├── salary_pdf_generator.py  # 給与明細PDF生成メインプログラム
├── batch_salary_pdf.py      # 一括処理ツール
├── salary_data_structures.py # 給与データ構造体定義
├── personal_salary_data.py  # 個人給与データクラス
├── test_structured_salary.py # 構造体テストファイル
├── test_personal_salary_data.py # 個人給与データクラステストファイル
├── requirements.txt         # 依存ライブラリリスト
├── README.md               # このファイル
└── 給与支給一覧令和7年.xlsx  # サンプルExcelファイル
```

## データ構造体

### EmployeeInfo（従業員情報構造体）
```python
@dataclass
class EmployeeInfo:
    社員番号: Optional[str] = None
    氏名: Optional[str] = None
    生年月日: Optional[datetime] = None
```

### SalaryData（給与データ構造体）
```python
@dataclass
class SalaryData:
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
    
    # その他の項目
    その他: Dict[str, Any] = None
```

### SalaryRecord（給与記録構造体）
```python
@dataclass
class SalaryRecord:
    従業員情報: EmployeeInfo
    給与データ: SalaryData
    月: int
```

### SalarySummary（給与サマリー構造体）
```python
@dataclass
class SalarySummary:
    従業員名: str
    給与記録: List[SalaryRecord]
```

## 個人給与データクラス

### PersonalSalaryData（個人給与データクラス）
```python
@dataclass
class PersonalSalaryData:
    # 給与データシートのプロパティ
    支給日: Optional[datetime] = None
    氏名: Optional[str] = None
    総支給額: Optional[Decimal] = None
    標準報酬月額: Optional[Decimal] = None
    健康保険: Optional[Decimal] = None
    厚生年金: Optional[Decimal] = None
    社会保険料控除後: Optional[Decimal] = None
    源泉所得税: Optional[Decimal] = None
    差引支給額: Optional[Decimal] = None
    振込金額: Optional[Decimal] = None
    健康保険料_従業員: Optional[Decimal] = None
    厚生年金_従業員: Optional[Decimal] = None
    社会保険料控除額: Optional[Decimal] = None
    扶養親族等の数: Optional[int] = None
    
    # 個人データシートのプロパティ
    社員番号: Optional[str] = None
    生年月日: Optional[datetime] = None
```

### PersonalSalaryDataManager（個人給与データ管理クラス）
```python
class PersonalSalaryDataManager:
    def add_personal_data(self, personal_data: PersonalSalaryData)
    def get_by_employee_name(self, employee_name: str) -> List[PersonalSalaryData]
    def get_by_month(self, month: int) -> List[PersonalSalaryData]
    def get_by_employee_and_month(self, employee_name: str, month: int) -> Optional[PersonalSalaryData]
    def get_total_salary_by_employee(self, employee_name: str) -> Optional[Decimal]
    def get_available_employees(self) -> List[str]
    def get_available_months(self) -> List[int]
    def export_to_dict(self) -> List[Dict[str, Any]]
```

## 使用方法

### 1. Excelファイル内容確認

```bash
python excel_reader.py
```

Excelファイルの構造と内容を確認できます。

### 2. 単一の給与明細PDF作成

```bash
python salary_pdf_generator.py
```

渡邉俊行さんの1月の給与明細PDFを作成します。

### 3. 一括処理ツール

```bash
python batch_salary_pdf.py
```

対話形式で以下の選択肢から選べます：

1. **全従業員の給与明細を作成**: 全ての従業員の全月分の給与明細を作成
2. **特定の従業員の給与明細を作成**: 指定した従業員の給与明細を作成
3. **特定の月の給与明細を作成**: 指定した月の全従業員の給与明細を作成

### 4. 構造体テスト

```bash
python test_structured_salary.py
```

構造体を使用した給与データの取得と処理をテストします。

### 5. 個人給与データクラステスト

```bash
python test_personal_salary_data.py
```

個人給与データクラスの機能をテストします。

## 出力ファイル

### 単一PDF作成
- `給与明細_渡邉俊行_1月.pdf`

### 一括処理
- `給与明細PDF/` ディレクトリ内に従業員別・月別のPDFファイルが作成されます
- 例: `給与明細_渡邉俊行_7月.pdf`

## PDFの内容

作成されるPDFには以下の情報が含まれます：

### 基本情報
- 氏名
- 社員番号
- 生年月日
- 支給年月
- 支給日

### 給与詳細
- 総支給額
- 標準報酬月額
- 健康保険料
- 厚生年金
- 社会保険料控除後
- 源泉所得税
- 差引支給額
- 振込金額

### 控除内訳
- 健康保険料
- 厚生年金
- 社会保険料控除額
- 扶養親族等の数

## 対応データ形式

### Excelファイル要件
- シート名: `給与データ` (給与情報)
- シート名: `個人データ` (従業員基本情報)
- 必須列:
  - 支給日 (A列)
  - 氏名 (B列)
  - 総支給額 (C列)
  - その他の給与項目

## 構造体の利点

### 1. 型安全性
- 各フィールドの型が明確に定義されている
- コンパイル時の型チェックが可能

### 2. データ検証
- `__post_init__`メソッドでデータの正規化と検証
- 数値項目の自動変換（文字列→Decimal）
- 文字列の自動トリム

### 3. 拡張性
- 新しいフィールドの追加が容易
- その他の項目を`その他`辞書で管理

### 4. 使いやすさ
- 属性アクセス（`salary_data.総支給額`）
- IDEの自動補完対応
- 明確なドキュメント

## 個人給与データクラスの利点

### 1. ExcelColumnsとの整合性
- ExcelColumnsで定義されているプロパティのみを含む
- 列番号のシンボル化による保守性の向上

### 2. 統合されたデータ管理
- 給与データと個人データを統合
- 一つのクラスで全ての情報を管理

### 3. 柔軟なデータ取得
- 従業員名、月、社員番号などでの検索
- 複数の検索条件に対応

### 4. データ管理機能
- PersonalSalaryDataManagerによる一括管理
- エクスポート機能

### 5. エクセルファイルからの直接読み込み
- `from_excel_file()`: エクセルファイルから全データを読み込み
- `from_excel_file_by_employee()`: 特定の従業員のデータを読み込み
- `from_excel_file_by_row()`: 特定の行のデータを読み込み
- データ型の自動変換（文字列→Decimal、日付変換など）
- エラーハンドリング機能

## エクセルファイルからのデータ読み込み

### PersonalSalaryDataクラスのメソッド

#### 1. from_excel_file()
```python
# 全データを読み込み
personal_data_list = PersonalSalaryData.from_excel_file("給与データ.xlsx")

# 特定の従業員のデータを読み込み
employee_data = PersonalSalaryData.from_excel_file(
    "給与データ.xlsx", employee_name="田中太郎"
)

# 特定の行のデータを読み込み
row_data = PersonalSalaryData.from_excel_file(
    "給与データ.xlsx", row_number=2
)
```

#### 2. from_excel_row()
```python
# エクセルの行データからインスタンスを作成
row_data = ["2024-01-15", "001", "田中太郎", "300000", ...]
personal_data = PersonalSalaryData.from_excel_row(row_data)
```

### PersonalSalaryDataManagerクラスのメソッド

#### 1. load_from_excel_file()
```python
manager = PersonalSalaryDataManager()
manager.load_from_excel_file("給与データ.xlsx")
```

#### 2. load_from_excel_file_by_employee()
```python
manager = PersonalSalaryDataManager()
manager.load_from_excel_file_by_employee("給与データ.xlsx", "田中太郎")
```

#### 3. load_from_excel_file_by_row()
```python
manager = PersonalSalaryDataManager()
manager.load_from_excel_file_by_row("給与データ.xlsx", 2)
```

### データ型の自動変換

- **日付**: Excelの日付オブジェクト → `datetime`
- **数値**: 文字列/数値 → `Decimal`
- **整数**: 文字列/数値 → `int`
- **文字列**: そのまま保持

### エラーハンドリング

- ファイルが存在しない場合のエラー処理
- データ型変換エラーの処理
- 空行や無効なデータのスキップ

## カスタマイズ

### フォント設定
Windows環境では自動的に日本語フォント（MS Gothic）が使用されます。
他の環境では `salary_pdf_generator.py` の `FONT_NAME` を変更してください。

### レイアウト変更
`salary_pdf_generator.py` の `create_payslip_pdf` メソッド内でレイアウトを調整できます。

### 構造体の拡張
`salary_data_structures.py` で新しいフィールドや構造体を追加できます。

### 個人給与データクラスの拡張
`personal_salary_data.py` で新しいプロパティやメソッドを追加できます。

## トラブルシューティング

### よくある問題

1. **フォントエラー**
   - Windows標準フォントが利用できない場合は、フォントパスを確認してください

2. **Excelファイル読み込みエラー**
   - ファイルパスが正しいか確認
   - ファイルが破損していないか確認

3. **PDF作成エラー**
   - 出力ディレクトリの書き込み権限を確認
   - 十分なディスク容量があるか確認

4. **構造体エラー**
   - データ型の不一致を確認
   - 必須フィールドの値が設定されているか確認

5. **個人給与データクラスエラー**
   - ExcelColumnsとの整合性を確認
   - データの型変換エラーを確認

## ライセンス

このプログラムは教育目的で作成されています。

## 更新履歴

- v1.2: 個人給与データクラス対応版
  - 個人給与データクラスの追加
  - ExcelColumnsとの整合性確保
  - データ管理機能の強化
  - テストファイルの追加

- v1.1: 構造体対応版
  - 給与データ構造体の追加
  - 型安全性の向上
  - データ検証機能の追加
  - テストファイルの追加

- v1.0: 初期バージョン
  - Excelファイル読み込み機能
  - 給与明細PDF生成機能
  - 一括処理機能 