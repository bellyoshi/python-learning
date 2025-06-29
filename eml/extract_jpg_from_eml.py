import os
import email
import email.header
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import mimetypes
import re

def extract_jpg_from_eml(eml_file_path, output_dir="extracted_images"):
    """
    EMLファイルからJPG画像を抽出する関数
    
    Args:
        eml_file_path (str): EMLファイルのパス
        output_dir (str): 抽出した画像を保存するディレクトリ
    """
    
    # 出力ディレクトリを作成
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"出力ディレクトリ '{output_dir}' を作成しました")
    
    # EMLファイルを読み込み
    try:
        with open(eml_file_path, 'rb') as f:
            msg = email.message_from_bytes(f.read())
    except Exception as e:
        print(f"EMLファイルの読み込みエラー: {e}")
        return
    
    print(f"EMLファイル '{eml_file_path}' を処理中...")
    
    extracted_count = 0
    
    def process_message(message, level=0):
        nonlocal extracted_count
        
        # メッセージの各部分を処理
        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            
            # ファイル名を取得
            filename = part.get_filename()
            if filename:
                # ファイル名のデコード
                try:
                    filename = email.header.decode_header(filename)[0][0]
                    if isinstance(filename, bytes):
                        filename = filename.decode('utf-8', errors='ignore')
                except:
                    pass
            
            # Content-Typeを確認
            content_type = part.get_content_type()
            
            # JPG画像かどうかをチェック
            is_jpg = False
            if content_type in ['image/jpeg', 'image/jpg']:
                is_jpg = True
            elif filename and filename.lower().endswith(('.jpg', '.jpeg')):
                is_jpg = True
            
            if is_jpg:
                # ファイル名がない場合はデフォルト名を生成
                if not filename:
                    filename = f"extracted_image_{extracted_count + 1}.jpg"
                
                # ファイル名の正規化（特殊文字を除去）
                safe_filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
                
                # 重複を避けるため、必要に応じて番号を追加
                base_name, ext = os.path.splitext(safe_filename)
                counter = 1
                final_filename = safe_filename
                while os.path.exists(os.path.join(output_dir, final_filename)):
                    final_filename = f"{base_name}_{counter}{ext}"
                    counter += 1
                
                # 画像データを取得
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        # ファイルに保存
                        output_path = os.path.join(output_dir, final_filename)
                        with open(output_path, 'wb') as img_file:
                            img_file.write(payload)
                        
                        extracted_count += 1
                        print(f"  抽出: {final_filename} ({len(payload)} bytes)")
                    else:
                        print(f"  警告: {filename} のデータが空です")
                        
                except Exception as e:
                    print(f"  エラー: {filename} の抽出に失敗: {e}")
    
    # メッセージを処理
    process_message(msg)
    
    print(f"完了: {extracted_count} 個のJPG画像を抽出しました")
    return extracted_count

def main():
    """メイン関数"""
    print("EMLファイルからJPG画像を抽出します...")
    print("=" * 50)
    
    # 現在のディレクトリのEMLファイルを検索
    eml_files = [f for f in os.listdir('.') if f.lower().endswith('.eml')]
    
    if not eml_files:
        print("EMLファイルが見つかりませんでした")
        return
    
    total_extracted = 0
    
    for eml_file in eml_files:
        print(f"\n処理中: {eml_file}")
        print("-" * 30)
        
        # 各EMLファイル用のサブディレクトリを作成
        base_name = os.path.splitext(eml_file)[0]
        output_dir = f"extracted_images_{base_name}"
        
        extracted = extract_jpg_from_eml(eml_file, output_dir)
        if extracted is not None:
            total_extracted += extracted
    
    print("\n" + "=" * 50)
    print(f"総抽出数: {total_extracted} 個のJPG画像")
    print("抽出完了!")

if __name__ == "__main__":
    main() 