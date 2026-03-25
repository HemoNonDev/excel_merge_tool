"""
複数のExcelファイルを1つに統合するツール
EC事業者の月別売上データ統合などに
"""

import pandas as pd
import glob
import os
from datetime import datetime

def merge_excel_files(input_folder, output_file):
    """
    指定フォルダ内の全Excelファイルを1つに統合
    
    Parameters:
    -----------
    input_folder : str
        読み込み元フォルダのパス
    output_file : str
        統合後のExcelファイルの保存先パス
    """
    
    # Excelファイルを検索してソート
    excel_files = glob.glob(f"{input_folder}/*.xlsx")
    excel_files.sort()

    if not excel_files:
        print(f"エラー: {input_folder} にExcelファイルが見つかりません")
        return
    
    print(f"見つかったファイル数: {len(excel_files)}")
    
    # ファイルの読み込み
    df_list = []
    for file in excel_files:
        try:
            df = pd.read_excel(file)
            # ファイル名を列として追加（データ元の識別用）
            df['元ファイル'] = os.path.basename(file)
            df_list.append(df)
            print(f"✓ 読み込み成功: {os.path.basename(file)} ({len(df.columns)}列)")
        except Exception as e:
            print(f"✗ 読み込み失敗: {os.path.basename(file)} - {e}")
    
    if not df_list:
        print("エラー: 読み込み可能なファイルが見つかりません")
        return
    
    # 全データを結合
    merged_df = pd.concat(df_list, ignore_index=True)
    
    # 結果を保存
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    merged_df.to_excel(output_file, index=False)
    print(f"\n統合完了！")
    print(f"出力ファイル: {output_file}")
    print(f"総行数: {len(merged_df)}")

if __name__ == "__main__":
    # 使用例
    input_folder = "sample_data"  # サンプルデータのフォルダ
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"output/統合結果_{timestamp}.xlsx"
    
    merge_excel_files(input_folder, output_file)