import os
import pandas as pd
import csv

# ---------------------------------------
# Excel・csvファイル行数確認
# created: 2023/04/18
# ---------------------------------------

# 行数カウント Excel
def get_excel_row_count(file_path):
    df = pd.read_excel(file_path)
    return len(df)

# 行数カウント csv
def get_csv_row_count(file_path, encoding='utf-8'):
    with open(file_path, 'r', encoding=encoding) as file:
        csv_reader = csv.reader(file)
        row_count = sum(1 for row in csv_reader)
    return row_count

# 指定した列がユニークな行数カウント Excel
def get_excel_unique_item_count(file_path):
    df = pd.read_excel(file_path)
    if 'ITEM' not in df.columns:
        return None
    unique_item_count = df['ITEM'].nunique()
    return unique_item_count

# 指定した列がユニークな行数カウント csv
def get_csv_unique_item_count(file_path, encoding='utf-8'):
    with open(file_path, 'r', encoding=encoding) as file:
        csv_reader = csv.DictReader(file)
        if 'ITEM' not in csv_reader.fieldnames:
            return None
        item_list = [row['ITEM'] for row in csv_reader]
    unique_item_count = len(set(item_list))
    return unique_item_count


# ---------------------------------------
# Main
# ---------------------------------------
# ディレクトリを指定 TODO 自分の環境を指定
directory = '/Users/XXX/XXXX/createFile-python'

# 結果を格納するデータフレームを作成
result_df = pd.DataFrame(columns=['ファイル名', '行数', 'アイテム数'])

# 指定したディレクトリ内のファイルをループ
for filename in os.listdir(directory):
    file_path = os.path.join(directory, filename)
    
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        # エクセルファイルの行数とユニークなITEM数を取得
        row_count = get_excel_row_count(file_path)
        unique_item_count = get_excel_unique_item_count(file_path)

    elif filename.endswith(".csv"):
        # CSVファイルの行数とユニークなITEM数を取得
        try:
            row_count = get_csv_row_count(file_path, encoding='utf-8')
            unique_item_count = get_csv_unique_item_count(file_path, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                row_count = get_csv_row_count(file_path, encoding='shift_jis')
                unique_item_count = get_csv_unique_item_count(file_path, encoding='shift_jis')
            except UnicodeDecodeError:
                try:
                    row_count = get_csv_row_count(file_path, encoding='cp932')
                    unique_item_count = get_csv_unique_item_count(file_path, encoding='cp932')
                except UnicodeDecodeError:
                    try:
                        row_count = get_csv_row_count(file_path, encoding='iso-2022-jp')
                        unique_item_count = get_csv_unique_item_count(file_path, encoding='iso-2022-jp')
                    except UnicodeDecodeError:
                        print(f"エンコーディングエラー: {filename}の適切なエンコーディングが見つかりませんでした。")
                        continue
    else:
        continue

    print(f"ファイル名:{filename}, 行数:{row_count}, アイテム数:{unique_item_count}")

    # データフレームに結果を追加（concatメソッドを使用）
    new_row = pd.DataFrame({'ファイル名': [filename], '行数': [row_count], 'アイテム数': [unique_item_count]})
    result_df = pd.concat([result_df, new_row], ignore_index=True)

print("ファイル取得 終了")

# 結果をエクセルファイルに書き込む
output_file = 'Files/file_check_sample.xlsx'
result_df.to_excel(output_file, index=False)

print("ファイル書き込み 終了")