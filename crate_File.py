import pandas as pd

# ---------------------------------------
# Excelファイル作成
# 10アイテム（1列目）固定×複数行
# created: 2023/04/19
# ---------------------------------------


# 列名をリストで指定
column_names = ['ITEM', 'CATEGORY', 'COLOR', 'PRICE', 'COST', 'SUBSIDIARY MATERIALS', 'QUANTITY', 'MEMO']

# ユニークな内容のリスト (「ITEM」列を除いたリスト)
unique_data = [
    ['tops','red','4000','','button','100',''],
    ['bottoms','navy','9000','','zipper','150',''],
    ['one piece','white','9000','','button','200',''],
    ['shoes','black','12000','','shoelace','80',''],
    ['cap','beige','3000','','embroidery thread','30',''],
]

# ITEMのリスト (各ITEMは10行ずつ繰り返し)
item_list = [
    '10114','10745','12369','11644','10345','11910','12444','10545','10880','10137',
]

# DataFrameを作成
df_list = []

print(f'itemの数：{len(item_list)}')

# 10行ずつデータを追加
for item in item_list:
    temp_df = pd.DataFrame(unique_data, columns=column_names[1:])
    temp_df.insert(0, 'ITEM', item)
    df_list.append(temp_df)

    print(f'item:{item}')

# DataFrameを連結（concat）
df = pd.concat(df_list, ignore_index=True)

# Excelファイルとして出力
df.to_excel('Files/sample.xlsx', index=False, engine='openpyxl')
print('ファイル出力 完了')