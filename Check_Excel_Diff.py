#２つのエクセルファイルを比較するスクリプト
import pandas as pd
import os

# ファイルパスとシート名
file1_path = 'SACLA運転状況集計まとめ_12-3自分.xlsm'
file2_path = 'SACLA運転状況集計まとめ_12-3最終系.xlsm'
sheet_name = '24-12'  # 比較するシート名を指定

# ファイルの存在確認
if os.path.exists(file1_path) and os.path.exists(file2_path):
    # エクセルファイルから指定したシートを読み込む
    try:
        file1 = pd.read_excel(file1_path, sheet_name=sheet_name, header=None)
        file2 = pd.read_excel(file2_path, sheet_name=sheet_name, header=None)

        print(file1_path,"____________________________________________________________________")
        print(file1)
        print(file2_path,"~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print(file2)
        print("====================================================================")

        # 比較
        diff = file1.compare(file2)

        # 差分があれば表示
        if not diff.empty:
            print("差分が見つかりました:")
            print(diff)
        else:
            print("差分はありません。")

    except ValueError as e:
        print(f"エラー: 指定したシート名 '{sheet_name}' が存在しません。エラー詳細: {e}")

else:
    print("指定されたファイルのいずれかが存在しません。")