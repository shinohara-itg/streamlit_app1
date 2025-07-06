import pandas as pd
import re
import os

def process_excel(excel_path: str, base_dir: str):
    # --- CSV書き出し用フォルダ ---
    output_dir = os.path.join(base_dir, "1 datatable")
    os.makedirs(output_dir, exist_ok=True)

    # --- テキスト出力用フォルダ ---
    text_output_dir = os.path.join(base_dir, "2 output")
    os.makedirs(text_output_dir, exist_ok=True)

    print(f"▶ 処理対象ファイル：{excel_path}")

    if not os.path.isfile(excel_path):
        raise FileNotFoundError(f"指定されたファイルが存在しません：{excel_path}")

    xls = pd.ExcelFile(excel_path)
    sheet_names = xls.sheet_names
    pattern = re.compile(r"^\d")
    target_sheets = [name for name in sheet_names if pattern.match(name)]

    for sheet in target_sheets:
        print(f"\n▶ 処理中シート：{sheet}")
        df_raw = pd.read_excel(excel_path, sheet_name=sheet, header=None)
        df = None

        # Q列形式
        if df_raw.shape[0] > 13 and df_raw.shape[1] > 17 and df_raw.iloc[13, 16] == "TOTAL":
            try:
                df = df_raw.iloc[12:].copy()
                columns_to_drop = list(range(0, 16)) + [17]
                df.drop(columns=columns_to_drop, inplace=True, errors='ignore')
                df.columns = df.iloc[0]
                df = df[1:]
                columns = list(df.columns)
                columns[0] = "セグメント名"
                df.columns = columns
            except Exception as e:
                print(f"⚠ エラー（Q列形式）：{e}")
                continue

        # B列形式
        elif df_raw.shape[0] > 22 and df_raw.shape[1] > 1 and df_raw.iloc[22, 1] == "TOTAL":
            try:
                df = df_raw.iloc[21:, [1] + list(range(3, df_raw.shape[1]))]
                df.columns = df.iloc[0]
                df = df[1:]
                columns = list(df.columns)
                columns[0] = "セグメント名"
                df.columns = columns
            except Exception as e:
                print(f"⚠ エラー（B列形式）：{e}")
                continue
        else:
            print("→ スキップ：該当形式なし（TOTALが見つからない）")
            continue

        if "TOTAL_x000d_" in df.columns[-1]:
            df.drop(columns=[df.columns[-1]], inplace=True)

        output_path = os.path.join(output_dir, f"{sheet}.csv")
        try:
            df.to_csv(output_path, index=False, encoding="utf-8-sig")
            print(f"✅ 保存完了：{output_path}")
        except Exception as e:
            print(f"❌ 保存失敗：{output_path} - {e}")

    # --- INDEXシート処理 ---
    try:
        index_df = pd.read_excel(excel_path, sheet_name="INDEX", header=None)
        for i in range(2, index_df.shape[0]):
            file_name = index_df.iloc[i, 1]
            question_text = index_df.iloc[i, 3]
            if pd.notna(file_name) and pd.notna(question_text):
                content = f"【分析対象】\n設問「{question_text}」に関するクロス集計の分析"
                text_path = os.path.join(text_output_dir, f"{file_name}.txt")
                with open(text_path, "w", encoding="utf-8") as f:
                    f.write(content)
                print(f"📝 テキスト出力：{text_path}")
    except Exception as e:
        print(f"⚠ INDEXシート処理エラー：{e}")

    print("\n🎉 すべての処理が完了しました。")
