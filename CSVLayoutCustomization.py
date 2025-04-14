import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json
import re
from tkinter.scrolledtext import ScrolledText
from tkinterdnd2 import DND_FILES, TkinterDnD

class CSVLayoutTool(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("CSVレイアウト変更ツール")
        self.geometry("900x650")
        
        self.profiles = {}
        self.current_profile_name = tk.StringVar()
        self.current_file = None
        self.preview_df = None
        self.encoding = tk.StringVar(value="shift_jis")  # デフォルトはshift_jis
        self.output_encoding = tk.StringVar(value="shift_jis")  # 出力のエンコーディング

        # --- 都道府県削除関連 ---
        self.remove_prefecture_var = tk.BooleanVar(value=False)
        self.remove_prefecture_column_var = tk.StringVar()

        # --- 都道府県コード取得関連 (新規追加) ---
        self.get_pref_code_var = tk.BooleanVar(value=False)
        self.get_pref_code_source_column_var = tk.StringVar()
        self.get_pref_code_new_column_var = tk.StringVar(value="都道府県コード")        

        # --- 都道府県名とコードのマッピング (新規追加) ---
        self.PREFECTURE_CODES = {
            "北海道": "01", "青森県": "02", "岩手県": "03", "宮城県": "04", "秋田県": "05",
            "山形県": "06", "福島県": "07", "茨城県": "08", "栃木県": "09", "群馬県": "10",
            "埼玉県": "11", "千葉県": "12", "東京都": "13", "神奈川県": "14", "新潟県": "15",
            "富山県": "16", "石川県": "17", "福井県": "18", "山梨県": "19", "長野県": "20",
            "岐阜県": "21", "静岡県": "22", "愛知県": "23", "三重県": "24", "滋賀県": "25",
            "京都府": "26", "大阪府": "27", "兵庫県": "28", "奈良県": "29", "和歌山県": "30",
            "鳥取県": "31", "島根県": "32", "岡山県": "33", "広島県": "34", "山口県": "35",
            "徳島県": "36", "香川県": "37", "愛媛県": "38", "高知県": "39", "福岡県": "40",
            "佐賀県": "41", "長崎県": "42", "熊本県": "43", "大分県": "44", "宮崎県": "45",
            "鹿児島県": "46", "沖縄県": "47"
        }
        # 逆引き用 (都道府県削除用、既存)
        self.PREFECTURES = list(self.PREFECTURE_CODES.keys())

        self.create_widgets()
        self.load_profiles()
        
    def create_widgets(self):
        # メインフレーム
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 左側フレーム（プロファイル操作）
        left_frame = ttk.LabelFrame(main_frame, text="プロファイル")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=5, pady=5)
        
        # プロファイル選択
        ttk.Label(left_frame, text="プロファイル:").pack(pady=(10, 5), padx=5, anchor=tk.W)
        
        profile_frame = ttk.Frame(left_frame)
        profile_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.profile_combobox = ttk.Combobox(profile_frame, textvariable=self.current_profile_name)
        self.profile_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.profile_combobox.bind("<<ComboboxSelected>>", self.load_profile)
        
        # プロファイル操作ボタン
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(button_frame, text="新規", command=self.new_profile, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="保存", command=self.save_profile, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="削除", command=self.delete_profile, width=6).pack(side=tk.LEFT, padx=2)
        
        # 設定エリア
        settings_frame = ttk.LabelFrame(left_frame, text="設定")
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 操作リスト
        ttk.Label(settings_frame, text="操作リスト:").pack(pady=(10, 5), padx=5, anchor=tk.W)
        
        # 操作タブ
        self.operations_notebook = ttk.Notebook(settings_frame)
        self.operations_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 並べ替えタブ
        reorder_frame = ttk.Frame(self.operations_notebook)
        self.operations_notebook.add(reorder_frame, text="並べ替え")
        
        ttk.Label(reorder_frame, text="項目の並び順 (カンマ区切り):").pack(pady=(10, 5), padx=5, anchor=tk.W)
        self.reorder_text = ScrolledText(reorder_frame, height=8, width=40, wrap=tk.WORD)
        self.reorder_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 結合タブ
        merge_frame = ttk.Frame(self.operations_notebook)
        self.operations_notebook.add(merge_frame, text="結合")
        
        ttk.Label(merge_frame, text="結合設定 (新項目名:結合元項目1,結合元項目2... 区切り文字)").pack(pady=(10, 5), padx=5, anchor=tk.W)
        self.merge_text = ScrolledText(merge_frame, height=8, width=40, wrap=tk.WORD)
        self.merge_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(merge_frame, text="例: 氏名:姓,名 / 住所:都道府県,市区町村,番地").pack(pady=5, padx=5, anchor=tk.W)

        # --- 文字列抽出タブ  ---
        extract_frame = ttk.Frame(self.operations_notebook)
        self.operations_notebook.add(extract_frame, text="文字列抽出")
        ttk.Label(extract_frame, text="抽出設定 (新項目名:抽出元項目:開始位置:文字数)").pack(pady=(10, 5), padx=5, anchor=tk.W)
        self.extract_text = ScrolledText(extract_frame, height=8, width=40, wrap=tk.WORD)
        self.extract_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(extract_frame, text="例: 商品コード前半:商品コード:1:5\n   郵便番号下4桁:郵便番号:5:4").pack(pady=5, padx=5, anchor=tk.W)
        ttk.Label(extract_frame, text="※開始位置は1から数えます").pack(pady=(0,5), padx=5, anchor=tk.W)

        # 文字除去タブ
        remove_frame = ttk.Frame(self.operations_notebook)
        self.operations_notebook.add(remove_frame, text="文字除去")
        
        ttk.Label(remove_frame, text="除去設定 (項目名:除去する文字)").pack(pady=(10, 5), padx=5, anchor=tk.W)
        self.remove_text = ScrolledText(remove_frame, height=8, width=40, wrap=tk.WORD)
        self.remove_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(remove_frame, text="例: 電話番号:-,( ) / 商品名:※").pack(pady=5, padx=5, anchor=tk.W)
        
        prefecture_frame = ttk.Frame(remove_frame)
        prefecture_frame.pack(fill=tk.X, padx=5, pady=(10, 5))

        self.remove_prefecture_check = ttk.Checkbutton(
            prefecture_frame,
            text="住所項目から都道府県を削除する",
            variable=self.remove_prefecture_var,
            onvalue=True,
            offvalue=False
        )
        self.remove_prefecture_check.pack(side=tk.LEFT, anchor=tk.W)

        prefecture_entry_frame = ttk.Frame(remove_frame)
        prefecture_entry_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        ttk.Label(prefecture_entry_frame, text="対象項目名 (複数可, カンマ区切り):").pack(side=tk.LEFT, padx=(0, 5))
        self.remove_prefecture_entry = ttk.Entry(
            prefecture_entry_frame,
            textvariable=self.remove_prefecture_column_var,
            width=25
        )
        self.remove_prefecture_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 文字追加タブ
        add_frame = ttk.Frame(self.operations_notebook)
        self.operations_notebook.add(add_frame, text="文字追加")
        
        ttk.Label(add_frame, text="追加設定 (項目名:位置:追加文字)").pack(pady=(10, 5), padx=5, anchor=tk.W)
        self.add_text = ScrolledText(add_frame, height=8, width=40, wrap=tk.WORD)
        self.add_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(add_frame, text="例: 商品コード:前:A- / 価格:後:円").pack(pady=5, padx=5, anchor=tk.W)

        # --- 都道府県コード取得タブ ---
        pref_code_frame = ttk.Frame(self.operations_notebook)
        self.operations_notebook.add(pref_code_frame, text="都道府県コード")

        pref_code_check_frame = ttk.Frame(pref_code_frame)
        pref_code_check_frame.pack(fill=tk.X, padx=5, pady=(10, 5))
        self.get_pref_code_check = ttk.Checkbutton(
            pref_code_check_frame,
            text="都道府県名からコードを取得して新しい列を追加する",
            variable=self.get_pref_code_var
        )
        self.get_pref_code_check.pack(anchor=tk.W)

        pref_code_source_frame = ttk.Frame(pref_code_frame)
        pref_code_source_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(pref_code_source_frame, text="都道府県名を含む項目名:").pack(side=tk.LEFT, padx=(0, 5))
        self.get_pref_code_source_entry = ttk.Entry(
            pref_code_source_frame,
            textvariable=self.get_pref_code_source_column_var,
            width=20
        )
        self.get_pref_code_source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        pref_code_new_frame = ttk.Frame(pref_code_frame)
        pref_code_new_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(pref_code_new_frame, text="新しい項目名 (コード列):").pack(side=tk.LEFT, padx=(0, 5))
        self.get_pref_code_new_entry = ttk.Entry(
            pref_code_new_frame,
            textvariable=self.get_pref_code_new_column_var,
            width=20
        )
        self.get_pref_code_new_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 右側フレーム（プレビューと実行）
        right_frame = ttk.LabelFrame(main_frame, text="プレビューと実行")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 文字コード選択フレーム
        encoding_frame = ttk.Frame(right_frame)
        encoding_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 入力文字コード選択
        ttk.Label(encoding_frame, text="入力文字コード:").pack(side=tk.LEFT, padx=(0, 5))
        encoding_combo = ttk.Combobox(encoding_frame, textvariable=self.encoding, values=["utf-8", "shift_jis"], width=10, state="readonly")
        encoding_combo.pack(side=tk.LEFT, padx=(0, 15))
        
        # 出力文字コード選択
        ttk.Label(encoding_frame, text="出力文字コード:").pack(side=tk.LEFT, padx=(0, 5))
        output_encoding_combo = ttk.Combobox(encoding_frame, textvariable=self.output_encoding, values=["utf-8", "shift_jis"], width=10, state="readonly")
        output_encoding_combo.pack(side=tk.LEFT)
        
        # ドラッグ&ドロップエリア
        self.drop_area = ttk.LabelFrame(right_frame, text="CSVファイルをドロップ")
        self.drop_area.pack(fill=tk.X, padx=5, pady=5)
        
        self.drop_label = ttk.Label(self.drop_area, text="ここにCSVファイルをドラッグ＆ドロップ", anchor=tk.CENTER)
        self.drop_label.pack(fill=tk.X, padx=20, pady=20)
        
        # ドロップ領域の設定
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.drop)
        
        # ファイル選択ボタン
        ttk.Button(right_frame, text="ファイルを選択", command=self.select_file).pack(fill=tk.X, padx=5, pady=5)
        
        # プレビュー領域
        preview_frame = ttk.LabelFrame(right_frame, text="プレビュー")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- スクロールバーを先に配置 ---
        # スクロールバー (Vertical)
        vsb = ttk.Scrollbar(preview_frame, orient="vertical")
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # 横スクロールバー (Horizontal)
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal")
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # --- Treeviewを最後に配置し、スクロールコマンドを設定 ---
        # プレビューのツリービュー
        # Treeview作成時にスクロールコマンドを設定する
        self.tree = ttk.Treeview(preview_frame, yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        # fill=tk.BOTH と expand=True で残りのスペースを埋めるように配置
        self.tree.pack(fill=tk.BOTH, expand=True)

        # --- スクロールバーのcommandを設定 ---
        # ScrollbarのcommandにTreeviewのメソッドを設定
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # 実行ボタン (元の位置に戻す)
        ttk.Button(right_frame, text="変換して保存", command=self.process_and_save).pack(fill=tk.X, padx=5, pady=5) # この行は元の場所にあるはず

        
    def load_profiles(self):
        try:
            if os.path.exists("csv_profiles.json"):
                with open("csv_profiles.json", "r", encoding="utf-8") as f:
                    self.profiles = json.load(f)
                    
                self.profile_combobox["values"] = list(self.profiles.keys())
                if self.profiles:
                    self.current_profile_name.set(list(self.profiles.keys())[0])
                    self.load_profile(None)
        except Exception as e:
            messagebox.showerror("エラー", f"プロファイルの読み込みに失敗しました: {str(e)}")
    
    def save_profiles(self):
        try:
            with open("csv_profiles.json", "w", encoding="utf-8") as f:
                json.dump(self.profiles, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("エラー", f"プロファイルの保存に失敗しました: {str(e)}")
    
    def new_profile(self):
        profile_name = tk.simpledialog.askstring("新規プロファイル", "プロファイル名を入力してください:")
        if profile_name and profile_name.strip():
            if profile_name in self.profiles:
                messagebox.showerror("エラー", "同名のプロファイルが既に存在します")
                return
                
            self.profiles[profile_name] = {
                "reorder": "",
                "merge": "",
                "extract": "",
                "remove": "",
                "add": "",
                "remove_prefecture": {"enabled": False, "column": ""},
                "get_pref_code": {"enabled": False, "source_column": "", "new_column": "都道府県コード"}
            }
            
            self.profile_combobox["values"] = list(self.profiles.keys())
            self.current_profile_name.set(profile_name)
            self.load_profile(None)
            self.save_profiles()
    
    def save_profile(self):
        profile_name = self.current_profile_name.get()
        if not profile_name:
            messagebox.showerror("エラー", "プロファイルが選択されていません")
            return
            
        self.profiles[profile_name] = {
            "reorder": self.reorder_text.get("1.0", tk.END).strip(),
            "merge": self.merge_text.get("1.0", tk.END).strip(),
            "extract": self.extract_text.get("1.0", tk.END).strip(),
            "remove": self.remove_text.get("1.0", tk.END).strip(),
            "add": self.add_text.get("1.0", tk.END).strip(),
            "remove_prefecture": {
                "enabled": self.remove_prefecture_var.get(),
                "column": self.remove_prefecture_column_var.get()
            },
            "get_pref_code": {
                "enabled": self.get_pref_code_var.get(),
                "source_column": self.get_pref_code_source_column_var.get(),
                "new_column": self.get_pref_code_new_column_var.get()
            }
        }
        
        self.save_profiles()
        messagebox.showinfo("保存完了", f"プロファイル「{profile_name}」を保存しました")
    
    def delete_profile(self):
        profile_name = self.current_profile_name.get()
        if not profile_name:
            messagebox.showerror("エラー", "プロファイルが選択されていません")
            return
            
        if messagebox.askyesno("確認", f"プロファイル「{profile_name}」を削除しますか？"):
            del self.profiles[profile_name]
            self.profile_combobox["values"] = list(self.profiles.keys())
            
            if self.profiles:
                self.current_profile_name.set(list(self.profiles.keys())[0])
                self.load_profile(None)
            else:
                self.current_profile_name.set("")
                self.reorder_text.delete("1.0", tk.END)
                self.merge_text.delete("1.0", tk.END)
                self.extract_text.delete("1.0", tk.END)
                self.remove_text.delete("1.0", tk.END)
                self.add_text.delete("1.0", tk.END)
                self.remove_prefecture_var.set(False)
                self.remove_prefecture_column_var.set("")
                self.get_pref_code_var.set(False)
                self.get_pref_code_source_column_var.set("")
                self.get_pref_code_new_column_var.set("都道府県コード")                
                
            self.save_profiles()
    
    def load_profile(self, event):
        profile_name = self.current_profile_name.get()
        if not profile_name or profile_name not in self.profiles:
            return
            
        profile = self.profiles[profile_name]
        
        # 既存の設定読み込み
        self.reorder_text.delete("1.0", tk.END)
        self.reorder_text.insert(tk.END, profile.get("reorder", ""))
        self.merge_text.delete("1.0", tk.END)
        self.merge_text.insert(tk.END, profile.get("merge", ""))
        self.extract_text.delete("1.0", tk.END)
        self.extract_text.insert(tk.END, profile.get("extract", ""))
        self.remove_text.delete("1.0", tk.END)
        self.remove_text.insert(tk.END, profile.get("remove", ""))
        self.add_text.delete("1.0", tk.END)
        self.add_text.insert(tk.END, profile.get("add", ""))

        # 都道府県削除設定の読み込み
        remove_pref_settings = profile.get("remove_prefecture", {"enabled": False, "column": ""}) # デフォルト値
        self.remove_prefecture_var.set(remove_pref_settings.get("enabled", False))
        self.remove_prefecture_column_var.set(remove_pref_settings.get("column", ""))

        # 都道府県コード設定の読み込み
        get_pref_code_settings = profile.get("get_pref_code", {"enabled": False, "source_column": "", "new_column": "都道府県コード"})
        self.get_pref_code_var.set(get_pref_code_settings.get("enabled", False))
        self.get_pref_code_source_column_var.set(get_pref_code_settings.get("source_column", ""))
        self.get_pref_code_new_column_var.set(get_pref_code_settings.get("new_column", "都道府県コード"))

        # 現在のファイルがあれば再プレビュー
        if self.current_file:
            self.preview_file(self.current_file)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        
        if file_path:
            self.preview_file(file_path)
    
    def drop(self, event):
        file_path = event.data
        
        # Windows形式のパスを修正
        file_path = file_path.replace("{", "").replace("}", "")
        
        # ダブルクォートを除去
        if file_path.startswith('"') and file_path.endswith('"'):
            file_path = file_path[1:-1]
        
        if os.path.isfile(file_path) and file_path.lower().endswith('.csv'):
            self.preview_file(file_path)
        else:
            messagebox.showerror("エラー", "有効なCSVファイルではありません")
    
    def preview_file(self, file_path):
        try:
            # 選択されたエンコーディングでファイルを読み込み
            selected_encoding = self.encoding.get()
            
            # エンコーディングでのエラーをキャッチするためのエラーハンドリング
            try:
                df = pd.read_csv(file_path, encoding=selected_encoding)
            except UnicodeDecodeError:
                # もし選択したエンコーディングでエラーが発生したら、代替を試行
                alternative_encoding = "shift_jis" if selected_encoding == "utf-8" else "utf-8"
                try:
                    df = pd.read_csv(file_path, encoding=alternative_encoding)
                    # 成功したら元のエンコーディング変数を更新
                    self.encoding.set(alternative_encoding)
                    messagebox.showinfo("エンコーディング変更", 
                                        f"選択されたエンコーディング({selected_encoding})では読み込めませんでした。\n"
                                        f"代わりに{alternative_encoding}で読み込みました。")
                except UnicodeDecodeError:
                    # 両方のエンコーディングが失敗した場合
                    messagebox.showerror("エラー", "UTF-8とShift-JISのどちらでもファイルを読み込めませんでした。\n"
                                                "別のCSVファイルを試してください。")
                    return
            
            # 現在のファイルとして設定
            self.current_file = file_path
            
            # タイトル更新
            self.title(f"CSVレイアウト変更ツール - {os.path.basename(file_path)}")
            
            # プレビューの作成
            self.preview_df = self.process_dataframe(df)
            self.update_preview()
            
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {str(e)}")
    
    def process_dataframe(self, df):

        # 除去対象都道府県リスト
        PREFECTURES = [
            "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
            "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
            "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
            "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
            "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
            "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
            "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"
        ]

        try:
            # 元のデータフレームをコピー
            result_df = df.copy()

            # --- 都道府県コード取得処理 (新規追加 - 他の処理より先に行うのが自然) ---
            if self.get_pref_code_var.get():
                source_col = self.get_pref_code_source_column_var.get().strip()
                new_col = self.get_pref_code_new_column_var.get().strip()

                if source_col and new_col: # ソース列名と新しい列名の両方が指定されている場合
                    if source_col in result_df.columns:
                        if new_col in result_df.columns:
                             # 新しい列名が既存の列と重複する場合は警告（上書きはしない）
                             print(f"警告: 新しい都道府県コード列名 '{new_col}' は既に存在します。別の名前を指定してください。")
                        else:
                            # 都道府県コードを取得する関数
                            def get_code(address):
                                if isinstance(address, str):
                                    for pref, code in self.PREFECTURE_CODES.items():
                                        if address.startswith(pref):
                                            return code
                                return "" # 見つからない場合は空文字

                            # 新しい列を追加してコードを格納
                            result_df[new_col] = result_df[source_col].apply(get_code)
                    else:
                        print(f"警告: 都道府県コード取得のソース列 '{source_col}' が見つかりません。")
                elif not source_col:
                     print("警告: 都道府県コード取得のソース列名が指定されていません。")
                elif not new_col:
                     print("警告: 都道府県コード取得の新しい列名が指定されていません。")

            # --- 都道府県名削除処理 ---
            if self.remove_prefecture_var.get():
                # カンマ区切りで項目名を取得し、前後の空白を除去
                target_columns_str = self.remove_prefecture_column_var.get().strip()
                if target_columns_str: # 入力がある場合のみ処理
                    target_columns = [col.strip() for col in target_columns_str.split(',') if col.strip()] # 空の項目名を除外

                    # 都道府県削除関数
                    def remove_pref(address):
                        if isinstance(address, str):
                            for pref in PREFECTURES:
                                if address.startswith(pref):
                                    return address[len(pref):]
                        return address

                    # 各対象列に対して処理を実行
                    for target_column in target_columns:
                        if target_column in result_df.columns:
                            result_df[target_column] = result_df[target_column].apply(remove_pref)
                        else:
                            # 対象列が存在しない場合、警告を表示
                            print(f"警告: 都道府県削除の対象列 '{target_column}' が見つかりません。")
            
            # --- 結合処理 ---
            merge_settings = self.merge_text.get("1.0", tk.END).strip()
            if merge_settings:
                for line in merge_settings.split('\n'):
                    if not line.strip():
                        continue
                        
                    parts = line.split(':')
                    if len(parts) >= 2:
                        new_column = parts[0].strip()
                        merge_info = ':'.join(parts[1:])
                        
                        # merge_info を ':' で分割した後の処理
                        merge_info_parts = merge_info.split(' ', 1)
                        source_columns_str = merge_info_parts[0]
                        source_columns = [col.strip() for col in source_columns_str.split(',')]

                        # separator_parts が存在するかどうかで区切り文字を決定
                        # merge_info_parts の長さが 2 以上なら区切り文字が指定されている
                        separator = merge_info_parts[1].strip() if len(merge_info_parts) > 1 else ''

                        # すべてのソース列が存在するか確認
                        if all(col in result_df.columns for col in source_columns):
                            result_df[new_column] = result_df[source_columns].apply(
                                lambda x: separator.join([str(item) if pd.notna(item) else '' for item in x]), axis=1
                            )

            # --- 文字列抽出処理 ---
            extract_settings = self.extract_text.get("1.0", tk.END).strip()
            if extract_settings:
                for line in extract_settings.split('\n'):
                    line = line.strip()
                    if not line: continue
                    try:
                        parts = line.split(':', 3) # 新項目名:抽出元項目:開始位置:文字数
                        if len(parts) == 4:
                            new_col = parts[0].strip()
                            source_col = parts[1].strip()
                            start_pos_str = parts[2].strip()
                            num_chars_str = parts[3].strip()

                            if not new_col:
                                print(f"警告: 文字列抽出設定で新しい項目名が空です: {line}")
                                continue
                            if not source_col:
                                print(f"警告: 文字列抽出設定で抽出元項目が指定されていません: {line}")
                                continue
                            if source_col not in result_df.columns:
                                print(f"警告: 文字列抽出の抽出元項目 '{source_col}' が見つかりません。")
                                continue
                            if new_col in result_df.columns:
                                print(f"警告: 文字列抽出の新しい項目名 '{new_col}' は既に存在します。上書きします。") # またはスキップ

                            # 開始位置と文字数を数値に変換
                            try:
                                start_pos = int(start_pos_str)
                                num_chars = int(num_chars_str)
                                if start_pos < 1:
                                    print(f"警告: 文字列抽出の開始位置は1以上である必要があります: {line}")
                                    continue
                                if num_chars < 0: # 0文字抽出は空文字になるので許容しても良いかも
                                    print(f"警告: 文字列抽出の文字数は0以上である必要があります: {line}")
                                    continue
                            except ValueError:
                                print(f"警告: 文字列抽出の開始位置または文字数が数値ではありません: {line}")
                                continue

                            # スライス処理 (開始位置は1ベースなので-1する)
                            # pandas の str.slice は stop に終了インデックス+1 を指定する
                            start_index = start_pos - 1
                            end_index = start_index + num_chars

                            # apply を使う方法 (NaNを安全に処理)
                            def extract_substring(text):
                                if pd.isna(text):
                                    return "" # NaNの場合は空文字を返す
                                text_str = str(text) # 数値なども文字列に変換
                                if start_index >= len(text_str):
                                    return "" # 開始位置が文字列長を超える場合は空文字
                                return text_str[start_index:end_index]

                            result_df[new_col] = result_df[source_col].apply(extract_substring)

                        else:
                            print(f"警告: 文字列抽出設定の形式が正しくありません ('新項目名:抽出元項目:開始位置:文字数') : {line}")
                    except Exception as e:
                        print(f"エラー: 文字列抽出処理中にエラーが発生しました ({line}): {e}")

            # --- 文字除去処理　---
            remove_settings = self.remove_text.get("1.0", tk.END).strip()
            if remove_settings:
                for line in remove_settings.split('\n'):
                    if not line.strip():
                        continue
                        
                    parts = line.split(':')
                    if len(parts) >= 2:
                        column = parts[0].strip()
                        chars_to_remove = parts[1].strip().split(',')
                        
                        if column in result_df.columns:
                            for char in chars_to_remove:
                                result_df[column] = result_df[column].astype(str).str.replace(char.strip(), '', regex=False)
            
            # --- 文字追加処理　---
            add_settings = self.add_text.get("1.0", tk.END).strip()
            if add_settings:
                for line in add_settings.split('\n'):
                    if not line.strip():
                        continue
                        
                    parts = line.split(':')
                    if len(parts) >= 3:
                        column = parts[0].strip()
                        position = parts[1].strip()
                        chars_to_add = parts[2].strip()
                        
                        if column in result_df.columns:
                            if position == "前":
                                result_df[column] = chars_to_add + result_df[column].astype(str)
                            elif position == "後":
                                result_df[column] = result_df[column].astype(str) + chars_to_add
            
            # --- 列の並べ替え ---
            reorder_settings = self.reorder_text.get("1.0", tk.END).strip()
            if reorder_settings:
                columns = [col.strip() for col in reorder_settings.split(',')]
                
                # 指定された列のみを対象に
                valid_columns = [col for col in columns if col in result_df.columns]
                
                # 指定されていない列を最後に追加
                remaining_columns = [col for col in result_df.columns if col not in valid_columns]
                
                # 最終的な列順序
                final_columns = valid_columns + remaining_columns
                
                # 並べ替え
                result_df = result_df[final_columns]
            
            return result_df
            
        except Exception as e:
            messagebox.showerror("エラー", f"データ処理中にエラーが発生しました: {str(e)}")
            return df
    
    def update_preview(self):
        if self.preview_df is None:
            return
            
        # ツリービューをクリア
        self.tree.delete(*self.tree.get_children())
        
        # 列の設定
        self.tree["columns"] = list(self.preview_df.columns)
        self.tree["show"] = "headings"
        
        # 各列の設定
        for col in self.preview_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        # データの追加（最大10行）
        preview_rows = min(10, len(self.preview_df))
        for i in range(preview_rows):
            values = list(self.preview_df.iloc[i])
            # 数値を文字列に変換
            values = [str(val) for val in values]
            self.tree.insert("", tk.END, values=values)
            
        if len(self.preview_df) > 10:
            self.tree.insert("", tk.END, values=["..."] * len(self.preview_df.columns))
    
    def process_and_save(self):
        if not self.current_file:
            messagebox.showerror("エラー", "処理するCSVファイルが選択されていません")
            return
            
        try:
            # 元のファイルを読み込み
            df = pd.read_csv(self.current_file, encoding=self.encoding.get())
            
            # 処理
            processed_df = self.process_dataframe(df)
            
            # 保存先を選択
            base_name = os.path.basename(self.current_file)
            name, ext = os.path.splitext(base_name)
            
            output_path = filedialog.asksaveasfilename(
                title="変換後のファイルを保存",
                initialfile=f"{name}_converted{ext}",
                filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
            )
            
            if output_path:
                # 選択された出力エンコーディングで保存
                processed_df.to_csv(output_path, index=False, encoding=self.output_encoding.get())
                messagebox.showinfo("成功", f"ファイルを保存しました: {output_path}")
                
        except Exception as e:
            messagebox.showerror("エラー", f"処理に失敗しました: {str(e)}")

if __name__ == "__main__":
    app = CSVLayoutTool()
    app.mainloop()
