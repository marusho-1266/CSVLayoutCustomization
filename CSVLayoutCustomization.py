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

        # --- ヘッダー除去関連 (新規追加) ---
        self.remove_header_var = tk.BooleanVar(value=False)

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

        # --- 空列マッピング用の一時変数 ---
        self._empty_col_mapping = {}

        # ウィジェット作成メソッドを呼び出す前にプレビュー関連の変数を初期化
        self.preview_frame = None
        self.tree = None
        self.vsb = None
        self.hsb = None

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
        
        # --- プレビュー領域のウィジェット ---
        self._create_preview_widgets(right_frame)

        # --- ヘッダー除去チェックボックス ---
        header_frame = ttk.Frame(right_frame)
        header_frame.pack(fill=tk.X, padx=5, pady=(5, 0)) # ボタンの少し上に配置
        self.remove_header_check = ttk.Checkbutton(
            header_frame,
            text="出力時にヘッダー行を除去する",
            variable=self.remove_header_var,
            onvalue=True,
            offvalue=False
        )
        self.remove_header_check.pack(side=tk.LEFT, anchor=tk.W)

        # 実行ボタン (元の位置に戻す)
        ttk.Button(right_frame, text="変換して保存", command=self.process_and_save).pack(fill=tk.X, padx=5, pady=5) # この行は元の場所にあるはず

    # --- プレビューウィジェット作成用メソッド ---
    def _create_preview_widgets(self, parent_frame):
        # プレビュー領域フレーム (インスタンス変数に格納)
        self.preview_frame = ttk.LabelFrame(parent_frame, text="プレビュー")
        self.preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # スクロールバー (インスタンス変数に格納)
        self.vsb = ttk.Scrollbar(self.preview_frame, orient="vertical")
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.hsb = ttk.Scrollbar(self.preview_frame, orient="horizontal")
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Treeview (インスタンス変数に格納)
        self.tree = ttk.Treeview(self.preview_frame, yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # スクロールバーのcommand設定
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

    # --- Treeviewを再生成するメソッド ---
    def _recreate_treeview(self):
        if self.tree:
            self.tree.destroy() # 古いTreeviewを破棄

        # 新しいTreeviewを作成し、インスタンス変数 self.tree を更新
        self.tree = ttk.Treeview(self.preview_frame, yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.tree.pack(fill=tk.BOTH, expand=True) # 再度packする

        # スクロールバーのcommandを新しいTreeviewに再設定
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

        # 新しいTreeviewをスクロールバーの前に表示させる (念のため)
        self.vsb.lift(self.tree)
        self.hsb.lift(self.tree)
        
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
                "get_pref_code": {"enabled": False, "source_column": "", "new_column": "都道府県コード"},
                "remove_header": False
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
            },
            "remove_header": self.remove_header_var.get()
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
                # Clear settings fields
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
                self.remove_header_var.set(False)

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

        # ヘッダー除去設定の読み込み
        self.remove_header_var.set(profile.get("remove_header", False))

        # 現在のファイルがあれば再プレビュー
        if self.current_file:
            self.preview_file(self.current_file)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            # --- ファイル選択時も drop と同様に after_idle を使う ---
            self.after_idle(lambda: self._clear_and_preview_logic(file_path))
    
    def drop(self, event):
        try:
            file_path = event.data
            file_path = file_path.replace("{", "").replace("}", "")
            if file_path.startswith('"') and file_path.endswith('"'):
                file_path = file_path[1:-1]

            if not os.path.isfile(file_path):
                messagebox.showerror("エラー", "有効なファイルではありません")
                return
            if not file_path.lower().endswith('.csv'):
                messagebox.showerror("エラー", "CSVファイルのみ対応しています")
                return

            # --- 修正箇所: Treeview再生成とプレビュー処理を遅延実行 ---
            self.after_idle(lambda: self._clear_and_preview_logic(file_path))
            # --- 修正箇所 終了 ---

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルのドロップ処理中にエラーが発生しました: {str(e)}")
            # エラー発生時のクリーンアップ (ここも after_idle が安全)
            self.after_idle(self._cleanup_on_error)


    # --- 新しいロジック用メソッド ---
    def _clear_and_preview_logic(self, file_path):
        """Treeviewを再生成し、新しいファイルをプレビューする"""
        try:
            # Treeviewを再生成
            self._recreate_treeview()

            # 現在のファイルとプレビューデータをクリア
            self.current_file = None
            self.preview_df = None

            # 新しいファイルをプレビュー
            self.preview_file(file_path)

        except Exception as task_error:
            messagebox.showerror("エラー", f"ファイル処理タスク中にエラーが発生しました: {str(task_error)}")
            self._cleanup_on_error()

    def _cleanup_on_error(self):
        """エラー発生時に状態をクリアし、Treeviewも再生成する"""
        self.current_file = None
        self.preview_df = None
        try:
            # エラー時もTreeviewを再生成してクリーンな状態にする
            self._recreate_treeview()
        except Exception as cleanup_error:
             print(f"警告: エラー後のTreeview再生成中にエラー: {cleanup_error}")
    
    def preview_file(self, file_path):
        # --- preview_file 内の update_preview() 呼び出し前のクリア処理は不要 ---
        # --- なぜなら _clear_and_preview_logic で Treeview が再生成されるため ---
        try:
            selected_encoding = self.encoding.get()
            df = None
            try:
                df = pd.read_csv(file_path, encoding=selected_encoding)
            except UnicodeDecodeError:
                alternative_encoding = "shift_jis" if selected_encoding == "utf-8" else "utf-8"
                try:
                    df = pd.read_csv(file_path, encoding=alternative_encoding)
                    self.encoding.set(alternative_encoding)
                    messagebox.showinfo("エンコーディング変更",
                                        f"選択されたエンコーディング({selected_encoding})では読み込めませんでした。\n"
                                        f"代わりに{alternative_encoding}で読み込みました。")
                except UnicodeDecodeError:
                    messagebox.showerror("エラー", "UTF-8とShift-JISのどちらでもファイルを読み込めませんでした。\n"
                                                "別のCSVファイルを試してください。")
                    self._cleanup_on_error() # エラー時はクリーンアップ
                    return
                except Exception as inner_e: # その他の読み込みエラー
                     messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました ({alternative_encoding}): {str(inner_e)}")
                     self._cleanup_on_error()
                     return
            except Exception as outer_e: # その他の読み込みエラー
                 messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました ({selected_encoding}): {str(outer_e)}")
                 self._cleanup_on_error()
                 return

            # dfが正常に読み込めた場合のみ続行
            if df is None:
                # 通常ここには来ないはずだが念のため
                self._cleanup_on_error()
                return

            self.current_file = file_path
            self.title(f"CSVレイアウト変更ツール - {os.path.basename(file_path)}")

            # プレビューの作成 (process_dataframe は df を受け取る)
            self.preview_df = self.process_dataframe(df) # df を渡す
            self.update_preview() # update_preview は self.preview_df を使う

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルプレビュー処理全体でエラーが発生しました: {str(e)}")
            self._cleanup_on_error() # プレビュー処理中のエラーでもクリーンアップ
    
    def process_dataframe(self, df: pd.DataFrame) -> pd.DataFrame: 

        try:
            result_df = df.copy()

            # --- 都道府県コード取得処理 ---
            if self.get_pref_code_var.get():
                source_col = self.get_pref_code_source_column_var.get().strip()
                new_col = self.get_pref_code_new_column_var.get().strip()
                if source_col and new_col:
                    if source_col in result_df.columns:
                        if new_col in result_df.columns:
                             print(f"警告: 新しい都道府県コード列名 '{new_col}' は既に存在します。") # messagebox.showwarning の方が良いかも
                        else:
                            def get_code(address):
                                if isinstance(address, str):
                                    for pref, code in self.PREFECTURE_CODES.items():
                                        if address.startswith(pref):
                                            return code
                                return ""
                            result_df[new_col] = result_df[source_col].apply(get_code)
                    else:
                        print(f"警告: 都道府県コード取得のソース列 '{source_col}' が見つかりません。") # messagebox.showwarning
                # ... (ソース列名、新列名がない場合の警告) ...

            # --- 都道府県名削除処理 ---
            if self.remove_prefecture_var.get():
                target_columns_str = self.remove_prefecture_column_var.get().strip()
                if target_columns_str:
                    target_columns = [col.strip() for col in target_columns_str.split(',') if col.strip()]
                    def remove_pref(address):
                        if isinstance(address, str):
                            # --- self.PREFECTURES を使用 ---
                            for pref in self.PREFECTURES:
                                if address.startswith(pref):
                                    return address[len(pref):]
                        return address
                    for target_column in target_columns:
                        if target_column in result_df.columns:
                            result_df[target_column] = result_df[target_column].apply(remove_pref)
                        else:
                            print(f"警告: 都道府県削除の対象列 '{target_column}' が見つかりません。") # messagebox.showwarning

            # --- 結合処理 ---
            merge_settings = self.merge_text.get("1.0", tk.END).strip()
            if merge_settings:
                for line in merge_settings.split('\n'):
                    line = line.strip()
                    if not line: continue
                    try: # 設定行のパースエラーをキャッチ
                        parts = line.split(':', 1) # 新項目名と残りで分割
                        if len(parts) == 2:
                            new_column = parts[0].strip()
                            merge_info = parts[1].strip()

                            # 区切り文字を特定 (最後のスペースで分割)
                            separator = ''
                            source_columns_str = merge_info
                            if ' ' in merge_info:
                                parts_merge = merge_info.rsplit(' ', 1)
                                source_columns_str = parts_merge[0].strip()
                                separator = parts_merge[1].strip() # 区切り文字はスペースを含む可能性あり

                            source_columns = [col.strip() for col in source_columns_str.split(',') if col.strip()]

                            if not new_column:
                                print(f"警告: 結合設定で新項目名が空です: {line}")
                                continue
                            if not source_columns:
                                print(f"警告: 結合設定で結合元項目が指定されていません: {line}")
                                continue

                            missing_cols = [col for col in source_columns if col not in result_df.columns]
                            if missing_cols:
                                print(f"警告: 結合の元項目が見つかりません: {', '.join(missing_cols)} (設定: {line})")
                                continue

                            if new_column in result_df.columns:
                                print(f"警告: 結合先の項目名 '{new_column}' は既に存在します。上書きします。")

                            # 結合実行
                            result_df[new_column] = result_df[source_columns].apply(
                                lambda x: separator.join([str(item) if pd.notna(item) else '' for item in x]), axis=1
                            )
                        else:
                             print(f"警告: 結合設定の形式が正しくありません (':') : {line}")
                    except Exception as merge_ex:
                        print(f"エラー: 結合処理中にエラーが発生しました ({line}): {merge_ex}")


            # --- 文字列抽出処理 ---
            extract_settings = self.extract_text.get("1.0", tk.END).strip()
            if extract_settings:
                for line in extract_settings.split('\n'):
                    line = line.strip()
                    if not line: continue
                    try:
                        parts = line.split(':', 3)
                        if len(parts) == 4:
                            new_col, source_col, start_pos_str, num_chars_str = [p.strip() for p in parts]

                            if not new_col: print(f"警告: 文字列抽出設定で新項目名が空: {line}"); continue
                            if not source_col: print(f"警告: 文字列抽出設定で抽出元項目が空: {line}"); continue
                            if source_col not in result_df.columns: print(f"警告: 文字列抽出の抽出元項目 '{source_col}' が見つかりません。"); continue
                            if new_col in result_df.columns: print(f"警告: 文字列抽出の新項目名 '{new_col}' は既に存在。上書きします。")

                            try:
                                start_pos = int(start_pos_str)
                                num_chars = int(num_chars_str)
                                if start_pos < 1: print(f"警告: 文字列抽出の開始位置は1以上: {line}"); continue
                                if num_chars < 0: print(f"警告: 文字列抽出の文字数は0以上: {line}"); continue
                            except ValueError: print(f"警告: 文字列抽出の開始位置/文字数が数値でない: {line}"); continue

                            start_index = start_pos - 1
                            end_index = start_index + num_chars

                            def extract_substring(text):
                                if pd.isna(text): return ""
                                text_str = str(text)
                                if start_index >= len(text_str): return ""
                                return text_str[start_index:end_index]

                            result_df[new_col] = result_df[source_col].apply(extract_substring)
                        else:
                            print(f"警告: 文字列抽出設定の形式が不正 ('新:元:開始:文字数'): {line}")
                    except Exception as extract_ex:
                        print(f"エラー: 文字列抽出処理中にエラー ({line}): {extract_ex}")

            # --- 文字除去処理 ---
            remove_settings = self.remove_text.get("1.0", tk.END).strip()
            if remove_settings:
                for line in remove_settings.split('\n'):
                    line = line.strip()
                    if not line: continue
                    try:
                        parts = line.split(':', 1)
                        if len(parts) == 2:
                            column = parts[0].strip()
                            # 除去文字はカンマ区切りではなく、指定された文字列全体として扱う方が自然かも？
                            # 例: "TEL:(03)1234-5678" から "TEL:", "(", ")", "-" を除去したい場合
                            # 設定例: 電話番号:TEL:,(,),- のようにするか、
                            # 設定例: 電話番号:TEL:():- のようにするか。後者の方が直感的か。
                            # ここでは元のカンマ区切りを維持するが、検討の余地あり。
                            chars_to_remove_str = parts[1].strip()
                            chars_to_remove_list = [c.strip() for c in chars_to_remove_str.split(',') if c.strip()] # カンマ区切り

                            if not column: print(f"警告: 文字除去設定で項目名が空: {line}"); continue
                            if not chars_to_remove_list: print(f"警告: 文字除去設定で除去文字が空: {line}"); continue
                            if column not in result_df.columns: print(f"警告: 文字除去の項目 '{column}' が見つかりません。"); continue

                            # 正規表現のエスケープが必要な文字も考慮するなら replace より re.sub が安全
                            # result_df[column] = result_df[column].astype(str)
                            # for char in chars_to_remove_list:
                            #     result_df[column] = result_df[column].str.replace(re.escape(char), '', regex=True)
                            # regex=False の方が意図通りならそのままでOK
                            for char in chars_to_remove_list:
                                result_df[column] = result_df[column].astype(str).str.replace(char, '', regex=False)
                        else:
                            print(f"警告: 文字除去設定の形式が不正 (':'): {line}")
                    except Exception as remove_ex:
                        print(f"エラー: 文字除去処理中にエラー ({line}): {remove_ex}")


            # --- 文字追加処理 ---
            add_settings = self.add_text.get("1.0", tk.END).strip()
            if add_settings:
                for line in add_settings.split('\n'):
                    line = line.strip()
                    if not line: continue
                    try:
                        parts = line.split(':', 2) # 項目名:位置:追加文字
                        if len(parts) == 3:
                            column, position, chars_to_add = [p.strip() for p in parts]

                            if not column: print(f"警告: 文字追加設定で項目名が空: {line}"); continue
                            if position not in ["前", "後"]: print(f"警告: 文字追加の位置は '前' または '後': {line}"); continue
                            # chars_to_add は空でも許可する（空文字を追加する意味はないがエラーではない）
                            if column not in result_df.columns: print(f"警告: 文字追加の項目 '{column}' が見つかりません。"); continue

                            if position == "前":
                                result_df[column] = chars_to_add + result_df[column].astype(str)
                            elif position == "後":
                                result_df[column] = result_df[column].astype(str) + chars_to_add
                        else:
                            print(f"警告: 文字追加設定の形式が不正 ('項目:位置:追加文字'): {line}")
                    except Exception as add_ex:
                        print(f"エラー: 文字追加処理中にエラー ({line}): {add_ex}")


            # --- 列の並べ替え ---
            reorder_settings = self.reorder_text.get("1.0", tk.END).strip()
            self._empty_col_mapping = {} # 並べ替え前にクリア
            if reorder_settings:
                specified_columns_with_blanks = [col.strip() for col in reorder_settings.split(',')]
                final_columns = []
                new_empty_cols_mapping = {}
                empty_col_counter = 1

                current_columns = list(result_df.columns) # 現在のDFの列リストを取得

                for col_name in specified_columns_with_blanks:
                    if col_name:
                        # 大文字小文字を区別せずにチェックする方が親切かもしれないが、現状維持
                        if col_name in current_columns:
                            final_columns.append(col_name)
                        else:
                            print(f"警告: 並べ替え指定の列 '{col_name}' は現在のデータに存在しません。")
                    else: # 空列を追加
                        placeholder_name = f"__EMPTY_COLUMN_{empty_col_counter}__"
                        while placeholder_name in current_columns or placeholder_name in new_empty_cols_mapping:
                            empty_col_counter += 1
                            placeholder_name = f"__EMPTY_COLUMN_{empty_col_counter}__"

                        result_df[placeholder_name] = "" # 空列をDFに追加
                        final_columns.append(placeholder_name)
                        new_empty_cols_mapping[placeholder_name] = ''
                        # current_columns にも追加しておく (後続の空列チェックのため)
                        current_columns.append(placeholder_name)
                        empty_col_counter += 1

                if final_columns:
                    try:
                        # 存在しない列を指定していないか最終チェック (念のため)
                        valid_final_columns = [col for col in final_columns if col in result_df.columns]
                        if len(valid_final_columns) != len(final_columns):
                             print("警告: 並べ替え列リストに不正な列が含まれていました。") # より詳細な情報が必要かも

                        result_df = result_df[valid_final_columns] # 存在する列のみで再構成
                        self._empty_col_mapping = new_empty_cols_mapping # マッピングを保存
                    except KeyError as e:
                         print(f"エラー: 列の選択中に予期せぬエラー。存在しない列: {e}")
                         # エラー時は元のDFを返す (並べ替えなし)
                         self._empty_col_mapping = {}
                         return df.copy() # 元のコピーを返す
                    except Exception as reorder_ex:
                         print(f"エラー: 列の並べ替え中に予期せぬエラー: {reorder_ex}")
                         self._empty_col_mapping = {}
                         return df.copy()
                else:
                    # 並べ替え指定が空、または有効な列が一つもなかった場合
                    if specified_columns_with_blanks: # 何か指定はあったのに有効なものがなかった
                        print("警告: 並べ替えで指定された有効な列がありません。出力は空になります。")
                    result_df = pd.DataFrame() # 空のDataFrameを返す
                    self._empty_col_mapping = {}

            return result_df

        except Exception as e:
            messagebox.showerror("エラー", f"データ処理中に予期せぬエラーが発生しました: {str(e)}")
            return df.copy() # エラー時は元のコピーを返す
    
    def update_preview(self):
        # --- update_preview は self.tree が再生成されている前提で動作 ---
        # --- 冒頭のクリア処理は不要 ---
        # for item in self.tree.get_children(): self.tree.delete(item)
        # self.tree["columns"] = ()
        # self.tree["displaycolumns"] = ()
        try:
            # --- TreeviewがNoneでないことを確認 ---
            if self.tree is None:
                 print("エラー: update_preview 呼び出し時に Treeview が存在しません。")
                 return

            # --- 既存の行があればクリア ---
            # (再生成されているはずだが、念のためクリア)
            for item in self.tree.get_children():
                try:
                    self.tree.delete(item)
                except tk.TclError: # アイテムが存在しない場合など
                    pass # 無視

            # --- 列定義もクリア (再生成されているはずだが念のため) ---
            try:
                self.tree["columns"] = ()
                self.tree["displaycolumns"] = ()
            except tk.TclError as e:
                 print(f"警告: update_previewでの列クリア中にTclError: {e}")
                 # ここでエラーが起きる場合、Treeviewの状態がまだおかしい可能性がある
                 # 再度再生成を試みるか？ -> 無限ループの可能性あり。一旦無視して進める。
                 pass
            except Exception as e:
                 print(f"警告: update_previewでの列クリア中にエラー: {e}")
                 pass


            if self.preview_df is None or self.preview_df.empty:
                # データがない場合はここで終了 (Treeviewは空の状態)
                return

            # --- 以降の処理はほぼ同じ ---
            columns = list(self.preview_df.columns)
            display_columns = []
            column_widths = {}
            column_headings = {}
            empty_col_mapping = getattr(self, '_empty_col_mapping', {})
            temp_columns_for_display = []

            for col in columns:
                temp_columns_for_display.append(col)
                column_headings[col] = "" if col in empty_col_mapping else col
                column_widths[col] = 100 # TODO: 列幅の自動調整や保存も検討
                display_columns.append(col)

            # --- Treeview設定 (try-exceptで囲む) ---
            try:
                self.tree["columns"] = temp_columns_for_display
                self.tree["displaycolumns"] = display_columns # 表示する列を指定
                self.tree["show"] = "headings"

                for col in display_columns: # display_columns を使う
                    self.tree.heading(col, text=column_headings[col], anchor=tk.W)
                    self.tree.column(col, width=column_widths[col], anchor=tk.W, stretch=tk.NO)
            except tk.TclError as e:
                 messagebox.showerror("プレビューエラー", f"Treeviewの列設定中にエラーが発生しました。\n{e}\nアプリケーションを再起動してください。")
                 self._cleanup_on_error() # エラー時はクリーンアップ
                 return
            except Exception as e:
                 messagebox.showerror("プレビューエラー", f"Treeviewの列設定中に予期せぬエラーが発生しました。\n{e}")
                 self._cleanup_on_error()
                 return


            # データの表示 (try-exceptで囲む)
            try:
                preview_rows = min(10, len(self.preview_df)) # 表示行数を制限
                for i in range(preview_rows):
                    row_data = self.preview_df.iloc[i]
                    # temp_columns_for_display を使って値を取得
                    values = [str(row_data[col]) if pd.notna(row_data[col]) else "" for col in temp_columns_for_display]
                    self.tree.insert("", tk.END, values=values)

                # 10行より多い場合は省略記号を表示
                if len(self.preview_df) > 10:
                    ellipsis_values = ["..."] * len(temp_columns_for_display)
                    self.tree.insert("", tk.END, values=ellipsis_values)
            except Exception as e:
                messagebox.showerror("エラー", f"プレビューデータの表示中にエラーが発生しました: {str(e)}")
                # データ表示エラーの場合もTreeviewをクリアしておく
                try:
                    for item in self.tree.get_children(): self.tree.delete(item)
                except: pass


        except Exception as e:
            messagebox.showerror("エラー", f"プレビューの更新中に予期せぬエラーが発生しました: {str(e)}")
            # 致命的なエラーの場合はクリーンアップ
            self._cleanup_on_error()
    
    def process_and_save(self):
        if not self.current_file:
            messagebox.showerror("エラー", "処理するCSVファイルが選択されていません")
            return
        if self.preview_df is None or self.preview_df.empty:
             messagebox.showerror("エラー", "処理対象のデータがありません。ファイルを確認してください。")
             return

        try:
            processed_df_preview = self.preview_df.copy()
            empty_col_mapping = getattr(self, '_empty_col_mapping', {})
            rename_dict = {ph: '' for ph in empty_col_mapping if ph in processed_df_preview.columns}

            processed_df_to_save = processed_df_preview.rename(columns=rename_dict) if rename_dict else processed_df_preview

            base_name = os.path.basename(self.current_file)
            name, ext = os.path.splitext(base_name)
            output_path = filedialog.asksaveasfilename(
                title="変換後のファイルを保存",
                initialfile=f"{name}_converted{ext}",
                defaultextension=".csv",
                filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
            )

            if output_path:
                selected_output_encoding = self.output_encoding.get()
                output_header = not self.remove_header_var.get()
                try:
                    processed_df_to_save.to_csv(
                        output_path,
                        index=False,
                        encoding=selected_output_encoding,
                        header=output_header,
                        quoting=csv.QUOTE_ALL # ダブルクォートで囲む場合
                        # quoting=csv.QUOTE_NONNUMERIC # 数値以外をダブルクォートで囲む場合
                        # quoting=csv.QUOTE_MINIMAL # デフォルト（特殊文字が含まれる場合のみ）
                    )
                    messagebox.showinfo("成功", f"ファイルを保存しました: {output_path}")
                except Exception as e:
                     # エンコーディングエラーの可能性も考慮
                     messagebox.showerror("保存エラー", f"ファイルの保存中にエラーが発生しました (エンコーディング: {selected_output_encoding}):\n{str(e)}")

        except Exception as e:
            messagebox.showerror("エラー", f"処理と保存中に予期せぬエラーが発生しました: {str(e)}")

# --- csvモジュールをインポート ---
import csv

if __name__ == "__main__":
    app = CSVLayoutTool()
    app.mainloop()