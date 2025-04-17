import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json
import re
from tkinter.scrolledtext import ScrolledText
from tkinterdnd2 import DND_FILES, TkinterDnD
import csv # csvモジュールをインポート

class CSVLayoutTool(TkinterDnD.Tk):
    # --- 定数定義 ---
    PROFILE_FILENAME = "csv_profiles.json"
    DEFAULT_PREF_CODE_COLUMN = "都道府県コード"
    EMPTY_COLUMN_PLACEHOLDER_PREFIX = "__EMPTY_COLUMN_"

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

        # --- ヘッダー除去関連 ---
        self.remove_header_var = tk.BooleanVar(value=False)

        # --- 都道府県削除関連 ---
        self.remove_prefecture_var = tk.BooleanVar(value=False)
        self.remove_prefecture_column_var = tk.StringVar()

        # --- 都道府県コード取得関連 ---
        self.get_pref_code_var = tk.BooleanVar(value=False)
        self.get_pref_code_source_column_var = tk.StringVar()
        self.get_pref_code_new_column_var = tk.StringVar(value=self.DEFAULT_PREF_CODE_COLUMN)

        # --- 都道府県名とコードのマッピング ---
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

        # 文字列抽出タブ
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

        # 文字置換タブ
        replace_frame = ttk.Frame(self.operations_notebook)
        self.operations_notebook.add(replace_frame, text="文字置換")

        ttk.Label(replace_frame, text="置換設定 (項目名:置換前文字列:置換後文字列)").pack(pady=(10, 5), padx=5, anchor=tk.W)
        self.replace_text = ScrolledText(replace_frame, height=8, width=40, wrap=tk.WORD)
        self.replace_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(replace_frame, text="例: ステータス:処理中:完了 / 商品名:旧製品:新製品").pack(pady=5, padx=5, anchor=tk.W)
        ttk.Label(replace_frame, text="※各設定は改行で区切ってください").pack(pady=(0,5), padx=5, anchor=tk.W)

        # 結合タブ
        merge_frame = ttk.Frame(self.operations_notebook)
        self.operations_notebook.add(merge_frame, text="結合")

        ttk.Label(merge_frame, text="結合設定 (新項目名:結合元項目1,結合元項目2...,区切り文字)").pack(pady=(10, 5), padx=5, anchor=tk.W)
        self.merge_text = ScrolledText(merge_frame, height=8, width=40, wrap=tk.WORD)
        self.merge_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        ttk.Label(merge_frame, text="例: 氏名:姓,名, / 住所:都道府県,市区町村,番地,-").pack(pady=5, padx=5, anchor=tk.W)
        ttk.Label(merge_frame, text="※最後のカンマ以降が区切り文字になります").pack(pady=(0,5), padx=5, anchor=tk.W)

        # 都道府県コード取得タブ
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

        # プレビュー領域のウィジェット
        self._create_preview_widgets(right_frame)

        # ヘッダー除去チェックボックス
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

        # 実行ボタン
        ttk.Button(right_frame, text="変換して保存", command=self.process_and_save).pack(fill=tk.X, padx=5, pady=5)

    def _create_preview_widgets(self, parent_frame):
        # プレビュー領域フレーム
        self.preview_frame = ttk.LabelFrame(parent_frame, text="プレビュー")
        self.preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # スクロールバー
        self.vsb = ttk.Scrollbar(self.preview_frame, orient="vertical")
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.hsb = ttk.Scrollbar(self.preview_frame, orient="horizontal")
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Treeview
        self.tree = ttk.Treeview(self.preview_frame, yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # スクロールバーのcommand設定
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

    def _recreate_treeview(self):
        if self.tree:
            try:
                self.tree.destroy() # 古いTreeviewを破棄
            except tk.TclError:
                pass # 既に破棄されている場合など

        # 新しいTreeviewを作成
        self.tree = ttk.Treeview(self.preview_frame, yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.tree.pack(fill=tk.BOTH, expand=True) # 再度packする

        # スクロールバーのcommandを新しいTreeviewに再設定
        self.vsb.config(command=self.tree.yview)
        self.hsb.config(command=self.tree.xview)

        # 新しいTreeviewをスクロールバーの前に表示させる
        self.vsb.lift(self.tree)
        self.hsb.lift(self.tree)

    def load_profiles(self):
        try:
            if os.path.exists(self.PROFILE_FILENAME):
                with open(self.PROFILE_FILENAME, "r", encoding="utf-8") as f:
                    self.profiles = json.load(f)

                self.profile_combobox["values"] = list(self.profiles.keys())
                if self.profiles:
                    self.current_profile_name.set(list(self.profiles.keys())[0])
                    self.load_profile(None)
        except Exception as e:
            messagebox.showerror("エラー", f"プロファイルの読み込みに失敗しました: {str(e)}")

    def save_profiles(self):
        try:
            with open(self.PROFILE_FILENAME, "w", encoding="utf-8") as f:
                json.dump(self.profiles, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("エラー", f"プロファイルの保存に失敗しました: {str(e)}")

    def new_profile(self):
        profile_name = tk.simpledialog.askstring("新規プロファイル", "プロファイル名を入力してください:")
        if profile_name and profile_name.strip():
            profile_name = profile_name.strip()
            if profile_name in self.profiles:
                messagebox.showerror("エラー", "同名のプロファイルが既に存在します")
                return

            self.profiles[profile_name] = {
                "reorder": "",
                "merge": "",
                "extract": "",
                "remove": "",
                "add": "",
                "replace": "",
                "remove_prefecture": {"enabled": False, "column": ""},
                "get_pref_code": {"enabled": False, "source_column": "", "new_column": self.DEFAULT_PREF_CODE_COLUMN},
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

        merge_lines = self.merge_text.get("1.0", tk.END).splitlines()
        while merge_lines and not merge_lines[-1].strip():
            merge_lines.pop()
        merge_setting_cleaned = "\n".join(merge_lines)
        reorder_lines = self.reorder_text.get("1.0", tk.END).splitlines()
        while reorder_lines and not reorder_lines[-1].strip(): reorder_lines.pop()
        reorder_setting_cleaned = "\n".join(reorder_lines)

        extract_lines = self.extract_text.get("1.0", tk.END).splitlines()
        while extract_lines and not extract_lines[-1].strip(): extract_lines.pop()
        extract_setting_cleaned = "\n".join(extract_lines)

        remove_lines = self.remove_text.get("1.0", tk.END).splitlines()
        while remove_lines and not remove_lines[-1].strip(): remove_lines.pop()
        remove_setting_cleaned = "\n".join(remove_lines)

        add_lines = self.add_text.get("1.0", tk.END).splitlines()
        while add_lines and not add_lines[-1].strip(): add_lines.pop()
        add_setting_cleaned = "\n".join(add_lines)

        replace_lines = self.replace_text.get("1.0", tk.END).splitlines()
        while replace_lines and not replace_lines[-1].strip(): replace_lines.pop()
        replace_setting_cleaned = "\n".join(replace_lines)

        self.profiles[profile_name] = {
            "reorder": reorder_setting_cleaned,
            "merge": merge_setting_cleaned,
            "extract": extract_setting_cleaned,
            "remove": remove_setting_cleaned,
            "add": add_setting_cleaned,
            "replace": replace_setting_cleaned,
            "remove_prefecture": {
                "enabled": self.remove_prefecture_var.get(),
                "column": self.remove_prefecture_column_var.get()
            },
            "get_pref_code": {
                "enabled": self.get_pref_code_var.get(),
                "source_column": self.get_pref_code_source_column_var.get(),
                "new_column": self.get_pref_code_new_column_var.get() or self.DEFAULT_PREF_CODE_COLUMN # 空の場合デフォルト値
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
                self.replace_text.delete("1.0", tk.END)
                self.remove_prefecture_var.set(False)
                self.remove_prefecture_column_var.set("")
                self.get_pref_code_var.set(False)
                self.get_pref_code_source_column_var.set("")
                self.get_pref_code_new_column_var.set(self.DEFAULT_PREF_CODE_COLUMN)
                self.remove_header_var.set(False)

            self.save_profiles()

    def load_profile(self, event):
        profile_name = self.current_profile_name.get()
        if not profile_name or profile_name not in self.profiles:
            # プロファイルが存在しない場合はフィールドをクリア
            self.current_profile_name.set("")
            self.reorder_text.delete("1.0", tk.END)
            self.merge_text.delete("1.0", tk.END)
            self.extract_text.delete("1.0", tk.END)
            self.remove_text.delete("1.0", tk.END)
            self.add_text.delete("1.0", tk.END)
            self.replace_text.delete("1.0", tk.END)
            self.remove_prefecture_var.set(False)
            self.remove_prefecture_column_var.set("")
            self.get_pref_code_var.set(False)
            self.get_pref_code_source_column_var.set("")
            self.get_pref_code_new_column_var.set(self.DEFAULT_PREF_CODE_COLUMN)
            self.remove_header_var.set(False)
            return

        profile = self.profiles[profile_name]

        # 設定読み込み (getのデフォルト値を空文字列に)
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
        self.replace_text.delete("1.0", tk.END)
        self.replace_text.insert(tk.END, profile.get("replace", ""))

        # 都道府県削除設定
        remove_pref_settings = profile.get("remove_prefecture", {"enabled": False, "column": ""})
        self.remove_prefecture_var.set(remove_pref_settings.get("enabled", False))
        self.remove_prefecture_column_var.set(remove_pref_settings.get("column", ""))

        # 都道府県コード設定
        get_pref_code_settings = profile.get("get_pref_code", {"enabled": False, "source_column": "", "new_column": self.DEFAULT_PREF_CODE_COLUMN})
        self.get_pref_code_var.set(get_pref_code_settings.get("enabled", False))
        self.get_pref_code_source_column_var.set(get_pref_code_settings.get("source_column", ""))
        self.get_pref_code_new_column_var.set(get_pref_code_settings.get("new_column", self.DEFAULT_PREF_CODE_COLUMN))

        # ヘッダー除去設定
        self.remove_header_var.set(profile.get("remove_header", False))

        # 現在のファイルがあれば再プレビュー
        if self.current_file:
            # プロファイル変更時もTreeview再生成とプレビューを行う
            self.after_idle(lambda: self._clear_and_preview_logic(self.current_file))

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            self.after_idle(lambda: self._clear_and_preview_logic(file_path))

    def drop(self, event):
        try:
            # ドラッグされたデータの解析 (より堅牢に)
            raw_path = event.data.strip()
            if raw_path.startswith('{') and raw_path.endswith('}'):
                raw_path = raw_path[1:-1]
            if raw_path.startswith('"') and raw_path.endswith('"'):
                raw_path = raw_path[1:-1]
            # 複数のファイルがドロップされた場合も考慮 (最初のファイルのみ処理)
            file_paths = raw_path.split() # スペース区切りを仮定
            if not file_paths:
                messagebox.showerror("エラー", "有効なファイルパスを取得できませんでした。")
                return
            file_path = file_paths[0] # 最初のファイルパスを使用

            if not os.path.isfile(file_path):
                messagebox.showerror("エラー", f"有効なファイルではありません:\n{file_path}")
                return
            if not file_path.lower().endswith('.csv'):
                messagebox.showerror("エラー", "CSVファイルのみ対応しています")
                return

            self.after_idle(lambda: self._clear_and_preview_logic(file_path))

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルのドロップ処理中にエラーが発生しました: {str(e)}")
            self.after_idle(self._cleanup_on_error)

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
             # ここでのエラーはコンソールに出力する程度に留める
             print(f"警告: エラー後のTreeview再生成中にエラー: {cleanup_error}")

    def preview_file(self, file_path):
        try:
            selected_encoding = self.encoding.get()
            df = None
            try:
                # keep_default_na=False で空文字列を NaN にしない
                # dtype=str ですべての列を文字列として読み込む
                df = pd.read_csv(file_path, encoding=selected_encoding, dtype=str, keep_default_na=False)
            except UnicodeDecodeError:
                alternative_encoding = "shift_jis" if selected_encoding == "utf-8" else "utf-8"
                try:
                    df = pd.read_csv(file_path, encoding=alternative_encoding, dtype=str, keep_default_na=False)
                    self.encoding.set(alternative_encoding)
                    messagebox.showinfo("エンコーディング変更",
                                        f"選択されたエンコーディング({selected_encoding})では読み込めませんでした。\n"
                                        f"代わりに{alternative_encoding}で読み込みました。")
                except UnicodeDecodeError:
                    messagebox.showerror("エラー", "UTF-8とShift-JISのどちらでもファイルを読み込めませんでした。\n"
                                                "ファイルの文字コードを確認してください。")
                    self._cleanup_on_error()
                    return
                except FileNotFoundError:
                     messagebox.showerror("エラー", f"ファイルが見つかりません:\n{file_path}")
                     self._cleanup_on_error()
                     return
                except pd.errors.EmptyDataError:
                     messagebox.showerror("エラー", f"ファイルが空か、CSVデータが含まれていません:\n{file_path}")
                     self._cleanup_on_error()
                     return
                except Exception as inner_e:
                     messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました ({alternative_encoding}):\n{str(inner_e)}")
                     self._cleanup_on_error()
                     return
            except FileNotFoundError:
                 messagebox.showerror("エラー", f"ファイルが見つかりません:\n{file_path}")
                 self._cleanup_on_error()
                 return
            except pd.errors.EmptyDataError:
                 messagebox.showerror("エラー", f"ファイルが空か、CSVデータが含まれていません:\n{file_path}")
                 self._cleanup_on_error()
                 return
            except Exception as outer_e:
                 messagebox.showerror("エラー", f"ファイルの読み込み中にエラーが発生しました ({selected_encoding}):\n{str(outer_e)}")
                 self._cleanup_on_error()
                 return

            # dfがNoneでないことを確認
            if df is None:
                messagebox.showerror("エラー", "データフレームの読み込みに失敗しました。")
                self._cleanup_on_error()
                return

            self.current_file = file_path
            self.title(f"CSVレイアウト変更ツール - {os.path.basename(file_path)}")

            # プレビューの作成
            self.preview_df = self.process_dataframe(df)
            self.update_preview()

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルプレビュー処理中に予期せぬエラーが発生しました: {str(e)}")
            self._cleanup_on_error()

    def process_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        DataFrameに対して定義された処理を実行する。
        エラーが発生した場合は元のDataFrameのコピーを返す。
        """
        try:
            result_df = df.copy()
            warnings = [] # 処理中の警告を収集するリスト

            result_df, warnings = self._process_get_pref_code(result_df, warnings)
            result_df, warnings = self._process_remove_prefecture(result_df, warnings)
            result_df, warnings = self._process_extract(result_df, warnings)
            result_df, warnings = self._process_remove(result_df, warnings)
            result_df, warnings = self._process_add(result_df, warnings)
            result_df, warnings = self._process_replace(result_df, warnings)
            result_df, warnings = self._process_merge(result_df, warnings)
            result_df, warnings = self._process_reorder(result_df, warnings)

            # 警告があればコンソールに出力 (必要に応じてUI表示に変更)
            if warnings:
                print("-" * 20 + " 処理中の警告 " + "-" * 20)
                for warn in warnings:
                    print(warn)
                print("-" * 55)

            return result_df

        except Exception as e:
            messagebox.showerror("データ処理エラー", f"データ処理中に予期せぬエラーが発生しました:\n{str(e)}")
            return df.copy() # エラー時は元のコピーを返す

    def _process_get_pref_code(self, df: pd.DataFrame, warnings: list) -> (pd.DataFrame, list):
        if self.get_pref_code_var.get():
            source_col = self.get_pref_code_source_column_var.get().strip()
            new_col = self.get_pref_code_new_column_var.get().strip() or self.DEFAULT_PREF_CODE_COLUMN

            if not source_col:
                warnings.append("都道府県コード取得: 都道府県名を含む項目名が指定されていません。")
                return df, warnings
            if not new_col:
                warnings.append("都道府県コード取得: 新しい項目名が指定されていません。")
                return df, warnings

            if source_col not in df.columns:
                warnings.append(f"都道府県コード取得: 対象項目 '{source_col}' が見つかりません。")
                return df, warnings
            if new_col in df.columns:
                warnings.append(f"都道府県コード取得: 新しい項目名 '{new_col}' は既に存在します。処理をスキップします。")
                return df, warnings

            def get_code(address):
                if isinstance(address, str):
                    for pref, code in self.PREFECTURE_CODES.items():
                        if address.startswith(pref):
                            return code
                return "" # 見つからない場合は空文字

            try:
                df[new_col] = df[source_col].apply(get_code)
            except Exception as e:
                 warnings.append(f"都道府県コード取得処理中にエラー: {e}")

        return df, warnings

    def _process_remove_prefecture(self, df: pd.DataFrame, warnings: list) -> (pd.DataFrame, list):
        if self.remove_prefecture_var.get():
            target_columns_str = self.remove_prefecture_column_var.get().strip()
            if not target_columns_str:
                warnings.append("都道府県名削除: 対象項目名が指定されていません。")
                return df, warnings

            target_columns = [col.strip() for col in target_columns_str.split(',') if col.strip()]
            valid_target_columns = []
            for target_column in target_columns:
                if target_column in df.columns:
                    valid_target_columns.append(target_column)
                else:
                    warnings.append(f"都道府県名削除: 対象項目 '{target_column}' が見つかりません。")

            if not valid_target_columns:
                return df, warnings # 有効な対象列がない

            def remove_pref(address):
                if isinstance(address, str):
                    for pref in self.PREFECTURES:
                        if address.startswith(pref):
                            return address[len(pref):]
                return address

            try:
                for target_column in valid_target_columns:
                    df[target_column] = df[target_column].apply(remove_pref)
            except Exception as e:
                warnings.append(f"都道府県名削除処理中にエラー: {e}")

        return df, warnings

    def _process_merge(self, df: pd.DataFrame, warnings: list) -> (pd.DataFrame, list):
        # merge_text から設定を取得し、末尾の不要な空行を除去
        merge_lines = self.merge_text.get("1.0", tk.END).splitlines()
        while merge_lines and not merge_lines[-1].strip():
            merge_lines.pop()
        merge_settings_content = "\n".join(merge_lines)

        if merge_settings_content:
            # 各行を処理 (enumerate は 1 始まり)
            for line_num, line in enumerate(merge_settings_content.split('\n'), 1):
                # 先に strip せずに、まず空行かどうかだけチェック
                if not line.strip(): continue # 実質的に空の行はスキップ

                try:
                    # コロンでの分割は元の line に対して行う
                    parts = line.split(':', 1)
                    if len(parts) != 2:
                        warnings.append(f"結合設定(行 {line_num}): 形式が不正です (':')。スキップします: {line.strip()}")
                        continue

                    # 新項目名部分のみ strip する
                    new_column = parts[0].strip()
                    # merge_info は strip せず、元の文字列（末尾スペースを含む可能性あり）を保持
                    merge_info = parts[1]

                    if not new_column:
                        warnings.append(f"結合設定(行 {line_num}): 新項目名が空です。スキップします: {line.strip()}")
                        continue

                    # --- 区切り文字抽出 ---
                    last_comma_index = merge_info.rfind(',')
                    if last_comma_index == -1:
                        # カンマがない場合: 結合元は1つ、区切り文字なしと解釈
                        source_columns_str = merge_info.strip()
                        separator = ''
                        source_columns = [source_columns_str] if source_columns_str else []
                        if not source_columns:
                             warnings.append(f"結合設定(行 {line_num}): 結合元項目が指定されていません。スキップします: {line.strip()}")
                             continue
                    else:
                        # 最後のカンマより前が結合元項目リスト、後ろが区切り文字
                        source_columns_str = merge_info[:last_comma_index].strip()
                        separator = merge_info[last_comma_index + 1:] # 末尾のスペース等も区切り文字の一部として保持
                        source_columns = [col.strip() for col in source_columns_str.split(',') if col.strip()]
                        if not source_columns:
                            warnings.append(f"結合設定(行 {line_num}): 結合元項目が指定されていません。スキップします: {line.strip()}")
                            continue

                    # 結合元項目がDataFrameに存在するかチェック
                    missing_cols = [col for col in source_columns if col not in df.columns]
                    if missing_cols:
                        warnings.append(f"結合設定(行 {line_num}): 結合元項目が見つかりません: {', '.join(missing_cols)}。スキップします: {line.strip()}")
                        continue

                    # 新しい項目名が既に存在する場合の警告
                    if new_column in df.columns:
                        warnings.append(f"結合設定(行 {line_num}): 結合先の項目名 '{new_column}' は既に存在します。上書きします。")

                    # 結合実行 (NaNを空文字に変換)
                    try:
                        # apply内のlambda関数
                        def join_items(row_items, sep):
                            # 各要素を文字列に変換（NaNは空文字に）
                            items = [str(item) if pd.notna(item) else '' for item in row_items]
                            # 区切り文字で結合
                            return sep.join(items)

                        # DataFrameに関数を適用
                        df[new_column] = df[source_columns].apply(lambda x: join_items(x, separator), axis=1)

                    except Exception as apply_ex:
                         warnings.append(f"結合設定(行 {line_num}): apply処理中にエラー: {apply_ex}。スキップ: {line.strip()}")
                         continue # この行の処理をスキップ

                except Exception as merge_ex:
                    warnings.append(f"結合設定(行 {line_num}): 処理中にエラーが発生しました: {merge_ex}。スキップします: {line.strip()}")
        return df, warnings

    def _process_extract(self, df: pd.DataFrame, warnings: list) -> (pd.DataFrame, list):
        extract_settings = self.extract_text.get("1.0", tk.END).strip()
        if extract_settings:
            for line_num, line in enumerate(extract_settings.split('\n'), 1):
                line = line.strip()
                if not line: continue
                try:
                    parts = line.split(':', 3)
                    if len(parts) != 4:
                        warnings.append(f"文字列抽出(行 {line_num}): 形式が不正 ('新:元:開始:文字数')。スキップします: {line}")
                        continue

                    new_col, source_col, start_pos_str, num_chars_str = [p.strip() for p in parts]

                    if not new_col: warnings.append(f"文字列抽出(行 {line_num}): 新項目名が空。スキップ: {line}"); continue
                    if not source_col: warnings.append(f"文字列抽出(行 {line_num}): 抽出元項目が空。スキップ: {line}"); continue
                    if source_col not in df.columns: warnings.append(f"文字列抽出(行 {line_num}): 抽出元項目 '{source_col}' が見つかりません。スキップ。"); continue
                    if new_col in df.columns: warnings.append(f"文字列抽出(行 {line_num}): 新項目名 '{new_col}' は既に存在。上書きします。")

                    try:
                        start_pos = int(start_pos_str)
                        num_chars = int(num_chars_str)
                        if start_pos < 1: warnings.append(f"文字列抽出(行 {line_num}): 開始位置は1以上。スキップ: {line}"); continue
                        if num_chars < 0: warnings.append(f"文字列抽出(行 {line_num}): 文字数は0以上。スキップ: {line}"); continue
                    except ValueError: warnings.append(f"文字列抽出(行 {line_num}): 開始位置/文字数が数値でない。スキップ: {line}"); continue

                    start_index = start_pos - 1 # 0-based index
                    end_index = start_index + num_chars

                    def extract_substring(text):
                        if pd.isna(text): return ""
                        text_str = str(text)
                        if start_index >= len(text_str): return "" # 開始位置が文字列長以上
                        # 抽出範囲が文字列を超える場合、最後まで抽出
                        actual_end_index = min(end_index, len(text_str))
                        return text_str[start_index:actual_end_index]

                    df[new_col] = df[source_col].apply(extract_substring)

                except Exception as extract_ex:
                    warnings.append(f"文字列抽出(行 {line_num}): 処理中にエラー: {extract_ex}。スキップ: {line}")
        return df, warnings

    def _process_remove(self, df: pd.DataFrame, warnings: list) -> (pd.DataFrame, list):
        remove_settings = self.remove_text.get("1.0", tk.END).strip()
        if remove_settings:
            for line_num, line in enumerate(remove_settings.split('\n'), 1):
                line = line.strip()
                if not line: continue
                try:
                    parts = line.split(':', 1)
                    if len(parts) != 2:
                        warnings.append(f"文字除去(行 {line_num}): 形式が不正 (':')。スキップ: {line}"); continue

                    column = parts[0].strip()
                    chars_to_remove_str = parts[1].strip() # 除去する文字列表記 (カンマ区切り)

                    if not column: warnings.append(f"文字除去(行 {line_num}): 項目名が空。スキップ: {line}"); continue
                    if not chars_to_remove_str: warnings.append(f"文字除去(行 {line_num}): 除去文字が空。スキップ: {line}"); continue
                    if column not in df.columns: warnings.append(f"文字除去(行 {line_num}): 項目 '{column}' が見つかりません。スキップ。"); continue

                    # カンマ区切りで除去文字リストを作成
                    chars_to_remove_list = [c.strip() for c in chars_to_remove_str.split(',') if c.strip()]
                    if not chars_to_remove_list: warnings.append(f"文字除去(行 {line_num}): 有効な除去文字がありません。スキップ: {line}"); continue

                    # 各除去文字に対してreplaceを実行
                    temp_series = df[column].astype(str) # 文字列に変換
                    for char in chars_to_remove_list:
                        # regex=False でリテラル文字列として置換
                        temp_series = temp_series.str.replace(char, '', regex=False)
                    df[column] = temp_series

                except Exception as remove_ex:
                    warnings.append(f"文字除去(行 {line_num}): 処理中にエラー: {remove_ex}。スキップ: {line}")
        return df, warnings

    def _process_add(self, df: pd.DataFrame, warnings: list) -> (pd.DataFrame, list):
        add_settings = self.add_text.get("1.0", tk.END).strip()
        if add_settings:
            for line_num, line in enumerate(add_settings.split('\n'), 1):
                line = line.strip()
                if not line: continue
                try:
                    parts = line.split(':', 2)
                    if len(parts) != 3:
                        warnings.append(f"文字追加(行 {line_num}): 形式が不正 ('項目:位置:追加文字')。スキップ: {line}"); continue

                    column, position, chars_to_add = [p.strip() for p in parts]

                    if not column: warnings.append(f"文字追加(行 {line_num}): 項目名が空。スキップ: {line}"); continue
                    if position not in ["前", "後"]: warnings.append(f"文字追加(行 {line_num}): 位置は '前' または '後'。スキップ: {line}"); continue
                    if column not in df.columns: warnings.append(f"文字追加(行 {line_num}): 項目 '{column}' が見つかりません。スキップ。"); continue

                    # fillna('') を使って NaN を空文字列に変換してから追加
                    col_series = df[column].fillna('').astype(str)
                    if position == "前":
                        df[column] = chars_to_add + col_series
                    elif position == "後":
                        df[column] = col_series + chars_to_add

                except Exception as add_ex:
                    warnings.append(f"文字追加(行 {line_num}): 処理中にエラー: {add_ex}。スキップ: {line}")
        return df, warnings

    def _process_replace(self, df: pd.DataFrame, warnings: list) -> (pd.DataFrame, list):
        replace_settings = self.replace_text.get("1.0", tk.END).strip()
        if replace_settings:
            for line_num, line in enumerate(replace_settings.split('\n'), 1):
                line = line.strip()
                if not line: continue
                try:
                    parts = line.split(':', 2)
                    if len(parts) != 3:
                        warnings.append(f"文字置換(行 {line_num}): 形式が不正 ('項目:置換前:置換後')。スキップ: {line}"); continue

                    column, old_str, new_str = [p.strip() for p in parts]

                    if not column: warnings.append(f"文字置換(行 {line_num}): 項目名が空。スキップ: {line}"); continue
                    if not old_str: warnings.append(f"文字置換(行 {line_num}): 置換前文字列が空。スキップ: {line}"); continue
                    if column not in df.columns: warnings.append(f"文字置換(行 {line_num}): 項目 '{column}' が見つかりません。スキップ。"); continue

                    # fillna('') でNaNを空文字に変換し、astype(str) で文字列型に統一
                    df[column] = df[column].fillna('').astype(str)
                    # .loc を使用して、old_str と完全に一致する値のみを new_str に置換
                    df.loc[df[column] == old_str, column] = new_str

                except Exception as replace_ex:
                    warnings.append(f"文字置換(行 {line_num}): 処理中にエラー: {replace_ex}。スキップ: {line}")
        return df, warnings

    def _process_reorder(self, df: pd.DataFrame, warnings: list) -> (pd.DataFrame, list):
        reorder_settings = self.reorder_text.get("1.0", tk.END).strip()
        self._empty_col_mapping = {} # 並べ替え前にクリア

        if reorder_settings:
            specified_columns_with_blanks = [col.strip() for col in reorder_settings.split(',')]
            final_columns = []
            new_empty_cols_mapping = {}
            empty_col_counter = 1
            current_columns = list(df.columns) # 現在のDFの列リスト

            for col_name in specified_columns_with_blanks:
                if col_name: # 通常の列名
                    if col_name in current_columns:
                        final_columns.append(col_name)
                    else:
                        warnings.append(f"並べ替え: 指定された列 '{col_name}' はデータに存在しません。無視します。")
                else: # 空列を追加 (,, の場合)
                    # 一意なプレースホルダー名を生成
                    placeholder_name = f"{self.EMPTY_COLUMN_PLACEHOLDER_PREFIX}{empty_col_counter}"
                    while placeholder_name in current_columns or placeholder_name in new_empty_cols_mapping:
                        empty_col_counter += 1
                        placeholder_name = f"{self.EMPTY_COLUMN_PLACEHOLDER_PREFIX}{empty_col_counter}"

                    df[placeholder_name] = "" # 空列をDataFrameに追加
                    final_columns.append(placeholder_name)
                    new_empty_cols_mapping[placeholder_name] = '' # マッピングに追加
                    current_columns.append(placeholder_name) # 後続の重複チェックのため
                    empty_col_counter += 1

            if final_columns:
                try:
                    # 存在する列のみでDataFrameを再構成
                    df = df[final_columns]
                    self._empty_col_mapping = new_empty_cols_mapping # マッピングを保存
                except KeyError as e:
                    warnings.append(f"並べ替え: 列の選択中にエラー。存在しない列: {e}。並べ替えは適用されません。")
                    self._empty_col_mapping = {}
                except Exception as reorder_ex:
                    warnings.append(f"並べ替え: 処理中に予期せぬエラー: {reorder_ex}。並べ替えは適用されません。")
                    self._empty_col_mapping = {}
            else:
                # 並べ替え指定が空、または有効な列が一つもなかった場合
                if specified_columns_with_blanks: # 何か指定はあったが無効だった
                    warnings.append("並べ替え: 指定された有効な列がありません。出力は空になります。")
                df = pd.DataFrame() # 空のDataFrameを返す
                self._empty_col_mapping = {}

        return df, warnings


    def update_preview(self):
        try:
            if self.tree is None:
                 # Treeviewがない場合は処理を中断
                 return

            # 既存の行と列をクリア
            try:
                for item in self.tree.get_children():
                    self.tree.delete(item)
                self.tree["columns"] = ()
                self.tree["displaycolumns"] = ()
            except tk.TclError:
                 # Treeviewの状態がおかしい可能性 -> 再生成を試みる
                 try:
                     self._recreate_treeview()
                 except Exception:
                     messagebox.showerror("プレビューエラー", "プレビュー表示の更新に失敗しました。")
                     return # これ以上進めない
                 # 再生成後、再度クリアを試みる (ここでは省略)
                 pass
            except Exception:
                 # その他のクリアエラーは無視して続行を試みる
                 pass

            if self.preview_df is None or self.preview_df.empty:
                # データがない場合はここで終了 (Treeviewは空の状態)
                return

            # 列情報の設定
            columns = list(self.preview_df.columns)
            # --- 修正ここから ---
            # display_columns = [] # 以前はここで初期化していた
            column_widths = {} # 将来的な幅調整用
            column_headings = {}
            empty_col_mapping = getattr(self, '_empty_col_mapping', {})

            for col in columns:
                is_empty_placeholder = col in empty_col_mapping
                column_headings[col] = "" if is_empty_placeholder else col # ヘッダーは空にする
                column_widths[col] = 100 # デフォルト幅
                # display_columns.append(col) # ここで追加するのではなく、下の displaycolumns = columns を使う
            # --- 修正ここまで ---

            # Treeview設定
            try:
                self.tree["columns"] = columns # 内部的な列リスト
                self.tree["displaycolumns"] = columns
                self.tree["show"] = "headings" # ヘッダーのみ表示

                for col in columns: # すべての内部列に対して設定
                    self.tree.heading(col, text=column_headings[col], anchor=tk.W)
                    self.tree.column(col, width=column_widths[col], anchor=tk.W, stretch=tk.NO)
            except tk.TclError as e:
                 messagebox.showerror("プレビューエラー", f"Treeviewの列設定中にエラーが発生しました。\n{e}\nアプリケーションを再起動してください。")
                 self._cleanup_on_error()
                 return
            except Exception as e:
                 messagebox.showerror("プレビューエラー", f"Treeviewの列設定中に予期せぬエラーが発生しました。\n{e}")
                 self._cleanup_on_error()
                 return

            # データ表示
            try:
                preview_rows = min(10, len(self.preview_df))
                for i in range(preview_rows):
                    row_data = self.preview_df.iloc[i]
                    values = [str(row_data[col]) if pd.notna(row_data[col]) else "" for col in columns]
                    self.tree.insert("", tk.END, values=values)

                if len(self.preview_df) > 10:
                    ellipsis_values = ["..."] * len(columns)
                    self.tree.insert("", tk.END, values=ellipsis_values)
            except Exception as e:
                messagebox.showerror("エラー", f"プレビューデータの表示中にエラーが発生しました: {str(e)}")
                # データ表示エラーの場合もクリア
                try:
                    for item in self.tree.get_children(): self.tree.delete(item)
                except: pass

        except Exception as e:
            messagebox.showerror("エラー", f"プレビューの更新中に予期せぬエラーが発生しました: {str(e)}")
            self._cleanup_on_error()

    def process_and_save(self):
        if not self.current_file:
            messagebox.showerror("エラー", "処理するCSVファイルが選択されていません")
            return
        if self.preview_df is None or self.preview_df.empty:
             if not self.reorder_text.get("1.0", tk.END).strip():
                 messagebox.showerror("エラー", "処理対象のデータがありません。ファイルを確認してください。")
                 return
             else:
                 if not messagebox.askyesno("確認", "処理結果が空ですが、空のファイルとして保存しますか？"):
                     return
                 processed_df_to_save = pd.DataFrame()
        else:
            processed_df_preview = self.preview_df.copy()
            empty_col_mapping = getattr(self, '_empty_col_mapping', {})
            rename_dict = {ph: '' for ph in empty_col_mapping if ph in processed_df_preview.columns}
            processed_df_to_save = processed_df_preview.rename(columns=rename_dict) if rename_dict else processed_df_preview


        try:
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
                        quoting=csv.QUOTE_ALL
                    )
                    messagebox.showinfo("成功", f"ファイルを保存しました:\n{output_path}")
                except Exception as e:
                     messagebox.showerror("保存エラー", f"ファイルの保存中にエラーが発生しました (エンコーディング: {selected_output_encoding}):\n{str(e)}")

        except Exception as e:
            messagebox.showerror("エラー", f"処理と保存中に予期せぬエラーが発生しました: {str(e)}")


if __name__ == "__main__":
    app = CSVLayoutTool()
    app.mainloop()
