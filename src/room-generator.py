import tkinter as tk
from tkinter import ttk
import pyperclip

class RoomGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("RoomGenerator - モード選択")
        self.root.geometry("300x200")  # 高さを増やした
        
        # Enterキーで現在フォーカスのあるボタンを実行
        self.root.bind("<Return>", lambda event: self.activate_focused_widget())

        self.setup_mode_selection()

    # 新しいメソッドを追加
    def activate_focused_widget(self):
        """現在フォーカスのあるウィジェットをアクティブにする"""
        focused = self.root.focus_get()
        if focused and 'invoke' in focused.__dir__():
            focused.invoke()

    def setup_mode_selection(self):
        """モード選択画面を表示"""
        # 画面中央にラベルを配置
        tk.Label(
            self.root, 
            text="部屋番号生成ツール", 
            font=('', 12, 'bold')
        ).pack(pady=10)
        
        # モード選択ボタン
        tk.Button(
            self.root, 
            text="通常", 
            command=self.show_normal_mode,
            width=20,
            height=2
        ).pack(pady=5)
        
        tk.Button(
            self.root, 
            text="連棟部屋作成", 
            command=self.show_extended_mode,
            width=20,
            height=2
        ).pack(pady=5)
        
        # アルファベットモードボタン追加
        tk.Button(
            self.root, 
            text="アルファベットで数える", 
            command=self.show_alphabet_mode,
            width=20,
            height=2
        ).pack(pady=5)

    def show_normal_mode(self):
        """通常モードのGUIを表示"""
        # 現在のウィンドウを閉じる
        self.root.destroy()
        
        self.root = tk.Tk()
        self.root.title("RoomGenerator - 通常")
        self.root.geometry("350x400")

        # Enterキーで現在フォーカスのあるボタンを実行
        self.root.bind("<Return>", lambda event: self.activate_focused_widget())

        # フォントサイズ設定
        font = ('', 9)
        
        # 明示的にウィンドウにフォーカスを設定
        self.root.focus_force()
        
        # 部屋数
        tk.Label(self.root, text="部屋数", font=font).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        room_number = tk.StringVar(value="2")
        room_entry = tk.Entry(self.root, textvariable=room_number, width=10, font=font)
        room_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # 階数
        tk.Label(self.root, text="階数", font=font).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        floor_number = tk.StringVar(value="2")
        floor_entry = tk.Entry(self.root, textvariable=floor_number, width=10, font=font)
        floor_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # チェックボックス
        include_4 = tk.BooleanVar()
        include_9 = tk.BooleanVar()
        cb1 = tk.Checkbutton(self.root, text="4の部屋を入れる", variable=include_4, font=font)
        cb1.var = include_4
        cb1.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        cb2 = tk.Checkbutton(self.root, text="9の部屋を入れる", variable=include_9, font=font)
        cb2.var = include_9
        cb2.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        # ボタン
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=2, pady=5)
        
        generate_btn = tk.Button(
            button_frame, 
            text="出力＆クリップボードに追加", 
            font=font,
            command=lambda: self.generate_normal(room_number.get(), floor_number.get(), include_4.get(), include_9.get())
        )
        generate_btn.grid(row=0, column=0, padx=5)
        
        close_btn = tk.Button(button_frame, text="閉じる", font=font, command=self.root.destroy)
        close_btn.grid(row=0, column=1, padx=5)
        
        back_btn = tk.Button(
            button_frame, 
            text="戻る", 
            font=font, 
            command=self.back_to_selection
        )
        back_btn.grid(row=0, column=2, padx=5)
        
        # 出力テキストエリア
        self.output_text = tk.Text(self.root, width=35, height=15, font=font)
        self.output_text.grid(row=5, column=0, columnspan=2, padx=5, pady=5)
        
        # スクロールバーの追加
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.output_text.yview)
        scrollbar.grid(row=5, column=2, sticky="ns")
        self.output_text.configure(yscrollcommand=scrollbar.set)
        
        # Ctrl+Enterのキーバインド
        self.root.bind('<Control-Return>', lambda event: self.copy_and_close())
        
        # Enter キーでボタンを押す
        self.root.bind('<Return>', lambda event: generate_btn.invoke())
        
        # Ctrl+Enterの説明ラベル
        info_label = tk.Label(
            self.root, 
            text="Ctrl + Enter で出力し、画面を閉じる", 
            font=('', 8),
            fg="gray50"
        )
        info_label.grid(row=6, column=0, columnspan=2, pady=(0, 5))
        
        self.root.mainloop()

    def show_extended_mode(self):
        """拡張モードのGUIを表示"""
        # 現在のウィンドウを閉じる
        self.root.destroy()
        
        self.root = tk.Tk()
        self.root.title("RoomGenerator - 連棟部屋作成")
        self.root.geometry("700x600")
        
        # Enterキーで現在フォーカスのあるボタンを実行
        self.root.bind("<Return>", lambda event: self.activate_focused_widget())
            
        # フォントサイズ設定
        font = ('', 9)

        # 明示的にウィンドウにフォーカスを設定
        self.root.focus_force()
        
        # メインフレーム（スクロール可能）
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=1)
        
        # キャンバスとスクロールバー
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # スクロール可能なフレーム
        scrollable_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        scrollable_frame.bind("<Configure>", configure_scroll_region)
        
        # 各棟の情報入力エリア
        building_info = []
        
        for i in range(6):
            frame = tk.LabelFrame(scrollable_frame, text=f"{i+1}棟目", padx=5, pady=5)
            frame.grid(row=i, column=0, padx=10, pady=5, sticky="ew")
            
            # 棟名
            tk.Label(frame, text="棟名", font=font).grid(row=0, column=0, padx=5, pady=2, sticky="w")
            bldg_name = tk.StringVar(value="A" if i == 0 else "B" if i == 1 else "")
            bldg_entry = tk.Entry(frame, textvariable=bldg_name, width=10, font=font)
            bldg_entry.grid(row=0, column=1, padx=5, pady=2, sticky="w")
            
            # 部屋数
            tk.Label(frame, text="部屋数", font=font).grid(row=0, column=2, padx=5, pady=2, sticky="w")
            room_number = tk.StringVar(value="2" if i < 2 else "")
            room_entry = tk.Entry(frame, textvariable=room_number, width=10, font=font)
            room_entry.grid(row=0, column=3, padx=5, pady=2, sticky="w")
            
            # 階数
            tk.Label(frame, text="階数", font=font).grid(row=0, column=4, padx=5, pady=2, sticky="w")
            floor_number = tk.StringVar(value="2" if i < 2 else "")
            floor_entry = tk.Entry(frame, textvariable=floor_number, width=10, font=font)
            floor_entry.grid(row=0, column=5, padx=5, pady=2, sticky="w")
            
            building_info.append((bldg_name, room_number, floor_number))
        
        # チェックボックス
        check_frame = tk.Frame(scrollable_frame)
        check_frame.grid(row=6, column=0, padx=10, pady=5, sticky="w")
        
        include_4 = tk.BooleanVar()
        cb = tk.Checkbutton(check_frame, text="4の部屋を入れる", variable=include_4, font=font)
        cb.var = include_4
        cb.pack(anchor="w")
        
        # ボタン
        button_frame = tk.Frame(scrollable_frame)
        button_frame.grid(row=7, column=0, padx=10, pady=5)
        
        generate_btn = tk.Button(
            button_frame, 
            text="出力＆クリップボードに追加", 
            font=font,
            command=lambda: self.generate_extended(building_info, include_4.get())
        )
        generate_btn.grid(row=0, column=0, padx=5)
        
        close_btn = tk.Button(button_frame, text="閉じる", font=font, command=self.root.destroy)
        close_btn.grid(row=0, column=1, padx=5)
        
        back_btn = tk.Button(
            button_frame, 
            text="戻る", 
            font=font, 
            command=self.back_to_selection
        )
        back_btn.grid(row=0, column=2, padx=5)
        
        # 出力テキストエリア
        output_frame = tk.LabelFrame(scrollable_frame, text="生成結果", padx=5, pady=5)
        output_frame.grid(row=8, column=0, padx=10, pady=5, sticky="ew")
        
        self.output_text = tk.Text(output_frame, width=80, height=15, font=font)
        self.output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 出力用スクロールバー
        output_scrollbar = ttk.Scrollbar(output_frame, orient="vertical", command=self.output_text.yview)
        output_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.output_text.configure(yscrollcommand=output_scrollbar.set)
        
        # Ctrl+Enterのキーバインド
        self.root.bind('<Control-Return>', lambda event: self.copy_and_close())
        
        # Enter キーでボタンを押す
        self.root.bind('<Return>', lambda event: generate_btn.invoke())
        
        # Ctrl+Enterの説明ラベル
        info_label = tk.Label(
            scrollable_frame, 
            text="Ctrl + Enter で出力し、画面を閉じる", 
            font=('', 8),
            fg="gray50"
        )
        info_label.grid(row=9, column=0, pady=(0, 5))
        
        self.root.mainloop()

    def show_alphabet_mode(self):
        """アルファベットモードのGUIを表示"""
        # 現在のウィンドウを閉じる
        self.root.destroy()
        
        self.root = tk.Tk()
        self.root.title("RoomGenerator - アルファベット")
        self.root.geometry("350x400")
        
        # Enterキーで現在フォーカスのあるボタンを実行
        self.root.bind("<Return>", lambda event: self.activate_focused_widget())

        # フォントサイズ設定
        font = ('', 9)
        
        # 明示的にウィンドウにフォーカスを設定
        self.root.focus_force()
        
        # 部屋数
        tk.Label(self.root, text="部屋数", font=font).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        room_number = tk.StringVar(value="3")
        room_entry = tk.Entry(self.root, textvariable=room_number, width=10, font=font)
        room_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # 階数
        tk.Label(self.root, text="階数", font=font).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        floor_number = tk.StringVar(value="2")
        floor_entry = tk.Entry(self.root, textvariable=floor_number, width=10, font=font)
        floor_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # チェックボックス
        add_floor = tk.BooleanVar(value=True)
        cb = tk.Checkbutton(self.root, text="階を追加する", variable=add_floor, font=font)
        cb.var = add_floor
        cb.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        # ボタン
        button_frame = tk.Frame(self.root)
        button_frame.grid(row=4, column=0, columnspan=2, pady=5)
        
        generate_btn = tk.Button(
            button_frame, 
            text="出力＆クリップボードに追加", 
            font=font,
            command=lambda: self.generate_alphabet(room_number.get(), floor_number.get(), add_floor.get())
        )
        generate_btn.grid(row=0, column=0, padx=5)
        
        close_btn = tk.Button(button_frame, text="閉じる", font=font, command=self.root.destroy)
        close_btn.grid(row=0, column=1, padx=5)
        
        back_btn = tk.Button(
            button_frame, 
            text="戻る", 
            font=font, 
            command=self.back_to_selection
        )
        back_btn.grid(row=0, column=2, padx=5)
        
        # 出力テキストエリア
        self.output_text = tk.Text(self.root, width=35, height=15, font=font)
        self.output_text.grid(row=5, column=0, columnspan=2, padx=5, pady=5)
        
        # スクロールバーの追加
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.output_text.yview)
        scrollbar.grid(row=5, column=2, sticky="ns")
        self.output_text.configure(yscrollcommand=scrollbar.set)
        
        # Ctrl+Enterのキーバインド
        self.root.bind('<Control-Return>', lambda event: self.copy_and_close())
        
        # Enter キーでボタンを押す
        self.root.bind('<Return>', lambda event: generate_btn.invoke())
        
        # Ctrl+Enterの説明ラベル
        info_label = tk.Label(
            self.root, 
            text="Ctrl + Enter で出力し、画面を閉じる", 
            font=('', 8),
            fg="gray50"
        )
        info_label.grid(row=6, column=0, columnspan=2, pady=(0, 5))
        
        self.root.mainloop()
        
    def generate_alphabet(self, rooms, floors, add_floor):
        """アルファベットモードの部屋番号生成"""
        try:
            rooms = int(rooms)
            floors = int(floors)
            
            if rooms <= 0 or floors <= 0:
                self.output_text.delete(1.0, tk.END)
                self.output_text.insert(tk.END, "エラー: 部屋数と階数は1以上の値を入力してください。")
                return
            
            output = ""
            
            for floor in range(1, floors + 1):
                for room in range(rooms):
                    # 0から始めて A, B, C...と変換
                    alpha = chr(65 + room)  # 65はAのASCIIコード
                    
                    if add_floor:
                        # 階数を追加
                        room_id = f"{floor}{alpha}"
                    else:
                        room_id = alpha
                    
                    output += f"{room_id} "
                
                output = output.rstrip() + "\n"
            
            # 最後の改行を削除
            output = output.rstrip()
            
            # 出力の更新
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, output)
            
            # クリップボードにコピー
            pyperclip.copy(output)
            
        except ValueError:
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, "エラー: 部屋数と階数には数字を入力してください。")

    def back_to_selection(self):
        """モード選択画面に戻る"""
        self.root.destroy()
        self.__init__()

    def generate_normal(self, rooms, floors, include_4, include_9):
        """通常モードの部屋番号生成"""
        try:
            rooms = int(rooms)
            floors = int(floors)
            
            if rooms <= 0 or floors <= 0:
                self.output_text.delete(1.0, tk.END)
                self.output_text.insert(tk.END, "エラー: 部屋数と階数は1以上の値を入力してください。")
                return
            
            output = ""
            
            for floor in range(1, floors + 1):
                floor_num = floor * 100
                room_counter = 0
                added_rooms = 0
                
                while added_rooms < rooms:
                    room_counter += 1
                    # 4または9で終わる部屋を含めるかどうか
                    if (room_counter % 10 == 4 and not include_4) or (room_counter % 10 == 9 and not include_9):
                        continue
                        
                    room = floor_num + room_counter
                    output += f"{room} "
                    added_rooms += 1
                
                output = output.rstrip() + "\n"
            
            # 最後の改行を削除
            output = output.rstrip()
            
            # 出力の更新
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, output)
            
            # クリップボードにコピー
            pyperclip.copy(output)
            
        except ValueError:
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, "エラー: 部屋数と階数には数字を入力してください。")

    def generate_extended(self, building_info, include_4):
        """拡張モードの部屋番号生成"""
        output = ""
        has_valid_building = False
        
        for i, (bldg_name, room_number, floor_number) in enumerate(building_info):
            bldg = bldg_name.get()
            
            try:
                # 空の入力をチェック
                if not room_number.get() or not floor_number.get():
                    continue
                    
                rooms = int(room_number.get())
                floors = int(floor_number.get())
                
                if rooms <= 0 or floors <= 0:
                    continue
                
                has_valid_building = True
                
                for floor in range(1, floors + 1):
                    floor_num = floor * 100
                    room_counter = 0
                    added_rooms = 0
                    
                    while added_rooms < rooms:
                        room_counter += 1
                        # 4で終わる部屋を含めるかどうか
                        if room_counter % 10 == 4 and not include_4:
                            continue
                            
                        room = floor_num + room_counter
                        # 棟名がある場合は追加
                        if bldg:
                            output += f"{bldg}{room} "
                        else:
                            output += f"{room} "
                        added_rooms += 1
                    
                    output = output.rstrip() + "\n"
                
                # 各棟の間に改行を追加
                output += "\n"
                
            except ValueError:
                continue
        
        # 最後の改行を削除
        output = output.rstrip()
        
        # 出力の更新
        self.output_text.delete(1.0, tk.END)
        
        if has_valid_building:
            self.output_text.insert(tk.END, output)
            # クリップボードにコピー
            pyperclip.copy(output)
        else:
            self.output_text.insert(tk.END, "エラー: 有効な棟情報がありません。\n少なくとも1つの棟に部屋数と階数を入力してください。")

    def copy_and_close(self):
        """出力を生成してクリップボードにコピーし、1秒後に画面を閉じる"""
        if hasattr(self, 'output_text') and self.output_text:
            # 現在の画面がどのモードか確認して出力生成
            try:
                # ウィジェットを探して現在のモードを判断
                entries = [w for w in self.root.winfo_children() if isinstance(w, tk.Entry)]
                checkbuttons = [w for w in self.root.winfo_children() if isinstance(w, tk.Checkbutton)]
                
                # ウィンドウタイトルで判断
                window_title = self.root.title()
                
                if "アルファベット" in window_title:
                    # アルファベットモード
                    if entries and len(entries) >= 2 and checkbuttons and len(checkbuttons) >= 1:
                        room_entry = entries[0]
                        floor_entry = entries[1]
                        add_floor_cb = checkbuttons[0]
                        add_floor = add_floor_cb.var.get() if hasattr(add_floor_cb, 'var') else True
                        
                        self.generate_alphabet(room_entry.get(), floor_entry.get(), add_floor)
                elif "通常" in window_title:
                    # 通常モード
                    if entries and len(entries) >= 2 and checkbuttons and len(checkbuttons) >= 1:
                        room_entry = entries[0]
                        floor_entry = entries[1]
                        include_4_cb = checkbuttons[0]
                        include_9_cb = None
                        
                        if len(checkbuttons) >= 2:
                            include_9_cb = checkbuttons[1]
                        
                        include_4 = include_4_cb.var.get() if hasattr(include_4_cb, 'var') else False
                        include_9 = include_9_cb.var.get() if include_9_cb and hasattr(include_9_cb, 'var') else False
                        
                        self.generate_normal(room_entry.get(), floor_entry.get(), include_4, include_9)
                elif "連棟部屋作成" in window_title:
                    # 拡張モード - キャンバスを探す
                    canvas = None
                    for widget in self.root.winfo_children():
                        if isinstance(widget, tk.Frame):
                            for w in widget.winfo_children():
                                if isinstance(w, tk.Canvas):
                                    canvas = w
                                    break
                    
                    if canvas:
                        # スクロール可能フレームを取得
                        scrollable_frame = canvas.children.get('!frame')
                        if scrollable_frame:
                            # 棟情報の取得
                            building_info = []
                            include_4 = False
                            
                            for widget in scrollable_frame.winfo_children():
                                if isinstance(widget, tk.LabelFrame) and "棟目" in widget['text']:
                                    entries = [w for w in widget.winfo_children() if isinstance(w, tk.Entry)]
                                    if len(entries) >= 3:
                                        bldg_name = tk.StringVar(value=entries[0].get())
                                        room_number = tk.StringVar(value=entries[1].get())
                                        floor_number = tk.StringVar(value=entries[2].get())
                                        building_info.append((bldg_name, room_number, floor_number))
                                elif isinstance(widget, tk.Frame):
                                    for w in widget.winfo_children():
                                        if isinstance(w, tk.Checkbutton) and "4の部屋" in str(w['text']):
                                            include_4 = w.var.get() if hasattr(w, 'var') else False
                            
                            if building_info:
                                self.generate_extended(building_info, include_4)
            except Exception as e:
                # エラーが発生した場合、現在の出力をそのまま使用
                print(f"出力生成中にエラーが発生しました: {e}")
                
            # 現在の出力テキストをコピー
            output = self.output_text.get(1.0, tk.END).rstrip()
            if output:
                pyperclip.copy(output)
                
                # 成功メッセージを表示
                status_label = tk.Label(
                    self.root,
                    text="クリップボードにコピーしました！",
                    font=('', 10),
                    fg="green",
                    bg="white"
                )
                
                # ラベルを最前面に表示
                if hasattr(self, 'output_text'):
                    x = self.output_text.winfo_x() + self.output_text.winfo_width() // 2 - 100
                    y = self.output_text.winfo_y() + self.output_text.winfo_height() // 2 - 10
                    status_label.place(x=x, y=y)
                else:
                    status_label.pack(pady=10)
                
                # ウィンドウを更新して表示を確実にする
                self.root.update()
                
                # 1秒後に画面を閉じる
                self.root.after(1000, self.root.destroy)
            else:
                # 出力が空の場合はすぐに閉じる
                self.root.destroy()
        else:
            # 出力テキストが見つからない場合はすぐに閉じる
            if self.root:
                self.root.destroy()

    def run(self):
        """アプリケーションを実行"""
        self.root.mainloop()

if __name__ == "__main__":
    app = RoomGenerator()
    app.run()