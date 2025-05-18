import win32com.client
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import IntVar, BooleanVar, StringVar
import pythoncom

def get_excel_application():
    """
    既存のExcelアプリケーションを取得するか、なければ新しく作成する
    """
    try:
        # 既存のExcelアプリケーションを取得
        excel = win32com.client.GetActiveObject("Excel.Application")
        return excel
    except pythoncom.com_error:
        # Excelが起動していない場合は新しいインスタンスを作成
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        return excel

def align_excel_objects(root_window):
    """
    Excelで選択された図形オブジェクトを等間隔に整列させる関数
    """
    # Excelアプリケーションの取得
    try:
        excel = get_excel_application()
    except Exception as e:
        messagebox.showerror("エラー", f"Excelアプリケーションの取得に失敗しました: {str(e)}")
        return
    
    try:
        # アクティブなブックのアクティブシートの取得
        workbook = excel.ActiveWorkbook
        if not workbook:
            messagebox.showwarning("警告", "開いているExcelブックがありません。")
            return
            
        worksheet = workbook.ActiveSheet
        
        # 選択されているオブジェクトの取得
        selected_shapes = excel.Selection.ShapeRange
        
        # 選択されているオブジェクトの数が3未満の場合、メッセージを表示して終了
        if selected_shapes.Count < 3:
            messagebox.showwarning("警告", "3つ以上の図形オブジェクトを選択してください。")
            return
        
        # 画面更新をオフにして高速化
        excel.ScreenUpdating = False
        
        # 選択されているオブジェクトの位置情報を取得
        shape_info = []
        for i in range(1, selected_shapes.Count + 1):
            shape = selected_shapes.Item(i)
            shape_info.append({
                'shape': shape,
                'left': shape.Left,
                'top': shape.Top,
                'width': shape.Width,
                'height': shape.Height
            })
        
        # 左上と右下の最大位置を確認
        min_left = min(shape['left'] for shape in shape_info)
        min_top = min(shape['top'] for shape in shape_info)
        max_right = max(shape['left'] + shape['width'] for shape in shape_info)
        max_bottom = max(shape['top'] + shape['height'] for shape in shape_info)
        
        # オブジェクトを行と列に分類
        # まず、Y座標でグループ化して行を特定（Y座標が近いものは同じ行とみなす）
        y_positions = [shape['top'] for shape in shape_info]
        y_positions.sort()
        
        # Y座標の差が大きい場所で区切り、行のグループを作成
        row_groups = []
        current_row = [y_positions[0]]
        
        for i in range(1, len(y_positions)):
            if y_positions[i] - y_positions[i-1] > 20:  # 20ピクセル以上離れていたら別の行と判断
                row_groups.append(current_row)
                current_row = [y_positions[i]]
            else:
                current_row.append(y_positions[i])
        
        row_groups.append(current_row)
        
        # 各行の代表値（平均値）を計算
        row_centers = [sum(row) / len(row) for row in row_groups]
        
        # 各オブジェクトを行ごとに分類
        rows = [[] for _ in range(len(row_groups))]
        for shape_data in shape_info:
            closest_row = min(range(len(row_centers)), key=lambda i: abs(shape_data['top'] - row_centers[i]))
            rows[closest_row].append(shape_data)
        
        # 各行内でオブジェクトをX座標でソート
        for row in rows:
            row.sort(key=lambda x: x['left'])
        
        # 行と列の数を決定
        num_rows = len(rows)
        num_cols = max(len(row) for row in rows)
        
        # 横方向と縦方向の間隔を計算
        shape_widths = [shape['width'] for shape in shape_info]
        shape_heights = [shape['height'] for shape in shape_info]
        avg_width = sum(shape_widths) / len(shape_widths)
        avg_height = sum(shape_heights) / len(shape_heights)
        
        # 利用可能な幅と高さ
        available_width = max_right - min_left
        available_height = max_bottom - min_top
        
        # 間隔の計算
        horizontal_gap = (available_width - (num_cols * avg_width)) / (num_cols - 1) if num_cols > 1 else 0
        vertical_gap = (available_height - (num_rows * avg_height)) / (num_rows - 1) if num_rows > 1 else 0
        
        # 各オブジェクトを格子状に配置
        for row_idx, row in enumerate(rows):
            for col_idx, shape_data in enumerate(row):
                shape = shape_data['shape']
                # 新しい位置を計算
                new_left = min_left + (col_idx * (avg_width + horizontal_gap))
                new_top = min_top + (row_idx * (avg_height + vertical_gap))
                
                # 位置を設定
                shape.Left = new_left
                shape.Top = new_top
        
        # 処理完了メッセージ
        messagebox.showinfo("完了", "図形の整列が完了しました。")
    except Exception as e:
        messagebox.showerror("エラー", f"図形の整列中にエラーが発生しました: {str(e)}")
    finally:
        # 画面更新を元に戻す
        excel.ScreenUpdating = True
    
    # 処理完了後にアプリケーションを終了しない
    # アプリケーションだけ閉じる
    if root_window:
        root_window.destroy()

def generate_room_numbers():
    """
    部屋番号の図形を生成する関数
    """
    global start_room_var, room_order_var, root
    
    # パラメータの取得
    room_count = room_count_var.get()
    floor_count = floor_count_var.get()
    include_4 = include_4_var.get()
    include_9 = include_9_var.get()
    start_room = start_room_var.get()
    room_order = room_order_var.get()  # 1: 若番順, 2: 老番順
    
    # 入力バリデーション
    if room_count <= 0 or floor_count <= 0:
        messagebox.showwarning("警告", "部屋数と階数は1以上の整数を入力してください。")
        return
    
    if not start_room:
        messagebox.showwarning("警告", "開始部屋番号を入力してください。")
        return
    
    # 開始部屋番号の解析
    start_is_numeric = start_room.isdigit()
    start_floor = 0
    start_room_num = ""
    
    if start_is_numeric:
        # 数字の場合、階数と部屋番号に分解
        if len(start_room) >= 3:
            start_floor = int(start_room[:-2])
            start_room_num = int(start_room[-2:])
        else:
            # 3桁未満の場合はそのまま部屋番号として扱う
            start_room_num = int(start_room)
    else:
        # 文字の場合、最初の文字を使用
        start_room_num = start_room[0]
    
    # Excelアプリケーションの取得
    try:
        excel = get_excel_application()
    except Exception as e:
        messagebox.showerror("エラー", f"Excelアプリケーションの取得に失敗しました: {str(e)}")
        return
    
    # 画面更新をオフにして高速化
    excel.ScreenUpdating = False
    
    try:
        # アクティブなブックのアクティブシートの取得
        workbook = excel.ActiveWorkbook
        if not workbook:
            messagebox.showwarning("警告", "開いているExcelブックがありません。")
            return
            
        worksheet = workbook.ActiveSheet
        
        # ピクセルからcmへの変換（1cm = 28.35ピクセル）
        cm_to_pixels = 28.35
        
        # テキストボックスのサイズと間隔を設定
        box_width = 1 * cm_to_pixels  # 1cm
        box_height = 1 * cm_to_pixels  # 1cm
        horizontal_spacing = 0.5 * cm_to_pixels  # 横方向の間隔1cm
        vertical_spacing = 1 * cm_to_pixels  # 縦方向の間隔1cm
        
        # アクティブセルの位置を使用して図形を配置
        active_cell = excel.ActiveCell
        left_position = active_cell.Left
        top_position = active_cell.Top
        
        # 部屋番号の生成と図形の作成
        # 階数が多い順に生成（上の階から配置するため）
        for floor in range(floor_count, 0, -1):
            floor_index = floor_count - floor  # 上の階から配置するための逆インデックス
            
            # 現在の階の部屋番号のリストを作成
            room_numbers = []
            
            if start_is_numeric:
                # 数字の部屋番号の場合
                current_room_num = start_room_num
                for i in range(room_count * 2):  # 余裕を持って多めに生成
                    if len(room_numbers) >= room_count:
                        break
                    
                    # 4号室と9号室のチェック
                    if (current_room_num % 10 == 4 and not include_4) or (current_room_num % 10 == 9 and not include_9):
                        current_room_num += 1
                        continue
                    
                    room_number = floor * 100 + current_room_num
                    room_numbers.append(str(room_number))
                    current_room_num += 1
            else:
                # 文字の部屋番号の場合
                current_char = start_room_num
                for i in range(room_count):
                    room_numbers.append(f"{floor}{current_char}")
                    # 次の文字へ（例：'A' -> 'B'）
                    current_char = chr(ord(current_char) + 1)
            
            # 老番順の場合はリストを逆順にする
            if room_order == 2:  # 老番順
                room_numbers.reverse()
            
            # 図形を作成
            for idx, room_number in enumerate(room_numbers):
                # 図形の位置計算
                current_left = left_position + idx * (box_width + horizontal_spacing)
                current_top = top_position + floor_index * (box_height + vertical_spacing)
                
                # テキストボックス作成
                shape = worksheet.Shapes.AddTextbox(
                    1,  # msoTextOrientationHorizontal
                    current_left,
                    current_top,
                    box_width,
                    box_height
                )
                
                # テキストと書式設定
                shape.TextFrame.Characters().Text = str(room_number)
                shape.TextFrame.HorizontalAlignment = -4108  # xlCenter
                shape.TextFrame.VerticalAlignment = -4108  # xlCenter
                
                # 内部の余白を設定 (0.08cm = ~2.27ピクセル)
                shape.TextFrame.MarginLeft = 2.27
                shape.TextFrame.MarginRight = 2.27
                shape.TextFrame.MarginTop = 2.27
                shape.TextFrame.MarginBottom = 2.27
                
                # フォント設定
                shape.TextFrame.Characters().Font.Name = "Meiryo UI"
                shape.TextFrame.Characters().Font.Size = 9
                shape.TextFrame.Characters().Font.Color = 0  # 黒
                
                # 枠線設定
                shape.Line.Weight = 0.75
                shape.Line.ForeColor.RGB = 0  # 黒
                
                # 塗りつぶし設定
                shape.Fill.ForeColor.RGB = 16777215  # 白
    except Exception as e:
        messagebox.showerror("エラー", f"部屋番号の生成中にエラーが発生しました: {str(e)}")
    finally:
        # 画面更新を元に戻す
        excel.ScreenUpdating = True
    
    messagebox.showinfo("完了", f"{floor_count}階分、各階{room_count}部屋の図形を生成しました。")
    
    # 処理完了後にアプリケーションを終了しない
    # アプリケーションだけ閉じる
    root.destroy()

def create_gui():
    """
    GUIを作成する関数
    """
    global root, room_count_var, floor_count_var, include_4_var, include_9_var, start_room_var, room_order_var
    
    root = tk.Tk()
    root.title("Excel図形操作ツール")
    root.geometry("450x500")  # ウィンドウサイズを大きくする
    
    # モード選択用のラジオボタン
    mode_frame = ttk.LabelFrame(root, text="機能選択")
    mode_frame.pack(padx=10, pady=10, fill="x")
    
    mode_var = IntVar(value=1)
    ttk.Radiobutton(mode_frame, text="図形の整列", variable=mode_var, value=1).pack(anchor="w", padx=10, pady=5)
    ttk.Radiobutton(mode_frame, text="部屋番号の生成", variable=mode_var, value=2).pack(anchor="w", padx=10, pady=5)
    
    # 部屋番号生成用の設定フレーム
    settings_frame = ttk.LabelFrame(root, text="部屋番号生成設定")
    settings_frame.pack(padx=10, pady=10, fill="x")
    
    # 部屋数入力
    room_count_frame = ttk.Frame(settings_frame)
    room_count_frame.pack(fill="x", padx=10, pady=5)
    ttk.Label(room_count_frame, text="1階あたりの部屋数:").pack(side="left")
    room_count_var = IntVar(value=3)
    ttk.Entry(room_count_frame, textvariable=room_count_var, width=5).pack(side="left", padx=5)
    
    # 階数入力
    floor_count_frame = ttk.Frame(settings_frame)
    floor_count_frame.pack(fill="x", padx=10, pady=5)
    ttk.Label(floor_count_frame, text="階数:").pack(side="left")
    floor_count_var = IntVar(value=2)
    ttk.Entry(floor_count_frame, textvariable=floor_count_var, width=5).pack(side="left", padx=5)
    
    # 開始部屋番号入力
    start_room_frame = ttk.Frame(settings_frame)
    start_room_frame.pack(fill="x", padx=10, pady=5)
    ttk.Label(start_room_frame, text="開始部屋番号:").pack(side="left")
    start_room_var = StringVar(value="101")
    ttk.Entry(start_room_frame, textvariable=start_room_var, width=10).pack(side="left", padx=5)
    ttk.Label(start_room_frame, text="(数字または文字を入力)").pack(side="left")
    
    # 部屋番号の順序
    room_order_frame = ttk.Frame(settings_frame)
    room_order_frame.pack(fill="x", padx=10, pady=5)
    room_order_var = IntVar(value=1)  # 1: 若番順, 2: 老番順
    ttk.Label(room_order_frame, text="部屋番号の順序:").pack(side="left")
    ttk.Radiobutton(room_order_frame, text="若番順 (101,102,103...)", variable=room_order_var, value=1).pack(side="left", padx=5)
    
    # 老番順のラジオボタンを別の行に配置
    elder_order_frame = ttk.Frame(settings_frame)
    elder_order_frame.pack(fill="x", padx=10, pady=2)
    ttk.Radiobutton(elder_order_frame, text="老番順 (...103,102,101)", variable=room_order_var, value=2).pack(side="left", padx=25)  # 余白を増やして整列
    
    # チェックボックス
    checkbox_frame = ttk.Frame(settings_frame)
    checkbox_frame.pack(fill="x", padx=10, pady=5)
    include_4_var = BooleanVar(value=False)
    include_9_var = BooleanVar(value=False)
    ttk.Checkbutton(checkbox_frame, text="4号室を含める", variable=include_4_var).pack(side="left", padx=5)
    ttk.Checkbutton(checkbox_frame, text="9号室を含める", variable=include_9_var).pack(side="left", padx=5)
    
    # 実行ボタンの関数定義
    def execute_function():
        if mode_var.get() == 1:
            align_excel_objects(root)
        else:
            generate_room_numbers()
    
    # 実行ボタンをフレームに配置して下部に表示
    button_frame = ttk.Frame(root)
    button_frame.pack(fill="x", padx=10, pady=20)
    ttk.Button(button_frame, text="実行", command=execute_function).pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()