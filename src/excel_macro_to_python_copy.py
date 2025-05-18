#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
VBAコード「星つきxlsx保存」の処理をPythonに変換したもの

※このコードは、ExcelのCOM経由でActiveWorkbook/ActiveSheetを操作するため、
　Excelが既に起動しており対象のブックがアクティブになっている必要があります。
※SaveAsダイアログの代替として、端末からの入力でファイル名を指定する処理にしています。
※ pywin32を利用してExcelを操作します。（pip install pywin32 でインストール可能）
"""

import os
import shutil
import glob
import datetime
import subprocess
import pythoncom
import win32com.client


def extract_prefecture(address):
    """住所から都道府県名を抽出する関数"""
    prefectures = ["北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
                   "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
                   "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
                   "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
                   "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
                   "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
                   "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"]

    for prefecture in prefectures:
        if address.startswith(prefecture):
            return prefecture
    return ""


def is_chubu_address(address, chubu_addresses):
    """住所が中部地方かどうかを判定する関数（都道府県名を抽出して判定）"""
    prefecture = extract_prefecture(address)
    return prefecture in chubu_addresses


def is_kansai_address(address, kansai_addresses):
    """住所が関西地方かどうかを判定する関数（都道府県名を抽出して判定）"""
    prefecture = extract_prefecture(address)
    return prefecture in kansai_addresses


def find_folders_starting_with(base_path, prefix):
    result = []
    if not base_path.endswith(os.sep):
        base_path += os.sep
    if not os.path.isdir(base_path):
        return result
    # os.scandir() を利用してディレクトリ走査を高速化
    with os.scandir(base_path) as it:
        for entry in it:
            if entry.is_dir():
                item = entry.name
                if item.startswith(prefix):
                    if len(item) == len(prefix):
                        result.append(entry.path)
                    else:
                        next_char = item[len(prefix):len(prefix)+1]
                        if next_char in ["_", " ", "-", "－"]:
                            result.append(entry.path)
    return result


def is_folder_open(folder_path):
    """
    指定されたフォルダが Windows Explorer で開かれている場合、そのウィンドウをアクティブにして True を返す関数
    """
    target = os.path.normpath(folder_path).lower()
    shell = win32com.client.Dispatch("Shell.Application")
    for window in shell.Windows():
        try:
            url = window.LocationURL
            if url.startswith("file:///"):
                # URLをファイルパスに変換
                open_folder = url[8:].replace('/', '\\')
                open_folder = os.path.normpath(open_folder).lower()
                if open_folder == target:
                    # 既に開かれているウィンドウをアクティブにする
                    try:
                        window.Document.parentWindow.focus()
                    except Exception as e:
                        print("ウィンドウをアクティブにできませんでした:", e)
                    return True
        except Exception:
            continue
    return False


def get_file_name(full_path):
    return os.path.basename(full_path)


def move_existing_file(target_file_path, base_folder):
    """
    移動先フォルダ内に今日の日付を使った old フォルダへ、ファイル名の先頭に (yyyy-mm-dd_作成) を付与して移動する関数
    """
    today_str = datetime.datetime.today().strftime("%Y-%m-%d")
    old_folder = os.path.join(base_folder, f"{today_str}_old")
    if not os.path.isdir(old_folder):
        try:
            os.makedirs(old_folder)
        except Exception as e:
            print(f"フォルダ '{today_str}_old' の作成に失敗しました。エラー: {e}")
            return False
    new_file_name = get_file_name(target_file_path)
    old_file_path = os.path.join(old_folder, new_file_name)

    if os.path.exists(old_file_path):
        print(
            f"ファイル '{get_file_name(target_file_path)}' は既に '{today_str}_old' フォルダに存在します。処理を中止します。")
        return False
    try:
        os.rename(target_file_path, old_file_path)
    except Exception as e:
        print(
            f"ファイル '{get_file_name(target_file_path)}' を '{today_str}_old' フォルダに移動できませんでした。エラー: {e}")
        return False
    return True


def save_starred_xlsx():
    pythoncom.CoInitialize()
    try:
        # Excelオブジェクトの取得
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.ActiveWorkbook
        if wb is None:
            print("アクティブなブックがありません。")
            return

        try:
            ws = excel.ActiveSheet
        except Exception:
            print("ActiveSheetの取得に失敗しました。")
            return
        excel.DisplayAlerts = False

        # 名前定義の辞書作成（ワークブック・ワークシート両レベル）
        defined_names_dict = {}
        # ワークブックレベル
        for name_obj in wb.Names:
            try:
                refers = name_obj.RefersTo
            except Exception:
                try:
                    refers = name_obj.RefersToLocal
                except:
                    refers = None
            if refers:
                defined_names_dict[name_obj.Name] = refers
                if '.' in name_obj.Name:
                    short_name = name_obj.Name.split('.')[-1]
                    if short_name not in defined_names_dict:
                        defined_names_dict[short_name] = refers
        # ワークシートレベル
        for ws_item in wb.Worksheets:
            for name_obj in ws_item.Names:
                try:
                    refers = name_obj.RefersTo
                except Exception:
                    try:
                        refers = name_obj.RefersToLocal
                    except:
                        refers = None
                if refers and name_obj.Name not in defined_names_dict:
                    defined_names_dict[name_obj.Name] = refers
                    if '.' in name_obj.Name:
                        short_name = name_obj.Name.split('.')[-1]
                        if short_name not in defined_names_dict:
                            defined_names_dict[short_name] = refers

        # ヘルパー関数：名前定義から値取得
        def get_value_by_defined_name(name):
            if name in defined_names_dict:
                try:
                    return excel.Range(name).Value
                except Exception:
                    pass
            if '.' in name:
                short = name.split('.')[-1]
                if short in defined_names_dict:
                    try:
                        return excel.Range(short).Value
                    except Exception:
                        pass
            try:
                return excel.Range(name).Value
            except Exception:
                if '.' in name:
                    try:
                        return excel.Range(name.split('.')[-1]).Value
                    except Exception:
                        return None
                return None

        # ユーザープロファイルの取得（フォールバックとしてexpanduserを利用）
        user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        download_folder_path = os.path.join(
            os.path.expanduser("~"), "Downloads") + os.sep

        # 値の取得と処理
        building_no_value = get_value_by_defined_name("HOUSES.BUILDING_NO")
        if isinstance(building_no_value, (int, float)):
            building_no = str(int(building_no_value))
        else:
            building_no = str(building_no_value or "")

        building_name = str(get_value_by_defined_name(
            "HOUSES.BUILDING_NAME") or "")
        addname1 = str(get_value_by_defined_name("ADD_NAME1") or "")
        addname2 = str(get_value_by_defined_name("ADD_NAME2") or "")
        houses_address = get_value_by_defined_name("HOUSES.ADDRESS") or ""
        tentative_name = get_value_by_defined_name(
            "HOUSES.TENTATIVE_NAME") or ""

        try:
            service_id = get_value_by_defined_name("HOUSES.SERVICE_ID")
            if service_id is None:
                raise Exception
        except Exception:
            service_id = input(
                "サービスが見つからないか名前定義が見つかりませんでした。サービス名を入力してください。(デフォルト: SS): ").strip() or "SS"
            ws.Range("DI3").Value = service_id

        try:
            building_state = get_value_by_defined_name("HOUSES.BUILDING_STATE")
            if building_state is None:
                raise Exception
        except Exception:
            building_state = input(
                "新築/既存が不明、もしくは名前定義が見つかりませんでした。新築/既存どちらかを入力してください。(デフォルト: 既存): ").strip() or "既存"
            ws.Range("EB16").Value = building_state

        invalid_chars = ["/", "\\", ":", "*", "?", "<", ">", "|", '"']
        for ch in invalid_chars:
            building_no = building_no.replace(ch, "_")
            building_name = building_name.replace(ch, "_")
            addname2 = addname2.replace(ch, "_")

        fix_addname1 = f"({addname1}) " if len(addname1) > 0 else ""
        full_file_name = os.path.join(
            download_folder_path, f"☆{fix_addname1}{building_no}_{building_name}_{addname2}.xlsx")

        if os.path.exists(full_file_name):
            if os.path.abspath(wb.FullName) == os.path.abspath(full_file_name):
                print("警告: 現在のブックと同じファイル名が使用されているため、既存ファイルの削除をスキップします。")
            else:
                try:
                    os.remove(full_file_name)
                except PermissionError as e:
                    print(
                        f"既存ファイル '{full_file_name}' の削除に失敗しました。ファイルが他のプロセスにより使用中です。エラー: {e}")
                    return
        wb.SaveAs(full_file_name, FileFormat=51, CreateBackup=False)

        # フォルダパスやキー名、漢字部分を適当に変更した例
        folder_map = {
            "FOO_new": os.path.join(user_profile, "Documents", "TestFolder", "新規案件"),
            "FOO_exist": os.path.join(user_profile, "Documents", "TestFolder", "既存案件"),
            "FOO_chubu": os.path.join(user_profile, "Documents", "TestFolder", "既存案件", "中部地方"),
            "FOO_kansai": os.path.join(user_profile, "Documents", "TestFolder", "関西支社", "関西案件"),
            "FOO_newex": os.path.join(user_profile, "Documents", "TestFolder", "新規案件", "導入済み"),
            "BAR_exist": os.path.join(user_profile, "Documents", "TestFolder", "BAR既存")
        }

        chubu_addresses = ["静岡県", "岐阜県", "長野県",
                           "愛知県", "新潟県", "三重県", "富山県", "石川県"]
        kansai_addresses = ["大阪府", "京都府", "兵庫県", "滋賀県", "和歌山県", "奈良県"]

        new_path = ""
        base_folder_path = ""
        if service_id == "FOO":
            if is_chubu_address(houses_address, chubu_addresses):
                base_folder_path = folder_map["FOO_chubu"]
            elif is_kansai_address(houses_address, kansai_addresses):
                base_folder_path = folder_map["FOO_kansai"]
            else:
                if building_state == "新築":
                    matched_ex = find_folders_starting_with(
                        folder_map["FOO_newex"], building_no)
                    if len(matched_ex) > 1:
                        print("導入済みフォルダ内に複数の該当フォルダが見つかりました。処理を中止します。")
                        print("\n".join(matched_ex))
                        return
                    elif len(matched_ex) == 1:
                        new_path = matched_ex[0]
                        print(f"導入済みフォルダに既存フォルダが見つかりました。格納先: {new_path}")
                    else:
                        matched_normal = find_folders_starting_with(
                            folder_map["FOO_new"], building_no)
                        if len(matched_normal) > 1:
                            print("新規案件フォルダ内に複数の該当フォルダが見つかりました。処理を中止します。")
                            print("\n".join(matched_normal))
                            return
                        elif len(matched_normal) == 1:
                            new_path = matched_normal[0]
                            print(f"新規案件フォルダに既存フォルダが見つかりました。格納先: {new_path}")
                        else:
                            new_path = os.path.join(
                                folder_map["FOO_new"], f"{building_no}_{tentative_name}")
                            if not os.path.isdir(new_path):
                                os.makedirs(new_path)
                                print(f"フォルダが存在しなかったため、新規に作成しました: {new_path}")
                elif building_state == "既存":
                    base_folder_path = folder_map["FOO_exist"]
                else:
                    print("HOUSES.BUILDING_STATEが新築でも既存でもありません。処理を中止します。")
                    return
        elif service_id == "BAR":
            base_folder_path = folder_map["BAR_exist"]
        else:
            print(f"ServiceIDがFOOでもBARでもありません: {service_id}")
            return

        if new_path == "":
            new_path = base_folder_path
            matched_folders = find_folders_starting_with(new_path, building_no)
            if len(matched_folders) > 1:
                print("BuildingNoを先頭に含むフォルダが複数見つかりました。処理を中止します。")
                print("\n".join(matched_folders))
                return
            elif len(matched_folders) == 1:
                new_path = matched_folders[0]
                print(f"以下のフォルダに自動的に格納します: {new_path}")
            elif len(matched_folders) == 0:
                if building_state == "新築":
                    new_path = os.path.join(
                        new_path, f"{building_no}_{tentative_name}")
                else:
                    new_path = os.path.join(
                        new_path, f"{building_no}_{building_name}")
                if not os.path.isdir(new_path):
                    os.makedirs(new_path)
                    print(f"フォルダが存在しなかったため、新規に作成しました: {new_path}")

        pdf_file_name = full_file_name.replace(".xlsx", ".pdf")
        wb.ExportAsFixedFormat(0, pdf_file_name, Quality=0, IncludeDocProperties=True,
                               IgnorePrintAreas=False, OpenAfterPublish=False)

        source_file_names = [full_file_name, pdf_file_name]
        target_file_names = [os.path.join(new_path, get_file_name(full_file_name)),
                             os.path.join(new_path, get_file_name(pdf_file_name))]
        moved_files = []
        copied_files = []

        for i, source in enumerate(source_file_names):
            target = target_file_names[i]
            if os.path.exists(target):
                file_creation_date = datetime.date.fromtimestamp(
                    os.path.getctime(target))
                if file_creation_date < datetime.date.today():
                    if not move_existing_file(target, new_path):
                        return
                    moved_files.append(get_file_name(target))
            try:
                shutil.copy2(source, target)
                copied_files.append(get_file_name(target))
            except Exception as e:
                print(f"ファイルコピーに失敗しました。エラー: {e}")
                return

        # ZIPファイルの移動処理
        zip_file_name = None

        # tentative_name が空の場合、building_noで始まるZIPファイルを検索する
        if tentative_name.strip() == "":
            pattern = os.path.join(download_folder_path, f"{building_no}*.zip")
            zip_files = glob.glob(pattern)
            if zip_files:
                zip_file_name = zip_files[0]  # 複数ヒットした場合は最初のものを使用
                print(
                    f"tentative_nameが空のため、対象ZIPファイルとして {zip_file_name} を選択しました。")
            else:
                print("ダウンロードフォルダ内に対象のZIPファイルが存在しません。")
        else:
            zip_file_name = os.path.join(
                download_folder_path, f"{building_no}_{tentative_name}.zip")
            if not os.path.exists(zip_file_name):
                print("ダウンロードフォルダ内に対象のZIPファイルが存在しません。")

        if zip_file_name and os.path.exists(zip_file_name):
            # リネーム処理：shutil.move() を利用して異なるドライブ間の移動にも対応
            if addname2 == "別件報告書":
                renamed_zip_file_name = os.path.join(
                    download_folder_path, "別件写真.zip")
            else:
                renamed_zip_file_name = os.path.join(
                    download_folder_path, "写真.zip")
            try:
                shutil.move(zip_file_name, renamed_zip_file_name)
            except Exception as e:
                print(f"ZIPファイルのリネームに失敗しました。エラー: {e}")
                return
            new_zip_file_name = os.path.join(
                new_path, get_file_name(renamed_zip_file_name))
            if os.path.exists(new_zip_file_name):
                if not move_existing_file(new_zip_file_name, new_path):
                    return
                moved_files.append(get_file_name(new_zip_file_name))
            try:
                shutil.move(renamed_zip_file_name, new_zip_file_name)
                copied_files.append(get_file_name(new_zip_file_name))
            except Exception as e:
                print(f"ZIPファイルを '{new_zip_file_name}' に移動できませんでした。エラー: {e}")
                return
        else:
            print("ZIPファイルが存在しないため、以降のZIP処理をスキップします。")

        print("処理完了:")
        print("格納フォルダ: ", new_path)
        print("格納ファイル:")
        for f in copied_files:
            print(f)

        if not is_folder_open(new_path):
            subprocess.Popen(["explorer.exe", new_path])
        edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        subprocess.Popen([edge_path, pdf_file_name])
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    save_starred_xlsx()
    import time
    time.sleep(3)
