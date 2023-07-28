import os

def add_suffix_to_filename(folder_path, suffix):
    """
    Oberview:
        指定したﾌｫﾙﾀﾞ内のすべてのﾌｧｲﾙ名の末尾に指定した文字列を追加
    Args:
        folder_path(str): ﾌｫﾙﾀﾞのﾊﾟｽ
        suffix(str): 追加する文字列
    Returns:
        None
    """

    # ﾌｫﾙﾀﾞ内のすべてのﾌｧｲﾙを取得
    file_list = os.listdir(folder_path)

    # 各ﾌｧｲﾙ名に指定した文字列を追加してﾘﾈｰﾑ
    for filename in file_list:
        # ﾌｧｲﾙ名と拡張子を分類
        file_name, file_extension = os.path.splitext(filename)

        # 拡張子を含むﾌｧｲﾙ名に指定した文字列を追加
        new_filename = file_name + suffix + file_extension

        old_path = os.path.join(folder_path, filename) # 元のﾌｧｲﾙﾊﾟｽ
        new_path = os.path.join(folder_path, new_filename) # 新しいﾌｧｲﾙのﾊﾟｽ
        os.rename(old_path, new_path) # ﾌｧｲﾙ名を変更

folder_path = r"" # 置換を行うﾌｫﾙﾀﾞﾊﾟｽ
suffix = "" # ﾌｧｲﾙに追加する文字列
add_suffix_to_filename(folder_path, suffix)
