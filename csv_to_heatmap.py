import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog

def select_csv_files():
    """ CSVﾌｧｲﾙﾊﾟｽ取得

    Overview:
        処理を行うCSVﾌｧｲﾙのパスを取得

    Returns:
        file_path (str) : 取得したﾃﾞｨﾚｸﾄﾘ
    """
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", ".csv")])
    root.destroy()

    return file_path

def create_heatmap_from_csv(csv_file_path):
    """ ﾋｰﾄﾏｯﾌﾟ作成

    Overview:
        指定したCSVﾌｧｲﾙからﾋｰﾄﾏｯﾌﾟを作成する

    Args:
        csv_file_path (str): ﾌｧｲﾙﾊﾟｽ
    
    Return:
        None
    """

    # CSVﾌｧｲﾙを pandas, DataFrames として読み込み
    df = pd.read_csv(csv_file_path, encoding='utf-8')

    # ﾃﾞｰﾀをﾋｰﾄﾏｯﾌﾟとして描画
    plt.rcParams["font.family"] = "MS Gothic" # 日本語対応
    plt.figure(figsize=(10, 8)) # ﾋｰﾄﾏｯﾌﾟのｻｲｽﾞを設定
    sns.heatmap(df.corr(), annot=True, cmap="YlGnBu", fmt=".1f") # ﾋｰﾄﾏｯﾌﾟを作成(相関係数)
    plt.title("") # ｸﾞﾗﾌﾀｲﾄﾙ
    plt.xlabel("") # X軸ﾗﾍﾞﾙ設定
    plt.ylabel("") # Y軸ﾗﾍﾞﾙ設定
    plt.show() # ｸﾞﾗﾌ表示

csv_file_path = select_csv_files()
if csv_file_path:
    create_heatmap_from_csv(csv_file_path)