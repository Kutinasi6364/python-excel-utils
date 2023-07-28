import pandas as pd
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog


def select_file_dialog():
    """ ﾌｧｲﾙ選択ﾀﾞｲｱﾛｸﾞ
    Overview:
        ﾌｧｲﾙをﾀﾞｲｱﾛｸﾞで選択する
    Args:
        None
    Returns:
        str: 選択したﾌｧｲﾙのﾊﾟｽ文字列
    """

    root = Tk()  # Tkinter ｳｲﾝﾄﾞｳ作成
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("", "")])  # ﾊﾟｽ取得

    root.destroy()  # ｳｲﾝﾄﾞｳ破棄

    return file_path


def csv_simple_linear_regression(csv_file_path, x_column, y_column):
    """ 単回帰分析
    Overview:
        CSVﾌｧｲﾙの線形回帰分析を行いｸﾞﾗﾌ表示
    Args:
        csv_file_path (str): CSVﾌｧｲﾙﾊﾟｽ
        x_column (str): X軸対象項目名
        y_column (str): Y軸対象項目名
    Return:
        None
    """

    # CSVファイルをpandas DataFrameとして読み込み
    df = pd.read_csv(csv_file_path)

    # xとyのデータを取得
    x_data = df[x_column].values.reshape(-1, 1)
    y_data = df[y_column].values

    # 単回帰モデルを作成して学習
    model = LinearRegression()
    model.fit(x_data, y_data)

    # モデルの係数と切片を取得
    slope = model.coef_[0]
    intercept = model.intercept_

    # 結果を出力
    print("回帰式: y = {:.2f}x + {:.2f}".format(slope, intercept))

    # 散布図と回帰直線をプロット
    plt.rcParams["font.family"] = "MS Gothic"  # 日本語対応
    plt.scatter(x_data, y_data, color='blue', label='データ')
    plt.plot(x_data, model.predict(x_data), color='red', label='回帰直線')
    plt.xlabel(x_column)
    plt.ylabel(y_column)
    plt.legend()
    plt.show()


csv_file_path = select_file_dialog()  # CSVﾌｧｲﾙを選択
x_column = "X"  # Y列名を指定(CSVﾌｧｲﾙ)
y_column = "Y"  # Y列名を指定(CSVﾌｧｲﾙ)
csv_simple_linear_regression(csv_file_path, x_column, y_column)

""" Reference
| 色名     | カラーコード |
|----------|--------------|
| 赤       | 'red'        |
| 緑       | 'green'      |
| 青       | 'blue'       |
| オレンジ | 'orange'     |
| ピンク   | 'pink'       |
| 黄色     | 'yellow'     |
| パープル | 'purple'     |
| グレー   | 'gray'       |
| ブラウン | 'brown'      |
"""
