## Excelファイルの値を貼りつけるツール（windows用）

## Ctrl+Vを押すだけでExcelファイルの指定列の値を貼り付けられます。

1. 「clipBoardWindows.py」にExcelファイルのパスと列を指定します。
2. 「clipBoardWindows.py」を実行します。
3. 実行後にエディタで「Ctrl+V」をすると値を貼り付けられます。

* デフォルトでは「C:\python\sample.xlsx」の1シート目・1列目を指定しています。

## 必要なパッケージ
* pywin32のインストール（pythonのversionに合わせる）
  * https://sourceforge.net/projects/pywin32/files/?source=navbar

* xlrdのインストール
  * https://pypi.python.org/pypi/xlrd
  * xlrd-1.0.0.tar.gzを解凍し、setup.pyのあるフォルダで下記コマンドを実行。
```
python setup.py install
```
