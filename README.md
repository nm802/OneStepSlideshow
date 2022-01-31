# OneStepSlideshow

画像ファイルからMicrosoft Powerpointのスライドショーを作ります。

# Features

Powerpoint標準の機能（フォトアルバム作成機能）はわりとめんどい。

画像を列挙したスライドが一発でほしいときに使えます。

# Requirement

* Microsoft Powerpoint
* python-pptx
* Pillow

# Installation

```bash
pip install python-pptx, Pillow
```

# Usage

#### python実行環境の準備
venvディレクトリ内に仮想環境ある前提の.batファイルになっています。
```make_slideshow.bat
call "venv\Scripts\activate.bat"
```
make_slideshow.bat 内実行環境を適宜修正してください。

#### 右クリックメニューに入れる 
「ファイル名を指定して実行」で「shell:sendto」と入力し，出てきたフォルダにmake_slideshow.batのショートカットを入れる。

名前を適宜「PowerPointスライドショーを作成」などに変える。

#### 実行
画像ファイルを一つまたは複数選択し，右クリック→送る→上記で作成したショートカットを選択。

画像ファイルがあるディレクトリに.pptxファイルができます。

# License
"OneStepSlideshow" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).



