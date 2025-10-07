開発者のGithubの0.8.0以降のWindows用リリース版をご使用の方は、
以下の方法で本GUIを使ってください。

（１）プログラムファイルのコピー

     下記の2つのファイルをLADAインストールフォルダーにコピー
     ※lada-cli.exeがあるフォルダ
　　　
     lada_gui.py
     LADA_LAUNCHER_FOR_GUI_RR.ps1

（２）ランチャーファイルのリネーム

     LADA_LAUNCHER_FOR_GUI_RR.ps1  を、
     LADA_LAUNCHER_FOR_GUI.ps1     に名前変更

（３）ffmpeg.exeのコピー

　　 LADAインストールフォルダの\_internal\bin\ffmpeg.exe　を
     LADAインストールフォルダにコピー

（３）起動方法

     LADAインストールフォルダから起動

     python lada_gui.py



＜主な変更点＞

Python関連のコード削除:

$PythonExe 変数を削除

PyTorch CUDAチェック機能を削除（EXEが内部で処理するため）

lada-cli.exeのパス変更:

$LadaCli = Join-Path $ScriptDir "lada-cli.exe"（ルートディレクトリに変更）

モデルファイルのパス変更:

検出モデルと修復モデルのパスを _internal\model_weights\ に変更

実行方法の変更:

Python経由での実行から直接 lada-cli.exe を実行するように変更

引数の構築方法を修正（lada-cli.exe を最初の引数から削除）

CUDA検出の簡素化:

PyTorchのCUDAチェックを削除し、NVIDIA GPUの存在のみで判断

バージョン表示の更新:

スクリプトのバナーに "(EXE Version)" を追加


