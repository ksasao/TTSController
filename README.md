# TTSController
各種 Text-to-Speech エンジンを統一的に操作するライブラリです。VOICEROIDなどを自動化する簡易Webサーバもあります。

## 対応プラットフォーム
- Windows 10 (64bit)

## 対応音声合成ライブラリ
- VOICEROID+ 各種
- 音街ウナTalkEx
- VOICEROID2 各種
- ガイノイドTALK 各種
- かんたん！AITalk3 / 関西風 / Lite
- CeVIO CS6 / CS7
- SAPI5 (Windows10標準の音声合成機能。スタートメニュー>設定>時刻と言語>音声認識>音声の管理>音声の追加から各国語の音声が追加できます。API仕様により追加しても列挙されない音声があります。)

### 動作確認済みリスト
ライブラリ名はインストールされたフォルダなどを参照して機械的に抽出しているため、リストにないものでも音声合成エンジンが共通であれば動作する可能性が高いです。
|音声合成エンジン|ライブラリ名|
|---|---|
|VOICEROID+ EX|民安ともえ EX, 東北ずん子, 東北きりたん, 京町セイカ|
|音街ウナTalkEx|音街ウナ|
|VOICEROID2|琴葉 茜, 琴葉 葵, 結月ゆかり, 紲星あかり, 東北イタコ, 桜乃そら, ついなちゃん（標準語）, ついなちゃん（関西弁）|
|VOICEROID2 (VOICEROID+ EX からのアップグレード)|民安ともえ(v1), 東北ずん子(v1), 東北きりたん(v1), 京町セイカ(v1)|
|かんたん！AITalk3|あんず, かほ, ななこ, のぞみ, せいじ|
|かんたん！AITalk3 関西風|みやび, やまと|
|かんたん！AITalk3 LITE|あんず(LITE), かほ(LITE), ななこ(LITE), のぞみ(LITE), せいじ(LITE)|
|ガイノイドTALK|鳴花ヒメ, 鳴花ミコト|
|CeVIO CS6 / CS7|さとうささら, すずきつづみ, タカハシ|
|SAPI5|Microsoft Haruka Desktop, Microsoft David Desktop, Microsoft Zira Desktop, Microsoft Irina Desktop|

## ブラウザで音声合成する
この実装は簡易実装であり、音声合成ライブラリと同一のPC上で実行することを想定しています。インターネット上への公開は、セキュリティ上のリスクや音声合成ライブラリのライセンス上の問題がある可能性があります。

- [ビルド済み実行ファイル(v0.0.4beta)](https://github.com/ksasao/TTSController/releases/download/v0.0.4beta/SpeechWebServer_v0.0.4.zip) (2020/12/26更新)

### 準備
- SpeechWebServer のプロジェクトを Visual Studio 2019 でビルドして ```SpeechWebServer.exe``` を実行します(管理者権限が必要です)

### 利用方法
- ブラウザで http://localhost:1000/ を開くと現在の時刻を発話します
- http://localhost:1000/?text=こんにちは を開くと「こんにちは」と発話します。「こんにちは」の部分は任意の文字列を指定できます
- http://localhost:1000/?text=おやすみなさい&range=1.2&volume=1.0&pitch=0.8&speed=0.8 のように、音量(volume), 話速(speed), 高さ(pitch), 抑揚(range) を指定できます (かんたん！AITalk3 LITE, CeVIO, SAPI5を除く)
- VOICEROID+ 東北きりたんがインストールされている場合、http://localhost:1000/?name=東北きりたん&text=こんばんは を開くと東北きりたんの声で発話します。他の VOICEROID を利用する場合は、アプリ起動時に表示される「インストール済み音声合成ライブラリ」の表記を参考に、適宜 name の引数を変更してください。
- http://localhost:1000/?text=おはよう&speaker=和室 のように音声を再生するスピーカー名を指定することができます。カッコ内の文字列を前方一致で検索します。なお、Google Home デバイスは Windows から Bluetoothスピーカーとして接続ができ、任意の名前(「和室」など)を付けることが可能です。

![スピーカー名の表示](https://user-images.githubusercontent.com/179872/103144037-c823f200-4765-11eb-93a3-e202a8621ad2.png)

## TODO

### 制御機能
- [x] 話者の一覧取得
- [x] 話者に応じたTTS切り替え
- [ ] Bluetooth スピーカーの安定動作のための無音区間挿入
- [ ] 同時起動対応(先に起動しているほうに処理を委譲)

### 音声コントロール
- [x] 再生
- [x] 音量の取得・変更
- [x] 話速の取得・変更
- [x] ピッチの取得・変更
- [x] 抑揚の取得・変更
- [x] 発話中の音声停止
- [x] 合成した音声の保存
- [ ] 連続して文字列が入力されたときの対応
- [ ] 音声合成対象の文字列の途中に .wav ファイルを差し込み
- [ ] 音声合成対象の文字列の途中に音声コントロールを埋め込み
- [x] 音声出力デバイス選択

