# monthly-program
YouTubeチャンネル情報を集めてExcelにまとめるプログラム
■ 準備するもの

このフォルダに入っている3つのファイル：

monthly.exe（または Macなら monthly）
settings.ini
README.txt（このファイル）
settings.iniファイルを開いて編集します。

■ settings.iniの書き方

## 設定ファイルの準備

1. `settings.ini.sample`ファイルをコピーして、`settings.ini`という名前で保存します。
2. `settings.ini`ファイルを開き、以下の情報を適切な値に置き換えてください。

   ```ini
   [entity]
   GCP_APIKEY=YOUR_API_KEY_HERE
   CHANNEL_IDS=YOUR_CHANNEL_ID1,YOUR_CHANNEL_ID2

■ プログラムの使い方

monthly.exe（または Macなら monthly）をダブルクリックすると自動で動きます。

完了すると、「monthly_channel_statistics.xlsx」という名前のExcelファイルが作られます。

■ 定期実行の設定方法

Windowsの場合（タスクスケジューラ）
スタートメニューから「タスクスケジューラ」を検索して開きます。
右側の「基本タスクの作成」をクリックします。
タスク名を入力（例：「monthly実行」など）し、「次へ」をクリック。
「毎日」や「毎週」など実行頻度を選択し、「次へ」。
実行開始日時を設定し、「次へ」。
「プログラムの開始」を選び、「次へ」。
「プログラム/スクリプト」欄に、monthly.exeのフルパス（例：C:\Users\username\Documents\monthly.exe）を入力または「参照」で選択。
「次へ」を押し、内容を確認して「完了」。
これで指定した時間に自動で実行されます。

Mac/Linuxの場合（cron）
ターミナルを開きます。

以下のコマンドでcron編集画面を開きます：

```bash
crontab -e
```

以下のような行を追加します（例は毎日午前9時に実行する場合）：

```bash
0 9 * * * /path/to/monthly
```

※ /path/to/monthly は実際のプログラムのフルパスに置き換えてください。

保存して終了するとcronが設定されます。

■ 注意点

定期実行では、settings.iniが正しい場所にあり、プログラムが実行できるパスにあることを必ず確認してください。

Windowsは.exeファイル、Mac/Linuxは実行ファイル（monthly）を指定してください。
