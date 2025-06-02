# YouTube チャンネル統計 Excel レポート作成ツール(作成途中)

このプロジェクトは、指定されたYouTubeチャンネルの統計情報を収集し、それを元に月ごとのレポートをExcelファイルに出力するツールです。各チャンネルの登録者数、動画の視聴数、長編動画やショート動画の情報を集計し、視覚的にわかりやすくまとめることができます。

## 機能

- YouTube Data APIを使用して、指定されたチャンネルから動画の統計情報を取得。
- 動画の公開日時や再生回数に基づいて、月ごとの統計を集計。
- 集計した情報をExcelファイルに書き込み、視覚的にわかりやすい形式で出力。
- 前月比などの追加指標を計算し、Excelに反映。

## 必要なパッケージ

以下のパッケージが必要です:

- `pandas`
- `openpyxl`
- `google-api-python-client`
- `isodate`
- `configparser`

これらは`requirements.txt`に記載されていますので、以下のコマンドでインストールできます。

```bash
pip install -r requirements.txt
```

設定方法
1.APIキーとチャンネルIDを設定
settings.iniという設定ファイルを用意してください。
このファイルにYouTube APIキーと、統計情報を取得するチャンネルのIDを設定します。

```ini
[entity]
GCP_APIKEY=your_google_api_key_here
CHANNEL_IDS=channel_id_1,channel_id_2,channel_id_3
```

・GCP_APIKEY: Google Cloud Consoleで取得したYouTube Data APIのAPIキーを入力します。

・CHANNEL_IDS: カンマ区切りで対象のYouTubeチャンネルIDを入力します。複数のチャンネルを指定可能です。

2.APIキーの取得
YouTube Data APIを利用するためには、Google CloudのプロジェクトでAPIキーを作成する必要があります。以下の手順でAPIキーを取得してください。

・Google Cloud Consoleにログインし、新しいプロジェクトを作成します。

・「APIとサービス」→「ライブラリ」から「YouTube Data API v3」を有効化します。

・「APIとサービス」→「認証情報」からAPIキーを作成します。このキーをsettings.iniファイルに記載します。

使用方法
1.必要なパッケージをインストールします。

```bash
pip install -r requirements.txt
```

2.settings.iniを設定し、チャンネルIDを入力します。

3.以下のコマンドでスクリプトを実行します。

```bash
python youtube_statistics_report.py
```

4.プログラムが実行されると、指定されたYouTubeチャンネルから統計情報が取得され、月ごとのレポートがmonthly_channel_statistics.xlsxというExcelファイルに書き込まれます。


出力内容
生成されるExcelファイルには、以下の項目が含まれます：

・月ごとの統計
・登録者数
・獲得登録者数（前月比）
・長編動画数
・ショート動画の視聴者数
・動画の視聴数（総再生回数）
・前月比
各項目に対して前月比が自動で計算され、増加分が表示されます。増加分がプラスの場合は青、マイナスの場合は赤で強調されます。

注意点
・settings.iniファイルに誤ったAPIキーやチャンネルIDを設定した場合、エラーが発生することがあります。
・YouTubeのAPI制限やネットワーク状況によっては、取得に時間がかかることがあります。






