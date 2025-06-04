# YouTube チャンネル統計 Excel レポート作成ツール(作成途中)

このプロジェクトは、指定されたYouTubeチャンネルの統計情報を収集し、それを元に月ごとのレポートをExcelファイルに出力するツールです。
各チャンネルの登録者数、動画の視聴数、長編動画やショート動画の情報を集計し、視覚的にわかりやすくまとめることができます。

## 機能

- YouTube Data APIを使用して、指定されたチャンネルから動画の統計情報を取得。
- 動画の公開日時や再生回数に基づいて、月ごとの統計を集計。
- 集計した情報をExcelファイルに書き込み、視覚的にわかりやすい形式で出力。
- 前月比などの追加指標を計算し、Excelに反映。

### 依存ライブラリ
以下のPythonパッケージをインストールしてください。

```bash
pip install pandas openpyxl google-api-python-client isodate
```

### 設定ファイル (settings.ini)

```ini
[entity]
GCP_APIKEY=your_google_api_key_here
CHANNEL_IDS=channel_id_1,channel_id_2,channel_id_3
```
・GCP_APIKEY: Google Cloud Consoleで取得したYouTube Data APIのAPIキーを入力してください。
・CHANNEL_IDS: カンマ区切りで対象のYouTubeチャンネルIDを入力します。複数指定可能です。

### APIキーの取得方法
1.Google Cloud Consoleにログインし、新しいプロジェクトを作成します。
2.「APIとサービス」→「ライブラリ」から「YouTube Data API v3」を有効化します。
3.「APIとサービス」→「認証情報」からAPIキーを作成し、settings.iniに記載します。

### 使用方法
1.依存パッケージをまとめてインストール

```bash
pip install -r requirements.txt
```

2.settings.iniを編集し、APIキーとチャンネルIDを設定します。
3.スクリプトを実行します。

```bash
python monthly.py
```

4.実行後、monthly_channel_statistics.xlsx というExcelファイルが作成され、月ごとのレポートが出力されます。

### 出力内容
生成されるExcelファイルには、以下の項目が含まれます：
・月ごとの統計
・登録者数
・獲得登録者数（前月比）
・長編動画数
・ショート動画の視聴者数
・動画の視聴数（総再生回数）
・各項目の前月比（増加分がプラスの場合は青、マイナスの場合は赤で強調表示）

### 注意点
・settings.iniのAPIキーやチャンネルIDに誤りがあるとエラーが発生する場合があります。
・YouTube APIの制限やネットワーク状況により、処理に時間がかかることがあります。





