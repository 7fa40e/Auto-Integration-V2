# Auto-Integration-V2
LINE API, GAS, Gemini API,  Google Spreadsheetを使用した画像読み取り及びデータ書き込みシステム

レシートの画像を送信すると、Gemini APIが「日付・店舗名・金額・品目」を自動で読み取り、指定したGoogleスプレッドシートへ自動記録するLINE Bot。

複数人での共同管理（家計簿やイベントの経費精算など）を想定しており、システム全体で一意なシート管理や、LINE上での手動修正機能、画像フォルダのGoogleドライブへの整理機能を備えた。

---

## 主な機能

* **AIレシート解析**: `Gemini 2.5 Flash` を利用し、送られたレシート画像からデータを抽出。
* **一意なシート管理機能**: システム全体（全ユーザー）で同じシート名・フォルダ名の重複作成を一律ブロックし、混同を防止。
* **共有キーシステム**: 各シートに固有の共有用キーが発行され、そのキーを入力するだけで簡単にシートの共同編集が可能。
* **LINE上での手動修正機能**: AIの誤認識があっても、LINEのクイックリプライボタンとチャット入力を用いて、その場で「日付・店舗・金額・品目」を上書き修正可能。
* **画像自動仕分け**: 記録したレシート画像は、スプレッドシートごとに自動生成されるGoogleドライブの専用フォルダへ自動保存。
* **削除検知**: 管理しているスプレッドシートが削除（ゴミ箱へ移動）された場合、選択肢から自動で除外・整理。

---

## システム構成

* **ユーザーインターフェース**: LINE公式アカウント (Messaging API)
* **バックエンド**: Google Apps Script (GAS)
* **データベース/ファイル管理**: Google スプレッドシート / Google ドライブ
* **処理AI**: Gemini API (`gemini-2.5-flash`)

---

## 準備・導入方法

### 1. LINE Messaging API の準備
1. [LINE Developers](https://developers.line.biz/) にログインし、プロバイダーとチャネル（Messaging API）を新規作成します。
2. 「Messaging API設定」タブから **チャネルアクセストークン（長期）** を発行し、メモしておきます。
3. 同タブにある **LINEアプリ用QRコード** から、Botを友達追加しておきます。

### 2. Gemini API キーの取得
1. [Google AI Studio](https://aistudio.google.com/) にアクセスします。
2. **Get API key** から、新しいAPIキーを発行し、メモしておきます。

### 3. 管理帳スプレッドシートの作成
1. Googleスプレッドシートを新規作成し、名前を「レシート管理帳_マスター」などに変更します。
2. 1行目に以下のヘッダー（列名）を設定します。
   * **A列**: `ユーザーID`
   * **B列**: `シートID`
   * **C列**: `シート名`
   * **D列**: `共有キー`
   * **E列**: `フォルダID`
3. このスプレッドシートの **URLからID（`https://docs.google.com/spreadsheets/d/【この部分】/edit`）** をコピーしてメモしておきます。

### 4. Google Apps Script (GAS) の設定
1. [Google Apps Script](https://script.google.com/) にアクセスし、「新しいプロジェクト」を作成します。
2. 本リポジトリの `コード.gs` の中身をすべて貼り付けます。
3. コードの最上部にある環境変数を、先ほどメモしたものに書き換えます。
   ```javascript
   const LINE_ACCESS_TOKEN = '任意のLINEアクセストークン';
   const GEMINI_API_KEY    = '任意のGemini APIキー';
   const MASTER_SPREADSHEET_ID = '任意のマスタースプレッドシートID';
