# C++ Builderのテスト

## 必要なもの
- Microsoft Excel 20xx
- C++ Builder 12 CE:　ツール＞オプション＞エディタ＞デフォルトのファイルエンコード を UTF8 にしています
- Git for Windows のインストール、GitHubのアカウントおよびアクセストークン(repoだけでOK)作成 があるとなおよい
- アクセストークンではなくSSH登録でももちろんOK

## 概要
C++ Builderで、OLEでVBAの戻り値取得はどうやるのー？という質問への回答

## 準備(1) (GitやGitHubを使っていない場合)
- GitHubから[ZIPファイル](https://github.com/elftune/cppb_sample_1/archive/refs/heads/main.zip)をダウンロードして適当なフォルダに展開

## 準備(2) (Git/GitHub/AccessTokenの準備ができている場合)
- C++ Builderのツール＞オプション＞バージョン管理＞Git からGitを使用可能にしてあれば、ファイル＞バージョン管理リポジトリから開く から https://github.com/elftune/cppb_sample_1.git を指定してクローンすればOK

## 使い方
- 上記準備(1)(2)のいずれかによりローカルにファイルを展開しておく
- ProjectGroup1.groupproj をダブルクリックしてC++ Builderを起動
- F9を押して実行。Excel VBAというボタンを押してしばらく待って メッセージ と表示されればOK
- エラーで止まったり、デバッグで途中で止めたりするとExcel本体が宙ぶらりんになってしまう(.xlsmファイルを開いた状態だとそのファイルも編集できない状態になる)のでタスクマネージャーでその都度Excelを強制終了してください
