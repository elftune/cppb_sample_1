# C++ Builderのテスト

## 必要なもの
- Microsoft Excel 20xx
- C++ Builder 12 CE:　ツール＞オプション＞エディタ＞デフォルトのファイルエンコード を UTF8 にしています
- Git for Windows のインストール、GitHubのアカウントおよびアクセストークン作成 があるとなおよい

## 概要
C++ BuilderでOLEでVBAを操作するのはどうやるのー？という質問への回答

## 使い方
- 1) GitやGitHubを使っていない場合: GitHubから[ZIPファイル](https://github.com/elftune/cppb_sample_1/archive/refs/heads/main.zip)を適当なフォルダにダウンロードして展開
- 2) GitHubのアカウントがあってアクセストークンを作ってあり、GitをインストールしてC++ Builderのツール＞オプション＞バージョン管理＞Git から登録してあれば、ファイル＞バージョン管理リポジトリから開く から https://github.com/elftune/cppb_sample_1.git を指定すればOK
- ProjectGroup1.groupproj をダブルクリックしてC++ Builderを起動
- F9を押して実行。Excel VBAというボタンを押してしばらく待って メッセージ と表示されればOK
