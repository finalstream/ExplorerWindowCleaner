# ExplorerWindowCleaner

[![Build status](https://ci.appveyor.com/api/projects/status/tiy31lavkila6ncy?svg=true)](https://ci.appveyor.com/project/finalstream/explorerwindowcleaner)　[![GitHub release](https://img.shields.io/github/release/finalstream/ExplorerWindowCleaner.svg)](https://github.com/finalstream/ExplorerWindowCleaner/releases/latest)　[![GitHub license](https://img.shields.io/github/license/finalstream/ExplorerWindowCleaner.svg)](https://github.com/finalstream/ExplorerWindowCleaner/blob/master/LICENSE)

エクスプローラで開いたウインドウが重複した場合に自動でクローズしたりするツールです。  
起動後、タスクトレイに常駐してエクスプローラで開いたウインドウで重複したパスが存在した場合、古いほうのウインドウを閉じます。そのほかにもエクスプローラ専用ランチャー的な機能がいろいろあります。
サクッと作りたかったのでMVVMでは作っていません。

Support Windows 10 / 8 / 7  
Require .NET Framework 4.5

### [Download](https://github.com/finalstream/ExplorerWindowCleaner/releases/latest)

## 作った経緯
僕はエクスプローラを開きすぎる癖があります。  
というのもウインドウリストを表示して探すというのが嫌い（そっちのほうが時間がかかるので）で１回使ったウインドウは２度と使わないことが多いです。  
そのため、エクスプローラのウインドウリストがスクロールするくらい開くことになります。  
メモリ不足でエクスプローラが開けなくなったときに古いの消していくのですが、前々からこれを自動化したいと思っていました。  

## 機能
エクスプローラに関連した以下の機能があります。

* 重複したパスのウインドウを自動クローズ。（ピン留めすると自動クローズを抑止できます）
* ピン留めをすることでお気に入り登録。
* お気に入りをエクスプローラで一括オープン。（すでにオープン済みのものは開きません）
* 開いているウインドウリスト(ロケーション)を表示。（最終更新日時、ロケーションパス表示）
* クローズドリスト（お気に入りと閉じたパスの履歴）に表示切替。
* ロケーションがネットワークドライブの場合、UNC(Universal Naming Convention)パスに変換するためメールとかにすぐ貼り付けれます。
* ウインドウリストでフルパスをクリップボードにコピー。
* 使用頻度が低いウインドウを自動でクローズ。（デフォルト有効）
* クローズしたウインドウ数を通知。(デフォルト有効)
* アクセントカラーを２３種類から変更が可能。（デフォルトCobalt）

![image](https://cloud.githubusercontent.com/assets/3516444/10563872/78302602-75d7-11e5-8eed-6d4cd8072ae2.png)

## 設定
ExplorerWindowCleaner.exe.configをテキストエディタで編集することで設定を変更できます。  

* Interval  
　監視間隔です。デフォルトは10secです。
* IsAutoCloseUnused  
　未使用のウインドウをクローズするかどうかを設定します。デフォルトは有効です。  
　これは起動後のデフォルト値になり、起動後にコンテキストメニューから変更可能です。  
* ExpireInterval  
　有効期間です。ウインドウのパスが更新されてから有効期間を経過したウインドウをクローズ対象にします。  
　IsAutoCloseUnusedが有効の場合のみ使用します。デフォルトは8hourです。  
* IsNotifyCloseWindow  
　ウインドウをクローズしたときにバルーンで通知します。デフォルトは有効です。
* AccentColor  
　アクセントカラーです。"Red", "Green", "Blue", "Purple", "Orange", "Lime", "Emerald", "Teal", "Cyan", "Cobalt", "Indigo", "Violet", "Pink", "Magenta", "Crimson", "Amber", "Yellow", "Brown", "Olive", "Steel", "Mauve", "Taupe", "Sienna"から選択してください。デフォルトは"Cobalt"です。
* ExportLimitNum  
　エクスポートする履歴の数です。デフォルトは30です。

## SpecialThanks

##### UI Framework : MahApps.Metro http://mahapps.com/
##### Json Library : Json.NET http://www.newtonsoft.com/json
##### Icon         : David Vignoni http://www.icon-king.com/ , Turbomilk http://turbomilk.com/

## TODO
10/18 すべて実装しました。

* ~~変更があるときだけ書き込む（現状、監視間隔ごとにwriteしているので効率悪し）~~
* ~~favoriteの登録が重複する？~~
* ~~下部にフルパス表示する？（同じフォルダ名を複数開くとわからなくなるので）~~
* ~~右クリックメニューのラベル修正~~
* ~~アクティブなウインドウは閉じないようにする。（急に閉じられるとびっくりする。できるか？）~~
* ~~フォルダ開くときルートで開いているのを変更する。（コマンドのオプションを設定で変更できるようにする？）~~
* ~~通知バルーンに閉じたフォルダのNAMEを表示する。~~
* ~~ソートのサイクルになしをいれる。~~
* ~~クローズドリストがソートされていない気がする。~~
* ~~たまに落ちる。（落ちないように処理する。）~~

次の実装予定。11月上旬までに実装予定。

* ピン留めのパスが変更されたとき、ピン留めパスのウインドウを開く。
* ライブラリのドキュメントとかも履歴から開けるようにする。
* クローズボタン追加。
* NowWindowからも開けるようにする。
