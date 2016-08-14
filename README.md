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
というのもウインドウリストを表示して探すというのが嫌い（時間がかかるし、Windows7あたりからフルパス表示できなくなったので探しにくい）で１回使ったウインドウは２度と使わないことが多いです。  
そのため、エクスプローラのウインドウリストがスクロールするくらい開くことになります。  
メモリ不足でエクスプローラが開けなくなったときに古いの消していくのですが、前々からこれを自動化したいと思っていました。  

## 機能
エクスプローラに関連した以下の機能があります。

* 重複したパスのウインドウを自動クローズ。（ピン留めすると自動クローズを抑止できます）
* ピン留めをすることで常にそのパスを開いたウインドウをキープ。
* お気に入りをエクスプローラで一括オープン。（すでにオープン済みのものは開きません）
* 開いているウインドウリスト(ロケーション)を表示。（最終更新日時、ロケーションパス表示）
* クローズドリスト（お気に入りと閉じたパスの履歴）に表示切替。
* ロケーションがネットワークドライブの場合、UNC(Universal Naming Convention)パスに変換するためメールとかにすぐ貼り付けれます。
* ウインドウリストでフルパスをクリップボードにコピー。
* 使用頻度が低いウインドウを自動でクローズ。（デフォルト有効）
* クローズしたウインドウ数を通知。(デフォルト有効)
* アクセントカラーを２３種類から変更が可能。（デフォルトCobalt）
* 閉じたフォルダの履歴などのショートカットメニューを表示（デスクトップのなにもないところを左ダブルクリックで出現）
* クリップボード履歴を表示可能（右ダブルクリックで出現）

![image](https://cloud.githubusercontent.com/assets/3516444/10712006/cfd43dd6-7ac7-11e5-9084-d44c6a1656d7.png)

## 設定
ExplorerWindowCleanerConfig.jsonをテキストエディタで編集することで設定を変更できます。  

* Interval  
　監視間隔です。デフォルトは10secです。
* IsAutoCloseUnused  
　未使用のウインドウをクローズするかどうかを設定します。デフォルトは有効です。  
　これは起動後のデフォルト値になり、起動後にコンテキストメニューから変更可能です。  
* ExpireInterval  
　有効期間です。ウインドウのパスが更新されてから有効期間を経過したウインドウをクローズ対象にします。  
　IsAutoCloseUnusedが有効の場合のみ使用します。デフォルトは5hourです。  
* IsNotifyCloseWindow  
　ウインドウをクローズしたときにバルーンで通知します。デフォルトは有効です。
* AccentColor  
　アクセントカラーです。"Red", "Green", "Blue", "Purple", "Orange", "Lime", "Emerald", "Teal", "Cyan", "Cobalt", "Indigo", "Violet", "Pink", "Magenta", "Crimson", "Amber", "Yellow", "Brown", "Olive", "Steel", "Mauve", "Taupe", "Sienna"から選択してください。デフォルトは"Cobalt"です。
* ExportLimitNum  
　エクスポートする履歴(お気に入り含む)の数です。デフォルトは30です。
* IsKeepPin  
　ピン留めしたパスをキープするかどうかです。（ピン留め後にパスが変更となった場合、ピン留め時のパスでウインドウを開きます）デフォルトは有効です。

## SpecialThanks

##### UI Framework : MahApps.Metro http://mahapps.com/
##### Json Library : Json.NET http://www.newtonsoft.com/json
##### Icon         : David Vignoni http://www.icon-king.com/ , Turbomilk http://turbomilk.com/ , Momentum http://www.momentumdesignlab.com/ (CC BY-SA 3.0 US) , Yusuke Kamiyamane http://p.yusukekamiyamane.com/ (CC BY 3.0) , Material Design Icons https://materialdesignicons.com/

## TODO
10/25 すべて実装しました。今後はIssuesで管理します。ご要望がありましたら追加ください。

* ~~更新しなくなるときがある。→原因調査中。。~~
* ~~ピン留めのパスが変更されたとき、ピン留めパスのウインドウを開く。~~
* ~~ライブラリのドキュメントとかも履歴から開けるようにする。（特殊フォルダ対応）~~
* ~~クローズボタン追加。~~
* ~~NowWindowからも開けるようにする。~~
* ~~右クリックでお気に入り追加~~
