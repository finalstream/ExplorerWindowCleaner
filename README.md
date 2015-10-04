# ExplorerWindowCleaner

[![Build status](https://ci.appveyor.com/api/projects/status/tiy31lavkila6ncy?svg=true)](https://ci.appveyor.com/project/finalstream/explorerwindowcleaner)　[![GitHub license](https://img.shields.io/github/license/finalstream/ExplorerWindowCleaner.svg)](https://github.com/finalstream/ExplorerWindowCleaner/blob/master/LICENSE)

エクスプローラで開いたウインドウが重複した場合に自動でクローズするツールです。  
起動後、タスクトレイに常駐してエクスプローラで開いたウインドウで重複したパスが存在した場合、古いほうのウインドウを閉じます。

Support Windows 10 / 8 / 7  
Require .NET Framework 4.5

### [Download](https://github.com/finalstream/ExplorerWindowCleaner/releases/latest)

## 機能
エクスプローラに関連した以下の機能があります。

* 重複したパスのウインドウを自動クローズ。
* 開いているウインドウリスト(フルパス)を表示。（最終更新日時、ローカルパス表示）
* ウインドウリストでフルパスをクリップボードにコピー。
* 使用頻度が低いウインドウを自動でクローズ。（デフォルト無効）
* クローズしたウインドウ数を通知。(デフォルト有効)

![image](https://cloud.githubusercontent.com/assets/3516444/10263309/735a5472-6a24-11e5-880a-dbacc7232786.png)

## 設定
ExplorerWindowCleaner.exe.configをテキストエディタで編集することで設定を変更できます。  

* Interval  
　監視間隔です。デフォルトは10secです。
* IsAutoCloseUnused  
　未使用のウインドウをクローズするかどうかを設定します。デフォルトは無効です。  
　これは起動後のデフォルト値になり、起動後にコンテキストメニューから変更可能です。  
* ExpireInterval  
　有効期間です。ウインドウのパスが更新されてから有効期間を経過したウインドウをクローズ対象にします。  
　IsAutoCloseUnusedが有効の場合のみ使用します。デフォルトは8hourです。  
* IsNotifyCloseWindow  
　ウインドウをクローズしたときにバルーンで通知します。デフォルトは有効です。
* AccentColor  
　アクセントカラーです。"Red", "Green", "Blue", "Purple", "Orange", "Lime", "Emerald", "Teal", "Cyan", "Cobalt", "Indigo", "Violet", "Pink", "Magenta", "Crimson", "Amber", "Yellow", "Brown", "Olive", "Steel", "Mauve", "Taupe", "Sienna"から選択してください。デフォルトは"Cobalt"です。

## SpecialThanks

##### UI Framework : MahApps.Metro http://mahapps.com/
##### Icon         : David Vignoni http://www.icon-king.com/
