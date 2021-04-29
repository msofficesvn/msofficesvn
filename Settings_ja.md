# カスタマイズ方法 #

## ショートカットキー設定 ##

### Excel ###

1. Excelを終了してください。

2. インストール時にコピーしたexcelsvn.iniの[ShortcutKey](ShortcutKey.md)セクションのShortcutKeyOnOffキーの値を1にしてください。
```
ShortcutKeyOnOff=1
```

2. 次にショートカットキーを割り当てている部分を必要に応じて編集してください。
```
Update="+^{u}"
Commit="+^{i}"
Diff="+^{d}"
RepoBrowser="+^{w}"
Log="+^{l}"
Lock="+^{k}"
Unlock="+^{n}"
Add="+^{a}"
Explorer="+^{e}"
```

"+" はShiftキー、"^" はCtrlキーを示します。また、{}でくくられたアルファベットはそのアルファベットキーを示します。 上記の例では、更新コマンド(Update)は"Shift"キー、"Ctrl"キー、"u"キーを同時に押すと実行されます。

キーの表記については [dev\_ExcelOnKey](Developers/dev_ExcelOnKey) をご参照ください。

不要なショートカットキー割り当ては削除するか、コメントアウトしてください。コメントアウトは行頭に";"を挿入してください。

> (例)更新コマンドのショートカットキー割り当てを無効にする場合
```
;Update="+^{u}"
```

3. Excelを起動してショートカットキーが機能することを確認してください。

### Word ###

1. Wordを終了してください。

2. インストール時にコピーしたwordsvn.iniの[ShortcutKey](ShortcutKey.md)セクションのShortcutKeyOnOffキーの値を1に、Registeredキーの値を0にしてください。
```
ShortcutKeyOnOff=1
Registered=0
```

3. 次にショートカットキーを割り当てている部分を必要に応じて編集してください。
```
;Shift+Ctrl+u
Update1=256
Update2=512
Update3=85

;Shift+Ctrl+i
Commit1=256
Commit2=512
Commit3=73

(以下略)
```

<コマンド名><番号>のキーに、ショートカットキーのキーコードを割り当ててください。

> (例)上記の例では、更新コマンド(Update)は"Shift"キー、"Ctrl"キー、"u"キーを同時に押すと実行されます。
| キー | キーコード |
|:-------|:----------------|
| Shift | 256 |
| Ctrl | 512 |
| u | 85 |

キーコードについては[dev\_WordKeyCode](Developers/dev_WordKeyCode.md)をご参照ください。

不要なショートカットキー割り当ては削除するか、コメントアウトしてください。コメントアウトは行頭に";"を挿入してください。

> (例)更新コマンドのショートカットキー割り当てを無効にする場合
```
;Shift+Ctrl+u
;Update1=256
;Update2=512
;Update3=85
```

4. Wordを起動してショートカットキーが機能することを確認してください。

5. ショートカットキーをさらに変更したい場合は、Wordを起動する前に、wordsvn.iniの[ShortcutKey](ShortcutKey.md)セクションのRegisteredキーの値を必ず0にしてください。Registeredキーの値はWordを起動してショートカットキーを登録すると1にセットされます。この値が0でなければ、Word起動時にショートカットキー登録は実行されません。

## ショートカットキー設定のクリア ##

### Excel ###

1. Excelを終了してください。

2. インストール時にコピーしたexcelsvn.iniの[ShortcutKey](ShortcutKey.md)セクションのShortcutKeyOnOffキーの値を0にしてください。
```
ShortcutKeyOnOff=0
```

3. Excelを起動してショートカットキーが無効となっているか、元の割り当て機能が実行されることを確認してください。

### Word ###

1. Wordを終了してください。

2. インストール時にコピーしたwordsvn.iniの[ShortcutKey](ShortcutKey.md)セクションのShortcutKeyOnOffキーの値を0に、Registeredキーの値を0にしてください。
```
ShortcutKeyOnOff=0
Registered=0
```

3. Wordを起動してショートカットキーが無効となっているか、元の割り当て機能が実行されることを確認してください。

## オプション設定 ##

### コマンド実行時のファイル保存メッセージのオン／オフ ###

#### Excel ####

1. excelsvn.iniの[Configuration](Configuration.md)セクションのDispAskSaveModMsgキーを設定してください。
| DispAskSaveModMsgの値 | 説明 |
|:------------------------|:-------|
| 0 | メッセージを表示しない |
| 1 | メッセージを表示する |

```
[Configuration]
DispAskSaveModMsg=1
```

#### Word ####

wordsvn.iniにExcelと同様に設定をしてください。

### コミットコマンド実行時のファイルの閉じる／再度開くのオン／オフ ###

#### Excel ####

1. excelsvn.iniの[Configuration](Configuration.md)セクションのCiCloseReopenFileキーを設定してください。
| CiCloseReopenFileの値 | 説明 |
|:------------------------|:-------|
| 0 | コミット実行時の前後でファイルを閉じたり開いたりしない |
| 1 | コミット実行前にファイルを閉じ、実行後にそのファイルを再度開く |
| 2 | svn:needs-lock属性が付いているファイルに対しては、コミット実行時にファイルを閉じたり開いたりするが、その属性が付いていないファイルに対しては、ファイルを閉じたり開いたりしない。 |

FileNameCharEncodingキーは、
```
CiCloseReopenFile=2
```
のときにアドインプログラムで使用します。通常は"shift-jis"のままで変更の必要はありません。

```
[Configuration]
FileNameCharEncoding="shift-jis"
CiCloseReopenFile=0
```

#### Word ####

wordsvn.iniにExcelと同様に設定をしてください。

### [更新]、[コミット]のコマンド終了時の進行ダイアログボックスの自動終了 ###

#### Excel ####

excelsvn.iniの[Configuration](Configuration.md)セクションのCiAutoCloseProgressDlgキーを設定してください。
| CiAutoCloseProgressDlgの値 | 説明 |
|:-----------------------------|:-------|
| 0 | 自動でダイアログを閉じません。 |
| 1 | エラーがなければ自動で閉じます。 |
| 2 | エラーや競合がなければ自動で閉じます。 |
| 3 | エラー、競合、マージがなければ自動で閉じます。 |
| 4 | エラー、競合、マージが手元の操作で起きなければ自動で閉じます。 |

```
[Configuration]
CiAutoCloseProgressDlg=3
```

#### Word ####

wordsvn.iniにExcelと同様に設定をしてください。

## autolock機能の設定 ##

### Excel ###

```
[ActiveContent]
AutoLock=1
```

| キー | 説明 |
|:-------|:-------|
| AutoLock | 1:autolock機能オン、0:autolock機能オフ |

### Word ###

```
[ActiveContent]
AutoLock=1
AutoLockCheckInterval=3
```

| キー | 説明 |
|:-------|:-------|
| AutoLock | 1:autolock機能オン、0:autolock機能オフ |
| AutoLockCheckInterval| autolock機能のための、ファイル状況確認の周期設定。1秒から60秒まで設定可。　|

### PowerPoint ###

autolock機能に対応していません。