# MSOfficeSVN customization guide

## Shortcut key settings

### Excel

1. Exit Excel.

2. Set the value of `ShortcutKeyOnOff` key to 1 in the `[ShortcutKey]` section of `excelsvn.ini` that has been copied during installation.
```
[ShortcutKey]
ShortcutKeyOnOff=1
```

3. Edit the parts that assign shortcut keys to SVN commands depending on your needs.
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

`+` means `Shift` key and `^` means `Ctrl` key. The alphabet that is braced by {} means the alphabet key.\
In the above example, the `Update` command is triggered when the `Shift` key, `Ctrl` key and `u` keys are pressed at the same time.

Refer to [dev\_ExcelOnKey](https://github.com/msofficesvn/msofficesvn/wiki/dev_ExcelOnKey) about key notation.

4. Re-start Excel and confirm that shortcut keys work.


### Word

1. Exit Word.

2. Set the value of `ShortcutKeyOnOff` key to 1, and the value of `Registered` key to 0 in the `[ShortcutKey]` section of `wordsvn.ini` that has been copied during installation.
```
[ShortcutKey]
ShortcutKeyOnOff=1
Registered=0
```

3. Edit the parts that assign shortcut keys to SVN commands depending on your needs.
```
;Shift+Ctrl+u
Update1=256
Update2=512
Update3=85

;Shift+Ctrl+i
Commit1=256
Commit2=512
Commit3=73

...
```

Assign shortcut key codes to `<Command Name><No.>` keys.

In the above example, the `Update` command is triggered when the `Shift` key, `Ctrl` key and `u` keys are pressed at the same time.
| Key | Key code |
|:----|:---------|
| Shift | 256 |
| Ctrl | 512 |
| u | 85 |

Refer to [dev\_WordKeyCode](https://github.com/msofficesvn/msofficesvn/wiki/dev_WordKeyCode) about key codes.

4. Re-start Word and confirm that shortcut keys work.

5. If you wish to change the shortcut key assignment furthermore, make sure to set the value of `Registered` to 0 in the `[ShortcutKey]` section of `wordsvn.ini` **before you start Word**. The value of `Registered` is set to 1 when you start Word and shortcut key assignment is registered. If the value is not 0, the shortcut key assignment is considered already registered, thus not applied.


## Clear the shortcut keys assignment

### Excel

1. Exit Excel.

2. **EITHER** remove or comment out the line(s) corresponding to the shortcut you want to disable.
```
;Update="+^{u}"
```

3. **OR** set the value of `ShortcutKeyOnOff` key to 0 in the `[ShortcutKey]` section of `excelsvn.ini` that has been copied during installation.
```
[ShortcutKey]
ShortcutKeyOnOff=0
```

4. Re-start Excel and confirm that shortcut keys are disabled or that originally assigned functions are invoked.


### Word ###

1. Exit Word.

2. **EITHER** remove or comment out the line(s) corresponding to the shortcut you want to disable.
```
;Shift+Ctrl+u
;Update1=256
;Update2=512
;Update3=85
```

3. **OR** set the value of `ShortcutKeyOnOff` key to 0, and the value of `Registered` key to 0 in the `[ShortcutKey]` section of `wordsvn.ini` that has been copied during installation.
```
[ShortcutKey]
ShortcutKeyOnOff=0
Registered=0
```

4. Re-start Word and confirm that shortcut keys are disabled or that originally assigned functions are invoked.


## Needs-lock detection (Auto-lock) settings

### Excel

```
[ActiveContent]
AutoLock=1
```

| Key | Description |
|:----|:------------|
| AutoLock | 1: auto-lock ON, 0: auto-lock OFF |


### Word

```
[ActiveContent]
AutoLock=1
AutoLockCheckInterval=3
```

| Key | Description |
|:----|:------------|
| AutoLock | 1: auto-lock ON, 0: auto-lock OFF |
| AutoLockCheckInterval| File status check interval setting for auto-lock. 1 second to 60 seconds available.　|


### PowerPoint

This feature is not supported in PowerPoint.


## Turn on/off the message asking data file saving while performing a command

Equally applicable for Excel and Word: Set the value of `DispAskSaveModMsg` key in the `[Configuration]` section of `excelsvn.ini` and/or `wordsvn.ini`.

| The value of `DispAskSaveModMsg` key | Description |
|:-----------------------------------|:------------|
| 0 | Don't display a message |
| 1 | Display a message |

```
[Configuration]
DispAskSaveModMsg=1
```


## Turn on/off closing and reopening the data file in performing commit command

Equally applicable for Excel and Word: Set the value of `CiCloseReopenFile` key in the `[Configuration]` section of `excelsvn.ini` and/or `wordsvn.ini`.

| The value of `CiCloseReopenFile` key | Description |
|:-----------------------------------|:------------|
| 0 | Don't close and reopen the data file in performing commit |
| 1 | Close and reopen the data file in performing commit |
| 2 | Close and reopen the data file in performing commit, but only if it has the svn:needs-lock property |

`FileNameCharEncoding` key is used by add-in when the following value is set.
```
CiCloseReopenFile=2
```

If you wish to use a language other than "iso-8859-1" (Western European) for the data file name, try to set the character set value to `FileNameCharEncoding` key. (Sorry, I'm not sure wether they realy work or not.)

| big5 |
|:-----|
| euc-jp |
| euc-kr |
| gb2312 |
| iso-2022-jp |
| iso-2022-kr |
| iso-8859-1 |
| iso-8859-2 |
| iso-8859-3 |
| iso-8859-4 |
| iso-8859-5 |
| iso-8859-6 |
| iso-8859-7 |
| iso-8859-8 |
| iso-8859-9 |
| koi8-r |
| shift-jis |
| us-ascii |
| utf-7 |
| utf-8 |

Refer to http://en.wikipedia.org/wiki/Character_encoding about character set values.

```
[Configuration]
FileNameCharEncoding="iso-8859-1"
CiCloseReopenFile=0
```


## Turn on/off auto closing of the progress dialog box in the end of performing Update and Commit

Equally applicable for Excel and Word: Set the value of `CiAutoCloseProgressDlg` key in the `[Configuration]` section of `excelsvn.ini` and/or `wordsvn.ini`.

| The value of `CiAutoCloseProgressDlg` key | Description |
|:----------------------------------------|:------------|
| 0 | don't close the dialog automatically |
| 1 | auto close if no errors |
| 2 | auto close if no errors and conflicts |
| 3 | auto close if no errors, conflicts and merges |
| 4 | auto close if no errors, conflicts and merges for local operations |

```
[Configuration]
CiAutoCloseProgressDlg=3
```
