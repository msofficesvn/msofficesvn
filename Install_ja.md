[English Page](Install.md)

# インストール方法 #

## 準備 ##

1. すでに旧バージョンのmsofficesvnをインストールしている場合は、アンインストールしてください。
アンインストール方法は、以下をご参照ください。

[インストール/アンインストール方法(ver.1.00)](Install_100_ja.md)

[インストール/アンインストール方法(ver.1.1.x)](Install_11x_ja.md)

[インストール/アンインストール方法(ver.1.2.x)](Install_12x_ja.md)

2. まず、TortoiseSVNをインストールしてください。

http://tortoisesvn.net/downloads

3. 本サイトのDownloadsページより、最新バージョンの `msofficesvn_<version>_ja.zip` ファイルをダウンロードし、解凍してください。以下のファイルが得られます。

```
├─2007
│  ├─excel
│  │      excelsvn.ini
│  │      excelsvn.xlam
│  │
│  ├─pptsvn
│  │      pptsvn.ini
│  │      pptsvn.ppam
│  │
│  └─word
│          wordsvn.dotm
│          wordsvn.ini
│
└─97-2003
    ├─excel
    │      excelsvn.ini
    │      excelsvn.xla
    │
    ├─pptsvn
    │      pptsvn.ini
    │      pptsvn.ppa
    │
    └─word
            wordsvn.dot
            wordsvn.ini
```

## Office97SR2～Office2003の場合 ##

### Excel ###

1. Excelアドイン用ファイルを以下のフォルダへコピーしてください。

  * excelsvn.xla
  * excelsvn.ini

| Office97 SR2 | C:\Program Files\Microsoft Office\Office\Library\ |
|:-------------|:--------------------------------------------------|
| Office2000 | C:\Program Files\Microsoft Office\Office\Library\ |
| OfficeXP | C:\Program Files\Microsoft Office\Office10\Library\ |
| Office2003 | %APPDATA%\Microsoft\AddIns\ |

%APPDATA%は、ログインユーザのApplication Dataフォルダを指します。

例えば、ログイン名が"koki"の場合は、%APPDATA%は以下のようになります。
```
C:\Documents and Settings\koki\Application Data
```
%APPDATA%をエクスプローラのアドレスバーに入力することでApplication Dataフォルダを表示することができます。

2. Excelを起動し、`[ツール]/[アドイン...]`メニューより表示される画面で、Excelsvnチェックボックスをオンにしてください。

3. `[Subversion]`メニューとコマンドバーが表示されます。

### Word ###

1. Wordアドイン用ファイルを以下のフォルダへコピーしてください。

  * wordsvn.dot
  * wordsvn.ini

| Office97 SR2 | C:\Program Files\Microsoft Office\Office\STARTUP\ |
|:-------------|:--------------------------------------------------|
| Office2000 | C:\Program Files\Microsoft Office\Office\STARTUP\ |
| OfficeXP | C:\Program Files\Microsoft Office\Office10\STARTUP\ |
| Office2003 | %APPDATA%\Microsoft\Word\STARTUP\ |

2. Wordを起動してください。

3. `[Subversion]`メニューとコマンドバーが表示されます。

### Power Point ###

以下はOffice2003でのインストール手順。他のバージョンでも同じかどうかは不明(特にセキュリティ設定周り)。

1. Excelアドインと同じフォルダへ以下のファイルをコピーしてください。

  * pptsvn.ppa
  * pptsvn.ini

2. Power Pointを起動してください。

3. `[ツール]/[オプション]/[セキュリティ]/[マクロセキュリティ]/[セキュリティレベル]`でセキュリティレベルの元の設定をメモしてから、"低"に設定してください。この設定をしないとアドインを追加できません。

4. `[ツール]/[アドイン]/[新規追加]`でコピーしたpptsvn.ppaを選択してください。アドイン画面の一覧に"pptsvn"が表示され、チェックマークが付いていることを確認してください。また、メインメニューに`[Subversion]`メニューと、`[Subversion]`ツールバーが表示されていることを確認してください。

5. 再度、`[ツール]/[オプション]/[セキュリティ]/[マクロセキュリティ]/[セキュリティレベル]`でセキュリティレベルを元の設定へ戻してください。

> 注記: セキュリティ設定を"低"に設定したままでは危険です。必ず、元のセキュリティレベルへ戻しましょう。


## Office2007, 2010の場合 ##

### Excel ###

1. Excelアドイン用ファイルを以下のフォルダへコピーしてください。

  * excelsvn.xlam
  * excelsvn.ini

```
%APPDATA%\Microsoft\AddIns\
```

2. Excelを起動し、`[Officeボタン]/[Excelのオプション]/[アドイン]/[管理]`で"Excelアドイン"を選択して、`[設定]`ボタンをクリックしてください。アドイン画面が表示されます。

3. アドイン画面で、Excelsvnチェックボックスが表示されますので、オンにしてください。

4. `[Subversion]`リボンが表示されます。

### Word ###

1. Wordアドイン用ファイルを以下のフォルダへコピーしてください。

  * wordsvn.dotm
  * wordsvn.ini

```
%APPDATA%\Microsoft\Word\STARTUP\
```

2. Wordを起動してください。

3. `[Subversion]`リボンが表示されます。


これでmsofficesvnを使用することができますが、よりカスタマイズをしたい場合は、[カスタマイズ方法](CustomSetting_ja.md)をご参照ください。

### Power Point ###

1. Excelアドインと同じフォルダへ以下のファイルをコピーしてください。

  * pptsvn.ppam
  * pptsvn.ini

2. Power Pointを起動してください。

3. ツール/アドイン/新規追加でコピーしたpptsvn.ppamを選択してください。アドイン画面の一覧に"pptsvn"が表示され、チェックマークが付いていることを確認してください。また、メインメニューに`[Subversion]`メニューと、`[Subversion]`ツールバーが表示されていることを確認してください。

# アンインストール方法 #

## Office97SR2～Office2003の場合 ##

### Excel ###

1. Excelを起動し、`[ツール]/[アドイン...]`メニューより表示される画面で、Excelsvnチェックボックスをオフにしてください。

2. `[Subversion]`メニューとコマンドバーが削除されます。

3. インストール時にコピーした、Excelアドイン用ファイルを削除してください。

### Word ###

1. Wordを起動し、`[ツール]/[テンプレートとアドイン...]`メニューより表示される画面で、wordsvn.dotチェックボックスをオフにしてください。

2. `[Subversion]`メニューとコマンドバーが削除されます。

3. インストール時にコピーした、Wordアドイン用ファイルを削除してください。

4. アンインストール後も、Subversionメニューやツールバーが残ってしまう場合は、Wordを終了しNormal.dotを削除し、再度Wordを起動してください。Normal.dotが新規作成され、残っていたSubversionメニューやツールバーが表示されなくなります。

### Power Point ###

1. Power Pointを起動し、`[ツール]/[アドイン...]`メニューより表示される画面で、pptsvnチェックボックスをオフにしてください。

2. `[Subversion]`メニューとコマンドバーが削除されます。

3. インストール時にコピーした、Excelアドイン用ファイルを削除してください。

## Office2007, 2010の場合 ##

### Excel ###

1. Excelを起動し、`[Officeボタン]/[Excelのオプション]/[アドイン]/[管理]`で"Excelアドイン"を選択して、`[設定]`ボタンをクリックしてください。アドイン画面が表示されます。

2. アドイン画面で、Excelsvnチェックボックスが表示されますので、オフにしてください。

3. `[Subversion]`リボンが削除されます。

4. インストール時にコピーした、Excelアドイン用ファイルを削除してください。

### Word ###

1. Wordを起動し、`[Officeボタン]/[Wordのオプション]/[アドイン]/[管理]`で"Wordアドイン"を選択して、`[設定]`ボタンをクリックしてください。アドイン画面が表示されます。

2. アドイン画面で、Wordsvnチェックボックスが表示されますので、オフにしてください。

3. `[Subversion]`リボンが削除されます。

4. インストール時にコピーした、Wordアドイン用ファイルを削除してください。

### Power Point ###

1. Power Pointを起動し、`[ツール]/[アドイン...]`メニューより表示される画面で、pptsvnチェックボックスをオフにしてください。

2. `[Subversion]`メニューとコマンドバーが削除されます。

3. インストール時にコピーした、PowerPoint アドイン用ファイルを削除してください。
