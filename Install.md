# MSOfficeSVN installation guide

## Preparation

1. Install TortoiseSVN: http://tortoisesvn.net/downloads
2. Download the latest MSOfficeSVN package: https://github.com/msofficesvn/msofficesvn/releases/
3. Unzip the archive to get the files of Word add-in and Excel add-in.

The resulting folder tree is the following:
```
+-2007-2010
|  +-excel
|  | excelsvn.ini
|  | excelsvn.xlam
|  |
|  +-pptsvn
|  |  pptsvn.ini
|  |  pptsvn.ppam
|  |
|  +-word
|    wordsvn.dotm
|    wordsvn.ini
|
+-97-2003
    +-excel
    | excelsvn.ini
    | excelsvn.xla
    |
    +-pptsvn
    |  pptsvn.ini
    |  pptsvn.ppa
    |
    +-word
      wordsvn.dot
      wordsvn.ini
```

Add-in files under the "2007-2010" folder has Office 2007 ribbon interface and work with Office 2007 and newer. Obviously they won't work with older Office versions.\
Add-in files under "97-2003" folder has legacy menu and toolbar interface and work with Office 97 to Office 2003.\
Choose the right add-ins based on the MS Office version installed on your computer.


## Office 2007 and newer

### Word

INSTALLATION:

1. Copy the files `wordsvn.ini` and `wordsvn.dotm` to the following folder: `%APPDATA%\Microsoft\Word\STARTUP\`
2. Start Word.
3. The `[Subversion]` ribbon should appear.

UNINSTALLATION:

1. Start Word and go to `[File]/[Options]/[Add-Ins]` (or `[Office button]/[Word Options]/[Add-Ins]` depending on your version)
2. Then select "Word Add-Ins" from the "Manage" drop-down list and click "Go". The Add-ins dialog box appears.
3. In the list, **uncheck** the "wordsvn.dotm" check box and close.
4. Delete the files that you have copied during the installation step 1.


### Excel

INSTALLATION:

1. Copy the files `excelsvn.ini` and `excelsvn.xlam` to the following folder: `%APPDATA%\Microsoft\AddIns\`
2. Start Excel and go to `[File]/[Options]/[Add-Ins]` (or `[Office button]/[Excel Options]/[Add-Ins]` depending on your version)
3. Then select "Excel Add-Ins" from the "Manage" drop-down list and click "Go". The Add-ins dialog box appears.
4. In the list, check the "Excelsvn" check box and close.
5. The `[Subversion]` ribbon should appear.

UNINSTALLATION:

1. Do the above steps 2 and 3.
2. In the list, **uncheck** the "Excelsvn" check box and close.
3. Delete the files that you have copied during the installation step 1.


### PowerPoint

INSTALLATION:

1. Copy the files `pptsvn.ini` and `pptsvn.ppa` to the following folder: `%APPDATA%\Microsoft\AddIns\`
2. Start PowerPoint and go to `[File]/[Options]/[Add-Ins]` (or `[Office button]/[PowerPoint Options]/[Add-Ins]` depending on your version)
3. Then select "PowerPoint Add-Ins" from the "Manage" drop-down list and click "Go". The Add-ins dialog box appears.
4. In the list, check the "pptsvn" check box and close.
5. The `[Subversion]` ribbon should appear.

UNINSTALLATION:

1. Do the above steps 2 and 3.
2. In the list, **uncheck** the "pptsvn" check box and close.
3. Delete the files that you have copied during the installation step 1.


### Customization

For any tweaking, please check the [customization documentation page](https://github.com/msofficesvn/msofficesvn/blob/master/Settings.md).


## Office 97 SR2 to Office 2003 (legacy documentation)

### Word

INSTALLATION:

1. Copy the files of Word add-in to the following folders.

  * wordsvn.dot
  * wordsvn.ini

| Office97 SR2 | C:\Program Files\Microsoft Office\Office\STARTUP\ |
|:-------------|:--------------------------------------------------|
| Office2000 | C:\Program Files\Microsoft Office\Office\STARTUP\ |
| Office XP | C:\Program Files\Microsoft Office\Office10\STARTUP\ |
| Office2003 | %APPDATA%\Microsoft\Word\STARTUP\ |

2. Start Word.

3. [Subversion](Subversion.md) menu and the command bar appear.

UNINSTALLATION:

1. Start Word and select `[Tool]/[Add-Ins...]` menu item of main menu. Templates and Add-Ins dialog box appears.
2. Clear the "wordsvn.dot" check box in `[Global templates and add-ins]` list box.
3. `[Subversion]` menu and the command bar disappear.
4. Delete the files of Word add-in that you copied in installation.
5. If Subversion menu and command bar still remain after uninstallation, exit Word and delete Normal.dot and start Word again. New Normal.dot will be created and the remaining menu and command bar will disappear.


### Excel

INSTALLATION:

1. Copy files of Excel add-in to the following folders.

  * excelsvn.ini
  * excelsvn.xla

| Office97 SR2 | C:\Program Files\Microsoft Office\Office\Library\ |
|:-------------|:--------------------------------------------------|
| Office2000 | C:\Program Files\Microsoft Office\Office\Library\ |
| Office XP | C:\Program Files\Microsoft Office\Office10\Library\ |
| Office2003 | %APPDATA%\Microsoft\AddIns\ |

%APPDATA% means the "Application Data folder" of the login user.

For example, when the login user is "koki", %APPDATA% points the following folder.

```
C:\Documents and Settings\koki\Application Data
```

If you input the %APPDATA% to the address bar of Explorer, it displays "Application Data folder".

2. Start Excel and check Excelsvn check box in `[Tool]/[Add-Ins...]` of main menu.

3. `[Subversion]` menu and the command bar appear.

UNINSTALLATION:

1. Start Excel and select `[Tool]/[Add-Ins...]` menu item of main menu. Add-Ins dialog box appears.
2. Clear the "Excelsvn" check box in `[Add-Ins available]` list box.
3. `[Subversion]` menu and the command bar disappear.
4. Delete the files of Excel add-in that you copied in installation.


### PowerPoint

INSTALLATION:

1. Copy files of PowerPoint add-in to the folders same as Excel add-in.

  * pptsvn.ppa
  * pptsvn.ini

2. Start PowerPoint.

3. Remember the current setting of `[Tool]/[Option]/[Security]/[Macro Security]/[Security Level]`, and set security level to "Low".

4. Select pptsvn.ppa in `[Tool]/[Add-Ins...]/[Add]`. Then, checked pptsvn item is displayed on the list in Addin dialog box. And, Subversion menu and tool bar are displayed on main menu.

5. Go back to `[Tool]/[Option]/[Security]/[Macro Security]/[Security Level]`, and set the security level back to the original setting.

NOTE: It's dangerous to leave the security level low. Make sure it back to the original level.

UNINSTALLATION:

1. Start PowerPoint and select `[Tool]/[Add-Ins...]` menu item of main menu. Add-Ins dialog box appears.
2. Clear the "pptsvn" check box in `[Add-Ins available]` list box.
3. `[Subversion]` menu and the command bar disappear.
4. Delete the files of Excel add-in that you copied in installation.


### Customization

For any tweaking, please check the [customization documentation page](https://github.com/msofficesvn/msofficesvn/blob/master/Settings.md).
