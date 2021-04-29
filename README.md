# MSOfficeSVN


## What is it

MSOfficeSVN is a set of add-ins for Microsoft Office Excel, Word and PowerPoint that assists for document version control through Subversion (SVN).\
Thanks to the MSOfficeSVN package, you can now easily version-control your MS Office files right from the MS Office menu!

MSOfficeSVN is based on [TortoiseSVN](https://tortoisesvn.net/), an established Subversion client for Windows.

Main features:
* Invoke the most frequently used version control commands directly from Microsoft Office: Update, Lock, Commit, Diff, Log, and others
* Allow shortcut keys to trigger SVN commands ; the [key mapping](#shortcut-keys) is customizable
* Notify the user if the `svn:needs-lock` property is in use (i.e. trying to edit a read-only file)

edition | screenshot
--- | ---
PowerPoint 2016 (EN) | ![PowerPoint 2016 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/msofficesvn_powerpoint2016.png)
Word 2007 (EN) | ![Word 2007 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/wd2007menu.jpg)
Excel 97 (JA) | ![Excel 97 ribbon menu, Japanese](https://github.com/msofficesvn/msofficesvn/raw/master/doc/ja/xl97menu.jpg)


## How to use

1. MSOfficeSVN is based on [TortoiseSVN](https://tortoisesvn.net/), that's why **you need to have previously installed [TortoiseSVN](https://tortoisesvn.net/)** for MSOfficeSVN to work!
2. MSOfficeSVN comes as a package of multiple add-ins. Follow the instructions to install the add-in(s) of your choice:  
  a. [English Instructions](https://github.com/msofficesvn/msofficesvn/Install.md)  
  b. [Japanese Instructions](https://github.com/msofficesvn/msofficesvn/Install_ja.md)


## Compatibility

TortoiseSVN is a Windows-only software, so does MSOfficeSVN.\
In other words, MSOfficeSVN is **not** compatible with Mac OS versions of Microsoft Office.

_As of April 29, 2021, we decided to stop supporting MS Office 2003 and older. [Already released versions](https://github.com/msofficesvn/msofficesvn/releases/) of MSOfficeSVN may fortunately be compatible with those old versions of MS Office._

The [latest release 1.4.0](https://github.com/msofficesvn/msofficesvn/releases/tag/rel-1.4.0) is compatible:
* Up to Microsoft Office 2019 (at least)
* Down to Microsoft Office 97
* With 32bit and 64bit versions of MS Office
* Up to TortoiseSVN 1.14.1 (at least)
* Down to TortoiseSVN 1.7
* With Windows OS (32bit, 64bit) where WSH is installed

Since [release 1.3.0](https://github.com/msofficesvn/msofficesvn/releases/tag/rel-1.3.0), MSOfficeSVN supports TortoiseSVN 1.7 (or later).


## Known limitations

The add-ins apply the commands to only the active document or book, so you can't use it when you wish to edit multiple files and commit them at the same time to make a "change set".


## How to contribute

If you want to add your contribution to our project feel free to fork the repository, commit your changes and submit a pull-request.\
You are encouraged to read the [GitHub fork guide](https://guides.github.com/activities/forking/).


## License(s) to use/share

This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.


## Credits

Koki Yamamoto (kokiya@gmail.com)

We appreciate the valuable contribution of the original authors who created the initial programs and opened their source code to the public:\
Mr. Osamu OKANO (http://dkiroku.com/2005-07-01-11.html) \
Mr. Kazuyuki NAGAMORI (http://www.nekoconeko.com/~nagamori/wordsvn/)


## Functional overview

### Commands

The add-ins can apply the following commands to the active workbook and document.

* SVN Update
* SVN Get lock
* SVN Commit
* SVN Diff
* SVN Show log
* SVN Repo-browser
* SVN Release lock
* SVN Add
* SVN Delete
* Open Explorer

Regarding `Update`, `Get lock`, `Commit` and `Release lock`, the active document or book is first closed then reopen after having applied the SVN command in order to refresh the document status.

`Open Explorer` opens the Windows Explorer and selects the active document or book.


### Needs-lock detection (a.k.a. Auto-lock feature)

Binary files stored inside a Subversion repository usually get the `svn:needs-lock` property. This forces the user to take an exclusive lock on a file before editing it, thus avoiding conflicts of two or more people working on the same file at the same time.\
http://svnbook.red-bean.com/en/1.8/svn.advanced.locking.html

As a consequence, TortoiseSVN applies a Read-only file-system flag on each file that has the `svn:needs-lock` property.\
https://tortoisesvn.net/docs/release/TortoiseSVN_en/tsvn-dug-locking.html

In MSOfficeSVN, we detect the `svn:needs-lock` and offer the possibility to take the lock on the file and make it writable. This is a way to avoid making modifications on a file you won't be able to save in the end.


### Ribbon interface and Tool bar support

This set of add-ins supports the ribbon interface of Office 2007 and newer.\
For versions where Office 2003 and older is supported (no ribbon), we provide a Tool bar instead.


### Shortcut keys

Shortcut keys are assigned by default for the Subversion commands.\
You may [edit the shortcuts mapping](https://github.com/msofficesvn/msofficesvn/Settings.md) if you want.


### Option settings

The following option settings are available.

* Turn on/off displaying a message asking data file saving while performing a command.
* Turn on/off closing and reopening the data file in performing commit command.
* Turn on/off auto closing of the progress dialog box in the end of performing `Update`, `Commit`.

Instructions for those settings to be edited are given by the [customization documentation page](https://github.com/msofficesvn/msofficesvn/Settings.md).


## Screenshots of the available packages

edition | screenshot
--- | ---
Excel 2016 (EN) | ![Excel 2016 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/msofficesvn_excel2016.png)
PowerPoint 2016 (EN) | ![PowerPoint 2016 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/msofficesvn_powerpoint2016.png)
Word 2016 (EN) | ![Word 2016 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/msofficesvn_word2016.png)
Word 2007 (EN) | ![Word 2007 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/wd2007menu.jpg)
Excel 2007 (EN) | ![Excel 2007 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/xl2007menu.jpg)
Word 2007 (JA) | ![Word 2007 ribbon menu, Japanese](https://github.com/msofficesvn/msofficesvn/raw/master/doc/ja/wd2007menu.jpg)
Excel 2007 (JA) | ![Excel 2007 ribbon menu, Japanese](https://github.com/msofficesvn/msofficesvn/raw/master/doc/ja/xl2007menu.jpg)
Word 97 (JA) | ![Word 97 ribbon menu, Japanese](https://github.com/msofficesvn/msofficesvn/raw/master/doc/ja/wd97menu.jpg)
Excel 97 (JA) | ![Excel 97 ribbon menu, Japanese](https://github.com/msofficesvn/msofficesvn/raw/master/doc/ja/xl97menu.jpg)
