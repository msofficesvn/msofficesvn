# MSOfficeSVN


## What is it

MSOfficeSVN is a set of add-ins for Microsoft Office Excel, Word and PowerPoint that assists for document version control through Subversion (SVN).\
Thanks to the MSOfficeSVN package, you can now easily version-control your MS Office files right from the MS Office menu!

Main features:
* Invoke the most frequently used version control commands directly from Microsoft Office: Update, Lock, Commit, Diff, Log, and others
* Allow shortcut keys to trigger SVN commands
* Notify the user if the `svn:needs-lock` property is in use (i.e. trying to edit a read-only file)

edition | screenshot
--- | ---
PowerPoint 2016 (EN) | ![PowerPoint 2016 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/msofficesvn_powerpoint2016.png)
Word 2007 (EN) | ![Word 2007 ribbon menu, English](https://github.com/msofficesvn/msofficesvn/raw/master/doc/en/wd2007menu.jpg)
Excel 97 (JA) | ![Excel 97 ribbon menu, Japanese](https://github.com/msofficesvn/msofficesvn/raw/master/doc/ja/xl97menu.jpg)


## How to use

1. MSOfficeSVN is based on [TortoiseSVN](https://tortoisesvn.net/), that's why **you need to have previously installed [TortoiseSVN](https://tortoisesvn.net/)** for MSOfficeSVN to work!
2. MSOfficeSVN comes as a package of multiple add-ins. Follow the instructions to install the add-in(s) of your choice:  
  a. [English Instruction](https://github.com/msofficesvn/msofficesvn/wiki/Install)  
  b. [Japanese Instruction](https://github.com/msofficesvn/msofficesvn/wiki/Install_ja)

More detailed information:\
[English documentation](https://github.com/msofficesvn/msofficesvn/wiki/Home)\
[Japanese documentation](https://github.com/msofficesvn/msofficesvn/wiki/Home_ja)


## Compatibility

TortoiseSVN is a Windows-only software, so does MSOfficeSVN.\
In other words, MSOfficeSVN is **not** compatible with Mac OS versions of Microsoft Office.

_As of April 29, 2021, we decided to stop supporting MS Office 2003 and older. [Already released versions](https://github.com/msofficesvn/msofficesvn/releases/) of MSOfficeSVN may fortunately be compatible with those old versions of MS Office._

The [latest release 1.4.0](https://github.com/msofficesvn/msofficesvn/releases/tag/rel-1.4.0) is compatible:
* Up to Microsoft Office 2019 (at least)
* With 64bit versions of MS Office
* Up to TortoiseSVN 1.14.1 (at least)
* With Windows 10

Since [release 1.3.0](https://github.com/msofficesvn/msofficesvn/releases/tag/rel-1.3.0), MSOfficeSVN supports TortoiseSVN 1.7 (or later).


## How to contribute

If you want to add your contribution to our project feel free to fork the repository, commit your changes and submit a pull-request.\
You are encouraged to read the [GitHub fork guide](https://guides.github.com/activities/forking/).


## License(s) to use/share

This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.


## Credits

Koki Yamamoto (kokiya@gmail.com)

Originally based on works found on:\
http://dkiroku.com/ (http://dkiroku.com/2005-07-01-11.html)\
http://www.nekoconeko.com/ (http://www.nekoconeko.com/~nagamori/wordsvn/)

I appreciate Mr. Osamu OKANO and Mr. Kazuyuki NAGAMORI, who created the original programs and opened their source code to the public.
