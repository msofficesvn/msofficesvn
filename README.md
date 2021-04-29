# MSOfficeSVN


## What is it

MSOfficeSVN is a set of add-ins for Microsoft Office Excel, Word and PowerPoint that assists for document version control through Subversion (SVN).\
Thanks to the MSOfficeSVN package, you can now easily version-control your MS Office files right from the MS Office menu!

![Word 2007 ribbon menu](https://github.com/msofficesvn/msofficesvn/raw/master/2007orlater/msofficesvn_common/doc/en/wd2007menu.jpg)
![Excel 2007 ribbon menu](https://github.com/msofficesvn/msofficesvn/raw/master/2007orlater/msofficesvn_common/doc/en/xl2007menu.jpg)

Main features:
* Invoke the most frequently used version control commands directly from Microsoft Office: Update, Lock, Commit, Diff, Log, and others
* Allow shortcut keys to trigger SVN commands
* Notify the user if the `svn:needs-lock` property is in use (i.e. trying to edit a read-only file)


## How to use

1. MSOfficeSVN is based on [TortoiseSVN](https://tortoisesvn.net/), that's why **you need to have previously installed [TortoiseSVN](https://tortoisesvn.net/)** for MSOfficeSVN to work!
2. MSOfficeSVN comes as a package of multiple add-ins. Follow the instructions to install the add-in(s) of your choice:  
  a. [English Instruction](https://github.com/msofficesvn/msofficesvn/wiki/Install)  
  b. [Japanese Instruction](https://github.com/msofficesvn/msofficesvn/wiki/Install_ja)

More detailed information:\
[English documentation](https://github.com/msofficesvn/msofficesvn/wiki)\
[Japanese documentation](https://github.com/msofficesvn/msofficesvn/wiki/Introduction_ja)


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


## Credits
Koki Yamamoto (kokiya@gmail.com)
