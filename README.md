Microsoft Office (Excel, Word, PowerPoint) add-ins that assist document version control.

You can invoke TortoiseSVN commands from tool bar or ribbon interface of Microsoft Office .

[TortoiseSVN](http://tortoisesvn.net/downloads.html) is established version control software and useful for [version control of Microsoft Office documents](http://newgeeks.blogspot.jp/2006/08/word-document-management-using-svn.html).
I recommend you to use TortoiseSVN for document management of Microsoft Office files.

**NOTE: msofficesvn ver.1.30 or later version supports TortoiseSVN ver.1.7. or later**

![http://msofficesvn.googlecode.com/svn/branches/rb-1.2.x/2007orlater/msofficesvn_common/doc/en/wd2007menu.jpg](http://msofficesvn.googlecode.com/svn/branches/rb-1.2.x/2007orlater/msofficesvn_common/doc/en/wd2007menu.jpg)

![http://msofficesvn.googlecode.com/svn/branches/rb-1.2.x/2007orlater/msofficesvn_common/doc/en/xl2007menu.jpg](http://msofficesvn.googlecode.com/svn/branches/rb-1.2.x/2007orlater/msofficesvn_common/doc/en/xl2007menu.jpg)

# Goal #

  * Integrate MS-Office in subversion
  * Make it easy that MS-Office users control the file version while they edit it.

# Specification #

  * Invoke frequently used version control commands(`*`) from MS-Office.
  * Support Short-cut key to invoke SVN commands.
  * Support seamless action of editing, saving and version control.
  * Support Office97sr2 to Office2010(32bit and maybe 64bit)

(`*`) Update, Lock, Commit, Diff, Log, Repository Browser

# Project Current Status #

Ver.1.40 is released.

Have you ever had any troubles by editing Word and Excel files under svn control without getting lock?

The latest version add-ins solve the problems. Just add svn:needs-lock property to Word and Excel files. If you are going to edit the files, add-ins ask you whether get lock or not for the files.

For detail

[English Introduction](http://code.google.com/p/msofficesvn/wiki/Introduction)

[Japanese Introduction](http://code.google.com/p/msofficesvn/wiki/Introduction_ja)

Download latest version

[English latest version](http://msofficesvn.googlecode.com/files/msofficesvn_140_en.zip)

[Japanese latest version](http://msofficesvn.googlecode.com/files/msofficesvn_140_ja.zip)

About installation

[English Instruction](http://code.google.com/p/msofficesvn/wiki/Install)

[Japanese Instruction](http://code.google.com/p/msofficesvn/wiki/Install_ja)

# History #

|Date|Description|
|:---|:----------|
|2013.01.13|Updated 1.4.0 English version dowload file. Because PowerPoint add-in menu was Japanse. Special thanks to Thomas.|
|2013.01.02|Released 1.4.0 Support 64bit MS-Office and autolock function.|
|2012.12.02|PreReleased 1.4.0 Word and Excel, not PowerPoint. Support 64bit MS-Office.|
|2012.02.06|Released 1.3.2 English version and Japanese version. Support PowerPoint.|
|2012.01.14|Released 1.3.1 English version and Japanese version. Walk round test folder problem. Special thanks to chiayung.|
|2012.01.09|Released 1.3.0 English version and Japanese version. Support TortoiseSVN 1.7 or later. Special thanks to all people who gave me advices.|
|2009.07.08|Released 1.2.0 English version and Japanese version. Support ribbon interface. Special thanks to Jeffrey and Akash.|
|2008.08.30|Released 1.1.1 English version and Japanese version. Support shortcut key and option setting.|
|2008.08.16|Released 1.1.0 Japanese version.|
|2008.01.27|Released 1.0.0 English version.|
|2008.01.21|Announced in `Subversion r8` of 2ch BBS.|
|2008.01.21|Released 1.0.0 Japanese version.|

# Contact #

If you have some comments about this project, e-mail me.

Koki Yamamoto <kokiya@gmail.com>


<a href='http://www.visualsvn.com/'>
<img src='http://www.visualsvn.com/images/VisualSVN_125x37.gif' alt='Powered by VisualSVN!' />
</a>