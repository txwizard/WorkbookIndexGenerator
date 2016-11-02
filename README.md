# WorkbookIndexGenerator

================================================================================
                         WorkbookIndexGenerator ReadMe

================================================================================
2016-11-01 21:49:58
================================================================================

WorkbookIndexGenerator is a Microsoft Excel add-in that indexes the active
workbook. Usage is straightforward.

1)  Copy the add-in that is most compatible with your version of Microsoft Excel
    into the directory where you keep your add-ins. Since this type of add-in
    can load from anywhere, if you don't have a designated add-ins directory, it
    can go anywhere, though it is best if the location is somewhere besides you
    My Documents directoy. If two or more users share the computer, put it in a
    location that is accessible via the same path for all users.

2)  Install the add-in into your copy of Microsoft Excel, and enable it.

3)  Open a workbook that contains more sheets than will fit across the bottom of
    the window.

4)  Press CTRL-SHIFT-X.

That's all there is to it. The first sheet in the active workbook is a new sheet
named Index, which contains an alphabetical list of the other worksheets in the
workbook. Each name is a hyperlink, which sends you to the upper left corner of
the sheet whose name appears in the cell.

If you add new worksheets to the book, press CTRL-SHIFT-X, and the index sheet
is regenerated.

That's all you need to know in order to use the WorkbookIndexGenerator.

--------------------------
Application Security
--------------------------

The VBA projects in both versions of the add-in are signed with a digital
certificate issued by GlobalSign CA that was valid when they were signed and
time stamped. This means that the application should run without issues
regardless of your macro security settings.

--------------------------
Supported Versions
--------------------------

TThere are two add-in workbooks.

1)  WorkbookIndexGenerator.XLA is compatible with all versions of Excel from
    97 onwards. Though it is technically compatible with Excel 2007 and newer,
    you can expect slightly better performance with the newer versions of Excel
    if you use WorkbookIndexGenerator.XLAM.

2)  WorkbookIndexGenerator.XLAM is compatible with all versions of Excel from
    2007 onwards.

--------------------------
About the Other Files
--------------------------

The following table lists and describes the other files included in the package.

--------------------------------------------------------------------------------
Name                                Description
----------------------------------  --------------------------------------------
WorkbookIndexGenerator.XLSB         This binary formatted workbook contains the
                                    complete code and data of the add-in.

WorkbookIndexGenerator.XLSM         This macro enabled workbook contains the
                                    complete code and data of the add-in.

WorkbookIndexGenerator_mMacros.BAS  This standard VBA module is the code that
                                    generates the index worksheets.

WWXLAppExceptions.CLS               This class module handles run-time errors,
                                    in the extremely unlikely event that one
                                    occurs. In over seven years of use in a
                                    variety of settings, I have never had even
                                    one run-time exception.
--------------------------------------------------------------------------------

--------------------------
Internal Documentation
--------------------------

The source code includes comprehenisve technical documentation. Argument names
follow Hungarian notation, to make the type immediately evident in most cases. A
lower case "p" precedes a type prefix, to differentiate function arguments from
local variables, followed by a lower case "a" to designate arguments that are
arrays. Object variables have an initial underscore and static variables begin
with "s_"; this naming scheme makes variable scope crystal clear.
================================================================================
