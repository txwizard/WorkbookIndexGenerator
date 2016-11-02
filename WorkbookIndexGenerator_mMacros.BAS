Attribute VB_Name = "mMacros"
Option Explicit

    Const DEFAULT_ERR_MSG_WKS_NM As String = "wsWWXLAppErrors"

    Const FORMULA_PREFIX As String = "="
    Const FORMULA_SHEET_NAME_DLM As String = "!"
    Const FORMULA_CELL_ADDR_DLM As String = ":"

    Const INDEX_SHEET_NAME As String = "Index"

    Const INSTR_NOT_FOUND As Integer = 0

    Const QUOTE_CHAR_DBL_WW As String = """" ' Yes, this resolves to a solitary double quote character.
    Const QUOTE_CHAR_SGL_WW As String = "'"

    Const RANGE_ORIGIN_CELL_INDEX As Long = 1

    Const RANGELIB_MSG_WKS_NM As String = "XLRangeLib_WW"

    Const SPACE_CHAR_WW As String = " "      ' Make it crystal clear that this is ONE space.


Public Sub IndexActiveWorkbook()
Attribute IndexActiveWorkbook.VB_Description = "Create or update the Index sheet of the active workbook."
Attribute IndexActiveWorkbook.VB_ProcData.VB_Invoke_Func = "X\n14"

'   ============================================================================
'
'   Name:               IndexActiveWorkbook
'
'   Synopsis:           This VBA routine simplifies managing workbooks that
'                       contain more sheets than it is practical to manage with
'                       the built-in tab controls by creating an index sheet,
'                       which is populated and given the focus.
'
'   Dependencies:       VBA class module WWXLAppExceptions, which lives in the
'                       following file, from which it was imported.
'
'                           C:\Documents and Settings\DAG\My Documents\_CLIENTS\NOLANB\01147\NEW_Code\VBA\WWXLAppExceptions.cls
'
'   Remarks:            This function began as Workbook_Open Event Procedure,
'                       running within a workbook that had too many tabls to
'                       effectively navigate using the supplied user interface.
'
'                       The final product is a combination of the material given
'                       in both of the following references.
'
'   References:         1)  "VBA: Macro to Create a Hyperlink Menu of Worksheets"
'                           http://www.automateexcel.com/2004/09/06/excel_vba_macro_to_create_a_hyperlink_me/
'
'                       2)  Samples by Excel MVP D. McRitchie
'                           http://www.mvps.org/dmcritchie/excel/code/buildtoc.txt
'
'   Author:             David A. Gray
'
'   ----------------------------------------------------------------------------
'   Revision History
'   ----------------------------------------------------------------------------
'
'   Date       Author Synopsis
'   ---------- ------ ----------------------------------------------------------
'   2010/05/03 DAG/WW Initial version
'   2010/05/04 DAG/WW Add code to sort the index, and remove the save command.
'   2012/10/12 DAG/WW Adapt for use from an add-in, to process the active book.
'   2012/10/13 DAG/WW Recover from the subscript out of bounds error that occurs
'                     when a workbook contains only the index sheet, and other
'                     sheets, if any, are hidden.
'   ============================================================================

    Const LINK_COL As Integer = 1
    Const WORKSHEET_TYPENAME As String = "Worksheet"

    Dim wsAll As Sheets
    Dim wsThis As Worksheet
    Dim wsIndex As Worksheet
    Dim wsFirst As Worksheet

    ActiveWorkbook.Activate
    Set wsAll = ActiveWorkbook.Sheets

    For Each wsThis In wsAll
        With wsThis
            If .Name = INDEX_SHEET_NAME Then
                Set wsIndex = wsThis
                Exit For    ' Leave early.
            Else
                If wsFirst Is Nothing Then
                    Set wsFirst = wsThis
                End If  ' If wsFirst Is Nothing Then
            End If  ' If .Name = INDEX_SHEET_NAME Then
        End With    ' With wsThis
    Next wsThis ' For Each wsThis In wsAll

    '   ------------------------------------------------------------------------
    '   If wsIndex is uninitialized, then wsFirst points to the frist sheet in
    '   the book, before which the index is inserted.
    '   ------------------------------------------------------------------------

    If wsIndex Is Nothing Then
        wsAll.Add Before:=wsFirst
        Set wsFirst = Nothing   ' It's done its job; dispose of its resources.
        Set wsIndex = ActiveWorkbook.ActiveSheet
        wsIndex.Name = INDEX_SHEET_NAME
    End If  ' If wsIndex Is Nothing Then

    Dim rngIndex As Range: Set rngIndex = wsIndex.Cells
    rngIndex.ClearContents

    Dim lngThisRow As Long

    For Each wsThis In wsAll
        With wsThis
            If TypeName(wsThis) = WORKSHEET_TYPENAME Then
                If IncludeSheetInIndex(wsThis) Then
                    lngThisRow = lngThisRow + 1
                    Dim rngThisEntry As Range: Set rngThisEntry = rngIndex.Cells(lngThisRow, _
                                                                                 LINK_COL)
                    Dim rngThisTarget As Range: Set rngThisTarget = wsThis.Cells(RANGE_ORIGIN_CELL_INDEX, _
                                                                                 RANGE_ORIGIN_CELL_INDEX)
                    wsIndex.Hyperlinks.Add Anchor:=rngThisEntry, _
                                           Address:=vbNullString, _
                                           SubAddress:=SheetNameForFormula(rngThisTarget.Worksheet) _
                                                                           & FORMULA_SHEET_NAME_DLM _
                                                                           & rngThisTarget.Address, _
                                           TextToDisplay:=wsThis.Name
                End If  ' If IncludeSheetInIndex(wsThis) Then
            End If  ' If TypeName(wsThis) = WORKSHEET_TYPENAME Then
        End With    ' With wsThis
    Next wsThis ' For Each wsThis In wsAll

    If lngThisRow > 0 Then
        Dim rngTheIndex As Range
        Set rngTheIndex = rngCellsFromRng_P6C(rngIndex, _
                                              RANGE_ORIGIN_CELL_INDEX, _
                                              RANGE_ORIGIN_CELL_INDEX, _
                                              lngThisRow, _
                                              RANGE_ORIGIN_CELL_INDEX)

        Dim rngSortKey1 As Range
        Set rngSortKey1 = rngCellsFromRng_P6C(rngTheIndex, _
                                              RANGE_ORIGIN_CELL_INDEX, _
                                              RANGE_ORIGIN_CELL_INDEX, _
                                              rngTheIndex.Rows.Count, _
                                              RANGE_ORIGIN_CELL_INDEX)

        rngTheIndex.Sort Key1:=rngSortKey1, _
                         Header:=xlNo, _
                         MatchCase:=True, _
                         Orientation:=xlSortColumns
    End If  ' If lngThisRow > 0 Then

    '   ------------------------------------------------------------------------
    '   Finally, save and activate the index worksheet.
    '   ------------------------------------------------------------------------

    wsIndex.Activate

End Sub


Private Function IncludeSheetInIndex(pws As Worksheet) As Boolean

'   ============================================================================
'
'   Name:               IncludeSheetInIndex
'
'   Synopsis:           This VBA routine simplifies managing workbooks that
'                       contain more sheets than it is practical to manage with
'                       the built-in tab controls by creating an index sheet,
'                       which is populated and given the focus.
'
'   In:                 pws As Worksheet    = Worksheet to evaluate
'
'   Out:                Return TRUE unless any of the following is true.
'
'                           1)  Name    = XLRangeLib_WW or wsWWXLAppErrors
'                           2)  Visible = xlSheetHidden or xlSheetVeryHidden
'
'   ============================================================================

    Dim fTestVisibility As Boolean

    With pws
        Select Case .Name
            Case INDEX_SHEET_NAME
                fTestVisibility = False
                IncludeSheetInIndex = False
            Case RANGELIB_MSG_WKS_NM
                fTestVisibility = False
                IncludeSheetInIndex = False
            Case DEFAULT_ERR_MSG_WKS_NM
                fTestVisibility = False
                IncludeSheetInIndex = False
            Case Else
                fTestVisibility = True
        End Select  ' Select Case .Name

        If fTestVisibility Then
            Select Case .Visible
                Case xlSheetVisible
                    IncludeSheetInIndex = True
                Case xlSheetHidden
                    IncludeSheetInIndex = False
                Case xlSheetVeryHidden
                    IncludeSheetInIndex = False
                Case Else
                    IncludeSheetInIndex = True
            End Select  ' Select Case .Visible
        End If  ' If fTestVisibility Then
    End With    ' With pws

End Function


Private Function rngCellsFromRng_P6C( _
        prngIn As Range, _
        plngUpperLeftRow As Long, _
        pintUpperLeftCol As Integer, _
        plngLowerRightRow As Long, _
        pintLowerRightCol As Integer) _
            As Range

'   ----------------------------------------------------------------------------
'
'    Name:         rngCellsFromRng_P6C
'
'    Synopsis:     Given a contiguous range and the upper left and lower right
'                  corners of a contiguous block of cells within it, return that
'                  block of cells as a new range object.
'
'    Arguments:    prngIn                 = Range object (as opposed to a Name).
'
'                  plngUpperLeftRow       = Long positive integer containing the
'                                           row number of the cell in the upper
'                                           left corner of the output range.
'
'                                           Please see Note 4 about the upper
'                                           limit of this value, which is enfor-
'                                           ced.
'
'                  pintUpperLeftCol       = Positive Integer containing the
'                                           column number of the cell in the
'                                           upper left corner of the output
'                                           range.
'
'                                           Please see Note 4 about the upper
'                                           limit of this value, which is enfor-
'                                           ced.
'
'                  plngLowerRightRow      = Long positive integer containing the
'                                           row number of the cell in the lower
'                                           right corner of the output range.
'
'                                           Please see Note 4 about the upper
'                                           limit of this value, which is enfor-
'                                           ced.
'
'                  pintLowerRightCol      = Positive Integer containing the
'                                           column number of the  cell in the
'                                           lower right output range.
'
'                                           Please see Note 4 about the upper
'                                           limit of this value, which is enfor-
'                                           ced.
'
'   Returns:        A reference to a Range object containing the desired subset
'                   of cells.
'
'                   Returns Nothing if any argument is invalid.
'                   Please see Note 4.
'
'   Notes:          1) Although a user defined function is creating the output
'                      range, the correct syntax reamins "set rng = ..." as it
'                      would with any other statement that creates a reference
'                      to an object.
'
'                   2) The output range must be a subset of the cells in the
'                      input range.
'
'                   3) The input range must be contiguous. Ranges that contain
'                      two or more Areas are unsupported.
'
'                   4) Although the document cited below, "Excel Specifications
'                      and Limits," states that it applies to Microsoft Office
'                      Excel 2003, since the limits in question affect the file
'                      format, I think it's safe to say that they apply to all
'                      versions of Excel that share the same format. This means
'                      that they apply to the following versions.
'
'                           Excel 97
'                           Excel 2000
'                           Excel 2002 (Excel XP)
'                           Excel 2003
'
'                      Since a contiguous range cannot span worksheets, it fol-
'                      lows that these limits also apply to any contiguous Range
'                      object, and to any Cells collection.
'
'                      Functions ValidColIndex_WW and ValidRowIndex_WW are
'                      called to perform preliminary validation of column
'                      indices and row indices, respectively.
'
'   References:     Excel Specifications and Limits
'                   http://office.microsoft.com/en-us/excel/HP051992911033.aspx
'
'   Date Written:   11-12 June 2003.
'
'   Author:         David A. Gray, Simple Soft Services, Inc., d/b/a WizardWrx.
'
'   Copyright:      (C) 2003-2008, Simple Soft Services, Inc., d/b/a WizardWrx.
'                   Irving, Texas, USA
'                   http://www.wizardwrx.com/
'                   All rights reserved world wide.
'
'   ----------------------------------------------------------------------------
'   Revision History
'   ----------------------------------------------------------------------------
'
'   Date       Version Author Synopsis
'   ---------- ------- ------ --------------------------------------------------
'   2004/11/22         DAG    Rewrite error handler so that it uses no global
'                             variables.
'
'   2008/08/17         DAG    1) Change arguments pintUpperLeftCol and
'                                pintLowerRightCol from Long to Integer, to
'                                remind users that a worksheet may contain only
'                                256 columns.
'
'                             2) Add code to validate all four agruments against
'                                the above limits.
'
'   2008/08/23         DAG    Implement data driven exception handling through
'                             class WWXLAppExceptions.
'
'   2008/08/30 1.03    DAG/WW Relocate to module XLRangeLib_WW, and remove some
'                             incorrect information from the documentation.
'
'   ----------------------------------------------------------------------------

    Const MODULE_STARTING_ERROR_MSG_INDEX As Long = 1

    Const COL_TOKEN As String = "column"
    Const END_OF_RANGE_INDEX As String = "Last $$INDEX$$ in range = "

    Const INDEX_TOKEN As String = "$$INDEX$$"

    Const INVALID_ARG_VALUE_LBL As String = "Argument value = "
    Const INVALID_LOWER_RIGHT_ROW As Long = MODULE_STARTING_ERROR_MSG_INDEX + 8
    Const INVALID_LOWER_RIGHT_COL As Long = MODULE_STARTING_ERROR_MSG_INDEX + 9
    Const INVALID_UPPER_LEFT_ROW As Long = MODULE_STARTING_ERROR_MSG_INDEX + 6
    Const INVALID_UPPER_LEFT_COL As Long = MODULE_STARTING_ERROR_MSG_INDEX + 7

    Const LOWER_VALUE_LBL As String = "Lower right corner value = "

    Const UPPER_LEFT_GT_LOWER_RIGHT_ROW As Long = MODULE_STARTING_ERROR_MSG_INDEX + 10
    Const UPPER_LEFT_GT_LOWER_RIGHT_COL As Long = MODULE_STARTING_ERROR_MSG_INDEX + 11

    Const INPUT_RANGE_IS_NULL As Long = MODULE_STARTING_ERROR_MSG_INDEX + 12

    Const REQUESTED_RANGE_PAST_LAST_ROW As Long = MODULE_STARTING_ERROR_MSG_INDEX + 13
    Const REQUESTED_RANGE_PAST_LAST_COL As Long = MODULE_STARTING_ERROR_MSG_INDEX + 14

    Const REQUESTED_INDEX As String = "Requested $$INDEX$$ in range = "
    Const ROW_TOKEN As String = "row"
    Const UPPER_VALUE_LBL As String = "Upper left corner value = "

    On Error GoTo rngCellsFromRng_P6C_Err

'   ----------------------------------------------------------------------------
'   Connect to table of error messages through our error handler object.
'   ----------------------------------------------------------------------------

    Dim wwException As WWXLAppExceptions
    Set wwException = New WWXLAppExceptions
    wwException.ErrorListSheet = RANGELIB_MSG_WKS_NM

'   ----------------------------------------------------------------------------
'   Initialize new Range object reference to Nothing and allocate local storage.
'   ----------------------------------------------------------------------------

    Set rngCellsFromRng_P6C = Nothing

    Dim intColsIn As Integer
    Dim lngRowsIn As Long

    Dim strSubsetRangeAddress As String
    Dim strLowerRightAddress As String
    Dim strUpperLeftAddress As String
    Dim strWorksheetName As String

'   ----------------------------------------------------------------------------
'   Confirm that the coordinates of the upper left and lower right corners are
'   consistent with the definition of upper left corner and lower right corner.
'   ----------------------------------------------------------------------------

    If ValidRowIndex_WW(plngUpperLeftRow) = False Then
        wwException.Throw _
            INVALID_UPPER_LEFT_ROW, _
            INVALID_ARG_VALUE_LBL & plngUpperLeftRow
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If ValidRowIndex_WW(plngUpperLeftRow) = False Then

    If ValidColIndex_WW(pintUpperLeftCol) = False Then
        wwException.Throw _
            INVALID_UPPER_LEFT_COL, _
            INVALID_ARG_VALUE_LBL & pintUpperLeftCol
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If ValidColIndex_WW(pintUpperLeftCol) = False Then

    If ValidRowIndex_WW(plngLowerRightRow) = False Then
        wwException.Throw _
            INVALID_LOWER_RIGHT_ROW, _
            INVALID_ARG_VALUE_LBL & plngLowerRightRow
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If ValidRowIndex_WW(plngLowerRightRow) = False Then

    If ValidColIndex_WW(pintLowerRightCol) = False Then
        wwException.Throw _
            INVALID_LOWER_RIGHT_COL, _
            INVALID_ARG_VALUE_LBL & pintLowerRightCol
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If ValidColIndex_WW(pintLowerRightCol) = False Then

'   ----------------------------------------------------------------------------
'   These two are simplified, because the lower bound has already passed muster.
'   ----------------------------------------------------------------------------

    If plngUpperLeftRow > plngLowerRightRow Then
        wwException.Throw _
            UPPER_LEFT_GT_LOWER_RIGHT_ROW, _
            UPPER_VALUE_LBL & plngUpperLeftRow & vbLf _
                & LOWER_VALUE_LBL & plngLowerRightRow
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If plngUpperLeftRow > plngLowerRightRow Then

    If pintUpperLeftCol > pintLowerRightCol Then
        wwException.Throw _
            UPPER_LEFT_GT_LOWER_RIGHT_COL, _
            UPPER_VALUE_LBL & pintUpperLeftCol & vbLf _
                & LOWER_VALUE_LBL & pintLowerRightCol
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If pintUpperLeftCol > pintLowerRightCol Then

'   ----------------------------------------------------------------------------
'   Confirm that the Range object reference actually points to a Range object
'   and that the lower right corner of the subset range is inside the range.
'   ----------------------------------------------------------------------------

    If prngIn Is Nothing Then
        wwException.Throw INPUT_RANGE_IS_NULL
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If prngIn Is Nothing Then

    intColsIn = prngIn.Columns.Count
    lngRowsIn = prngIn.Rows.Count

    If plngLowerRightRow > lngRowsIn Then
        wwException.Throw _
            REQUESTED_RANGE_PAST_LAST_ROW, _
            Replace(END_OF_RANGE_INDEX, INDEX_TOKEN, ROW_TOKEN) & lngRowsIn & vbLf _
                & Replace(REQUESTED_INDEX, INDEX_TOKEN, ROW_TOKEN) & plngLowerRightRow
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If plngLowerRightRow > lngRowsIn Then

    If pintLowerRightCol > intColsIn Then
        wwException.Throw _
            REQUESTED_RANGE_PAST_LAST_COL, _
            Replace(END_OF_RANGE_INDEX, INDEX_TOKEN, COL_TOKEN) & intColsIn & vbLf _
                & Replace(REQUESTED_INDEX, INDEX_TOKEN, COL_TOKEN) & pintLowerRightCol
        GoTo rngCellsFromRng_P6C_End    ' This is now unreachable, but left to guard against tigers.
    End If  ' If pintLowerRightCol > intColsIn Then

'   ----------------------------------------------------------------------------
'   The input Range object and the coordinates of the requested subset of cells
'   are all valid. Construct the address of the output range, and create the
'   Range object.
'   ----------------------------------------------------------------------------

    With prngIn
        strWorksheetName = .Worksheet.Name
        strUpperLeftAddress = .Cells(plngUpperLeftRow, pintUpperLeftCol).Address
        strLowerRightAddress = .Cells(plngLowerRightRow, pintLowerRightCol).Address
    End With    ' With prngIn

    strSubsetRangeAddress = FORMULA_PREFIX & strWorksheetName _
                            & FORMULA_SHEET_NAME_DLM & strUpperLeftAddress _
                            & FORMULA_CELL_ADDR_DLM & strLowerRightAddress

    Set rngCellsFromRng_P6C = Range(strSubsetRangeAddress)
    GoTo rngCellsFromRng_P6C_End

rngCellsFromRng_P6C_Err:

    wwException.Show "rngCellsFromRng_P6C"

    If IsObject(rngCellsFromRng_P6C) Then
        Set rngCellsFromRng_P6C = Nothing
    End If

    Err.Raise wwException.Number, _
              wwException.Source, _
              wwException.Description
    Resume rngCellsFromRng_P6C_End      ' This is now unreachable, but left to guard against tigers.

rngCellsFromRng_P6C_End:

    Set wwException = Nothing           ' However we get here, discard the object.
    Exit Function

End Function


Private Function SheetNameForFormula(pws As Worksheet)

    If InStr(pws.Name, SPACE_CHAR_WW) > INSTR_NOT_FOUND Then
        SheetNameForFormula = QUOTE_CHAR_SGL_WW & pws.Name & QUOTE_CHAR_SGL_WW
    Else
        SheetNameForFormula = pws.Name
    End If

End Function


Private Function ValidColIndex_WW( _
        ByVal pintColIndex As Integer) _
            As Boolean

    '   ------------------------------------------------------------------------
    '
    '   Name:           ValidColIndex_WW
    '
    '   Synopsis:       Test the input value to ensure that it is a valid value
    '                   for the index of a Column in a Range or Worksheet.
    '
    '   Arguments:      pintColIndex    = Integer Column index to evaluate.
    '
    '   Returns:        TRUE if argument pintColIndex is a valid Column index.
    '
    '                   FALSE if argument pintColIndex is an invalid Column
    '                   index.
    '
    '   Notes:          Although the document cited below, "Excel Specifications
    '                   and Limits," states that it applies to Microsoft Office
    '                   Excel 2003, since the limits in question affect the file
    '                   format, which has remained the same since Excel 97, I
    '                   think it's safe to say that they apply to all versions
    '                   of Excel that share the same file format. This means
    '                   that they apply to the following versions.
    '
    '                           Excel 97
    '                           Excel 2000
    '                           Excel 2002 (Excel XP)
    '                           Office Excel 2003
    '
    '                   Since a contiguous range cannot span worksheets, it
    '                   follows that these limits also apply to any contiguous
    '                   Range object.
    '
    '                   Unlike most of the functions that I write, this one has
    '                   no error handler, since the runtime library won't allow
    '                   its caller to pass an uninitialized integer into it, and
    '                   the only assignment is to its own return value.
    '
    '                   For Excel 97-2003, the maximum number of columns that a
    '                   Worksheet may contain is 256. It follows, a priori, that
    '                   the same limit applies to the Column index of an Area of
    '                   a Range object, and to the Column index of a Cell in a
    '                   Cells collection, a Range, or a Worksheet.
    '
    '                   For Excel 2007 and 2010, this limit increases to 16384
    '                   columns.
    '
    '    References:    Excel Specifications and Limits
    '                   http://office.microsoft.com/en-us/excel/HP051992911033.aspx
    '
    '    Date Written:  Sunday, 17 August 2008.
    '
    '    Author:        David A. Gray, Simple Soft Services, Inc., d/b/a WizardWrx
    '
    '    Copyright:     (C) 2003-2012, Simple Soft Services, Inc., d/b/a WizardWrx
    '                   Irving, Texas, USA
    '                   http://www.wizardwrx.com/
    '                   All rights reserved world wide.
    '
    '   ------------------------------------------------------------------------
    '   Revision History
    '   ------------------------------------------------------------------------
    '
    '   Date       Version Author Synopsis
    '   ---------- ------- ------ ----------------------------------------------
    '   2008/08/24 1.00    DAG/WW Move this function from XLSUtis_P6C.
    '   2012/10/13 2.00    DAG/WW Replace hard coded upper limit with a version
    '                             aware computation.
    '   ========================================================================

    If pintColIndex > 0 And pintColIndex <= MaxColsInWorksheet() Then
        ValidColIndex_WW = True
    Else
        ValidColIndex_WW = False
    End If

End Function


Private Function ValidRowIndex_WW( _
        ByVal plngRowIndex As Long) _
            As Boolean

    '   ------------------------------------------------------------------------
    '
    '   Name:           ValidRowIndex_WW
    '
    '   Synopsis:       Test the input value to ensure that it is a valid value
    '                   for the index of a Row in a Range or Worksheet.
    '
    '   Arguments:      plngRowIndex    = Row index to evaluate.
    '
    '   Returns:        TRUE if plngRowIndex is a valid Row index.
    '
    '                   FALSE if plngRowIndex is an invalid Row index.
    '
    '   Notes:          Although the document cited below, "Excel Specifications
    '                   and Limits," states that it applies to Microsoft Office
    '                   Excel 2003, since the limits in question affect the file
    '                   format, which has remained the same since Excel 97, I
    '                   think it's safe to say that they apply to all versions
    '                   of Excel that share the same file format. This means
    '                   that they apply to the following versions.
    '
    '                           Excel 97
    '                           Excel 2000
    '                           Excel 2002 (Excel XP)
    '                           Office Excel 2003
    '
    '                   Since a contiguous range cannot span worksheets, it
    '                   follows that these limits also apply to any contiguous
    '                   Range object.
    '
    '                   Unlike most of the functions that I write, this one has
    '                   no error handler, since the runtime library won't allow
    '                   its caller to pass an uninitialized integer into it, and
    '                   the only assignment is to its own return value.
    '
    '                   For Excel 97-2003, the maximum number of rows that a
    '                   Worksheet may contain is 65535. It follows, a priori,
    '                   that the same limit applies to the Row index of an Area
    '                   of a Range object, and to the Column index of a Cell in
    '                   a Cells collection, a Range, or a Worksheet.
    '
    '                   For Excel 2007 and 2010, this limit increases to 1048576
    '                   rows.
    '
    '    References:    Excel Specifications and Limits
    '                   http://office.microsoft.com/en-us/excel/HP051992911033.aspx
    '
    '    Date Written:  Sunday, 17 August 2008.
    '
    '    Author:        David A. Gray, Simple Soft Services, Inc., d/b/a WizardWrx
    '
    '    Copyright:     (C) 2003-2012, Simple Soft Services, Inc., d/b/a WizardWrx
    '                   Irving, Texas, USA
    '                   http://www.wizardwrx.com/
    '                   All rights reserved world wide.
    '
    '   ------------------------------------------------------------------------
    '   Revision History
    '   ------------------------------------------------------------------------
    '
    '   Date       Version Author Synopsis
    '   ---------- ------- ------ ----------------------------------------------
    '   2008/08/24 1.00    DAG/WW Move this function from XLSUtis_P6C.
    '   2012/10/13 2.00    DAG/WW Replace hard coded upper limit with a version
    '                             aware computation.
    '   ========================================================================

    If plngRowIndex > 0 And plngRowIndex <= MaxRowsInWorksheet() Then
        ValidRowIndex_WW = True
    Else
        ValidRowIndex_WW = False
    End If

End Function


Public Function MaxColsInWorksheet() As Long

'   ============================================================================
'
'   Name:               MaxColsInWorksheet
'
'   Synopsis:           Return the number of columns in the active worksheet.
'
'   Arguments:          None.
'
'   Return Value:       The return value is the upper limit of the index
'                       (number) of a column in the active worksheet, expressed
'                       as a long integer.
'
'   ============================================================================

    Dim ws As Worksheet: Set ws = ActiveWorkbook.ActiveSheet
    MaxColsInWorksheet = ws.Columns.Count

End Function


Public Function MaxRowsInWorksheet() As Long

'   ============================================================================
'
'   Name:               MaxRowsInWorksheet
'
'   Synopsis:           Return the number of rows in the active worksheet.
'
'   Arguments:          None.
'
'   Return Value:       The return value is the upper limit of the index
'                       (number) of a row in the active worksheet, expressed as
'                       a long integer.
'
'   ============================================================================

    Dim ws As Worksheet: Set ws = ActiveWorkbook.ActiveSheet
    MaxRowsInWorksheet = ws.Rows.Count

End Function
