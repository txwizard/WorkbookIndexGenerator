VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WWXLAppExceptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '   ========================================================================
    '
    '   Class Name:     WWXLAppExceptions
    '
    '   Synopsis:       Format and display error messages from a message number
    '                   that starts from 1, and increases by 1 for however many
    '                   unique errors are defined for the project. Messages,
    '                   numbered from 1024, are stored in a designated worksheet
    '                   range. See Notes 1 and 2.
    '
    '   Remarks:        1) Error message numbers and text are stored in a range
    '                      named rngWWXLAppExceptionList of a worksheet named
    '                      wsWWXLAppErrors.
    '
    '                   2) The starting error number is stored in range
    '                      rngStartingErrorNumber, a one-cell range, in the same
    '                      workbook. This class uses this value only indirectly.
    '                      Its real use is as the seed for the error numbers in
    '                      the second (B) column of rngWWXLAppExceptionList. The
    '                      class DOES use those numbers.
    '
    '                   3) The project name is stored in rngWWXLAProjectName,
    '                      another single-cell range. You will almost certainly
    '                      need to change this value if you import this class
    '                      and its accompanying worksheet into another project,
    '                      because the class DOES use this value.
    '
    '                   4) The default module name is stored in range
    '                      rngWWXLADefaultModuleName, another single-cell range.
    '                      If your project contains two or more modules, you can
    '                      override this name in code by setting the ModuleName
    '                      property.
    '
    '                   5) As of version 1.04, these names are scoped to the
    '                      individual error message worksheets. I've moved the
    '                      corresponding string constants into private method,
    '                      Read_Props_From_Sheet. However, for easy reference, I
    '                      am leaving their names and definitions here.
    '
    '                           ERROR_MESSAGE_RNG_NM = Error message list
    '                           DEFAULT_ERR_MSG_WKS_NM = Error Message Sheet
    '                           PROJECT_NAME_RNG_NM = VBA Project Name
    '                           DEFAULT_MODULE_RNG_NM = VBA Module name
    '                           STARTING_ERROR_NUMBER_RNG_NM = First error #
    '
    '                      See the article cited in the second reference for an
    '                      explanation of how I was able to confine the scope of
    '                      these ranges to a single sheet.
    '
    '                   6) Magic number 32752 is 2 to the 15th power, less 2 to
    '                      the 4th power. This number reserves a range of 15
    '                      error numbers for use by this class, which is the only
    '                      class in a project that cannot use the error message
    '                      worksheets.
    '
    '                   7) Constant DEFAULT_PROJECT_NAME is set to the default
    '                      name of a VBA project, VBAProject. Change this name
    '                      to correspond to the actual name of the VBA project,
    '                      as shown in the project property sheet.
    '
    '                   8) Beginning with version 1.10, this class uses the
    '                      ThisWorkbook property, rather than ActiveWorkbook, so
    '                      that the code is usable in an add-in.
    '
    '   References:     1)  "The VBA Err Object," Microsoft Office XP Developer,
    '                       MSDN Library, Office Developer Center
    '                       http://msdn.microsoft.com/en-us/library/aa164768(office.10).aspx
    '
    '                   2)  "Using the Same Range Name on Different Worksheets,"
    '                       Allan Wyatt's ExcelTips
    '                       http://exceltips.vitalnews.com/Pages/T002662_Using_the_Same_Range_Name_on_Different_Worksheets.html.
    '
    '   Created:        Saturday, 09 August 2008
    '
    '   Author:         David A. Gray, Chief Wizard
    '                   Simple Soft Services, Inc., d/b/a WizardWrx
    '                   http://www.wizardwrx.com/
    '
    '   Copyright:      2008-2012, Simple Soft Services, Inc., d/b/a WizardWrx
    '                   All rights reserved world wide.
    '
    '   ------------------------------------------------------------------------
    '   Revision History
    '   ------------------------------------------------------------------------
    '
    '   Date       Version Author Synopsis
    '   ---------- ------- ------ ----------------------------------------------
    '   2008/08/11 1.00    DAG/WW Add code to save after update.
    '
    '   2008/08/13 1.01    DAG/WW Move all the data that will or may change from
    '                             project to project into named ranges in sheet
    '                             wsWWXLAppErrors, and make them available as
    '                             read only properties.
    '
    '   2008/08/16 1.02    DAG/WW Cosmetic cleanup of documentation.
    '
    '   2008/08/21 1.03    DAG/WW Expose the error number and message as
    '                             properties, so that I can call the Raise
    '                             method on the regular Err object, to force an
    '                             exception to bubble up the call stack.
    '
    '   2008/08/23 1.04    DAG/WW 1) Add Read/Write ErrorListSheet property,
    '                                which allows the default exception report
    '                                worksheet, wsWWXLAppErrors, to be changed.
    '
    '                             2) Add a Read only TargetSite property,
    '                                which method Throw sets, to leave
    '                                a record of the routine that throew the
    '                                exception, and add it to the error message.
    '
    '                             3) Eliminate redundant property, ObjectError.
    '
    '                             4) Add the TargetSite property to the message
    '                                string stored and displayed by the Show
    '                                method.
    '
    '                             5) Shorten the name of the ThrowException
    '                                method to Throw.
    '
    '   2008/08/24 1.05    DAG/WW Add code to the Show method to detect when it
    '                             is being called directly by the error handler
    '                             for a VBA routine to handle a run-time error
    '                             that was generated at the current level of the
    '                             call stack.
    '
    '   2008/08/27 1.06    DAG/WW Handle the exception caused by attempting to
    '                             set property ErrorListSheet to an invalid
    '                             worksheet.
    '
    '   2008/08/30 1.07    DAG/WW 1) Add comprehenisve error trapping to private
    '                                method, Read_Props_From_Sheet.
    '
    '                             2) Define commonly used error constants, and
    '                                expose them as read only properties.
    '
    '                                Property Name       Property Value
    '                                ------------------- -----------------------
    '                                ErrNbrNoError       NO_ERROR
    '                                ErrSubscrOutOfRange ERR_SUBSCR_OUT_OF_RANGE
    '                                ErrObjectDefined    ERR_OBJ_DEFINED_ERROR
    '                                ErrObjOrWithNotSet  ERR_OBJ_OR_WITH_NOT_SET
    '                                ------------------- -----------------------
    '
    '   2008/09/05 1.08    DAG/WW Reorganize constant definitions, and improve
    '                             the internal technical documentation.
    '
    '   2009/11/01 1.09    DAG/WW Add an optional Boolean argument RuntimeError
    '                             to the Throw method, with a default value of
    '                             False (to avoid breeaking existing code), and
    '                             have the method take the error number at face
    '                             value if RuntimeError is TRUE.
    '
    '   2012/10/13 1.10    DAG/WW 1) Change reference to ActiveWorkbook property
    '                                to ThisWorkbook, so that this code looks in
    '                                the workbook that contains it for its
    '                                reference worksheets. See Remarks.
    '
    '                             2) Clean up the internal documentation, adding
    '                                omitted arguments and correcting errors due
    '                                to copying of comment blocks between its
    '                                routines.
    '
    '   ========================================================================

    '   ------------------------------------------------------------------------
    '   Private constants.
    '   ------------------------------------------------------------------------

    Const DEFAULT_PROJECT_NAME As String = "VBAProject"     ' May be adjusted to match project name, per Project Properties

    Const ERR_SUBSCR_OUT_OF_RANGE As Long = 9
    Const MODULE_STARTING_ERROR_NUMBER As Long = 32752      ' Used only in the following constant definitions.

    Const ERR_NBR_RANGENAME_IS_UNDEFINED As Long = MODULE_STARTING_ERROR_NUMBER + 2
    Const ERR_NBR_SHEETNAME_IS_BLANK As Long = MODULE_STARTING_ERROR_NUMBER
    Const ERR_NBR_SHEETNAME_IS_INVALID As Long = MODULE_STARTING_ERROR_NUMBER + 1

    Const ERR_OBJ_DEFINED_ERROR As Long = 1004
    Const ERR_OBJ_OR_WITH_NOT_SET As Long = 91

    Const NO_ERROR As Long = 0

    Const UNDEFINED As String = "Undefined"

    '   ------------------------------------------------------------------------
    '   Private storage for properties
    '   ------------------------------------------------------------------------

    Dim mlngStartingErrorNumber As Long
    Dim mlngErrorNumber As Long

    Dim mstrErrorListSheet As String
    Dim mstrErrorMessage As String
    Dim mstrModuleName As String
    Dim mstrProjectName As String
    Dim mstrTargetSite As String

    '   ------------------------------------------------------------------------
    '   Private storage for internal use
    '   ------------------------------------------------------------------------

    Dim mblnBubbleUp As Boolean                 ' Are we bulling an exception up?
    Dim mwksErrorData As Worksheet              ' Messages live in this sheet.
    Dim mrngWWXLAppExceptionList As Range       ' This range has a standard name.


    '   +----------------------------------------------------------------------+
    '   |                                                                      |
    '   |                          P R O P E R T I E S                         |
    '   |                                                                      |
    '   +----------------------------------------------------------------------+


Public Property Get Description() As String

    '   ------------------------------------------------------------------------
    '   The Description and Message properties are synonymns, created to
    '   simplify alternating between working in VBA and in VB.NET.
    '   ------------------------------------------------------------------------

    Description = mstrErrorMessage

End Property


Public Property Get ErrNbrNoError() As Long

    '   ------------------------------------------------------------------------
    '   Since VBA forbids public and global constants in class modules, the only
    '   way to make a constant public is to return it as a read only property.
    '   ------------------------------------------------------------------------

    ErrNbrNoError = NO_ERROR

End Property


Public Property Get ErrObjectDefined() As Long

    '   ------------------------------------------------------------------------
    '   Since VBA forbids public and global constants in class modules, the only
    '   way to make a constant public is to return it as a read only property.
    '   ------------------------------------------------------------------------

    ErrObjectDefined = ERR_OBJ_DEFINED_ERROR

End Property

Public Property Get ErrObjOrWithNotSet() As Long

    '   ------------------------------------------------------------------------
    '   Since VBA forbids public and global constants in class modules, the only
    '   way to make a constant public is to return it as a read only property.
    '   ------------------------------------------------------------------------

    ErrObjOrWithNotSet = ERR_OBJ_OR_WITH_NOT_SET

End Property


Public Property Get ErrSubscrOutOfRange() As Long

    '   ------------------------------------------------------------------------
    '   Since VBA forbids public and global constants in class modules, the only
    '   way to make a constant public is to return it as a read only property.
    '   ------------------------------------------------------------------------

    ErrSubscrOutOfRange = ERR_SUBSCR_OUT_OF_RANGE

End Property


Public Property Get ErrorListSheet() As String

    ErrorListSheet = mstrErrorListSheet

End Property

Public Property Let ErrorListSheet(ByRef pstrErrorListSheet As String)

    Dim fTrapInvalidSheetName As Boolean
    On Error GoTo ErrorListSheet_Err

    If pstrErrorListSheet <> vbNullString Then
        Set mwksErrorData = Nothing
        fTrapInvalidSheetName = True
        Set mwksErrorData = ThisWorkbook.Worksheets(pstrErrorListSheet)
        fTrapInvalidSheetName = False

        mstrErrorListSheet = pstrErrorListSheet
        Call Read_Props_From_Sheet
    Else
        mlngErrorNumber = ERR_NBR_SHEETNAME_IS_BLANK
        mstrErrorMessage = "The ErrorListSheet property cannot be blank."
        Err.Raise mlngErrorNumber Or vbObjectError, _
                  DEFAULT_PROJECT_NAME & ".ErrorListSheet_Property", _
                  mstrErrorMessage
    End If

    GoTo ErrorListSheet_End

ErrorListSheet_Err:

    If fTrapInvalidSheetName Then
        If Err.Number = ERR_SUBSCR_OUT_OF_RANGE Then
            mlngErrorNumber = ERR_NBR_SHEETNAME_IS_INVALID
            mstrErrorMessage = "Specified worksheet name, " _
                & pstrErrorListSheet & _
                ", is invalid."
            Err.Raise mlngErrorNumber Or vbObjectError, _
                      DEFAULT_PROJECT_NAME & ".ErrorListSheet_Property", _
                      mstrErrorMessage
        Else
            mlngErrorNumber = Err.Number
            mstrErrorMessage = Err.Description
            Err.Raise mlngErrorNumber Or vbObjectError, _
                      DEFAULT_PROJECT_NAME & ".ErrorListSheet_Property", _
                      mstrErrorMessage
        End If
    Else
        mlngErrorNumber = Err.Number
        mstrErrorMessage = Err.Description
        Err.Raise mlngErrorNumber, _
                  DEFAULT_PROJECT_NAME & ".ErrorListSheet_Property", _
                  mstrErrorMessage
    End If

ErrorListSheet_End:

End Property


Public Property Get ErrorSource() As String

    Const ERROR_SOURCE_TMPLATE As String = "$$PROJ$$.$$MODULE"
    Const ERROR_SOURCE_PROJECT_TOKEN As String = "$$PROJ$$"
    Const ERROR_SOURCE_MODULE_TOKEN As String = "$$MODULE"

    ErrorSource = Replace( _
        Replace( _
            ERROR_SOURCE_TMPLATE, _
            ERROR_SOURCE_PROJECT_TOKEN, _
            mstrProjectName), _
        ERROR_SOURCE_MODULE_TOKEN, _
        mstrModuleName)

End Property


Public Property Get Message() As String

    '   ------------------------------------------------------------------------
    '   The Description and Message properties are synonymns, created to
    '   simplify alternating between working in VBA and in VB.NET.
    '   ------------------------------------------------------------------------

    Message = mstrErrorMessage

End Property


Public Property Get ModuleName() As String

    ModuleName = mstrModuleName

End Property

Public Property Let ModuleName(ByRef pstrModuleName As String)

    If pstrModuleName <> vbNullString Then
        mstrModuleName = pstrModuleName
    End If

End Property


Public Property Get Number() As Long

    '   ------------------------------------------------------------------------
    '   The Exception object of the Microsoft .NET Framework doesn't have a
    '   Number property, per se, though it has a protected HRESULT property.
    '   ------------------------------------------------------------------------

    Number = mlngErrorNumber

End Property


Public Property Get ProjectName() As String

    ProjectName = mstrProjectName

End Property


Public Property Get Source() As String

    '   ------------------------------------------------------------------------
    '   The ErrorSource and Source properties are synonymns. I added the latter
    '   for compatibility with the native Err object. However, since ErrorSource
    '   is a computed property, this property is implemented as a call to the
    '   ErrorSource property.
    '   ------------------------------------------------------------------------

    Source = Me.ErrorSource

End Property


Public Property Get StartingErrorNumber() As Long

    '   ------------------------------------------------------------------------
    '   This property reports the starting error number, per the active error
    '   message worksheet, which is identified by the ErrorListSheet property.
    '   ------------------------------------------------------------------------

    StartingErrorNumber = mlngStartingErrorNumber

End Property


Public Property Get TargetSite() As String

    '   ------------------------------------------------------------------------
    '   This property is analogous to the like named property in the Exception
    '   class of the Microsoft .NET Framework.
    '   ------------------------------------------------------------------------

    TargetSite = mstrTargetSite

End Property


    '   +----------------------------------------------------------------------+
    '   |                                                                      |
    '   |                        C O N S T R U C T O R                         |
    '   |                                                                      |
    '   +----------------------------------------------------------------------+


Private Sub Class_Initialize()

    '   ========================================================================
    '
    '   Name:           Class_Initialize, a VBA object constructor.
    '
    '   Synopsis:       Constructor for a VBA class.
    '
    '   Arguments:      None.
    '
    '   Returns:        Nothing. It's a subroutine (a VOID function).
    '
    '   Remarks:        This code runs whenever a new instance of this class is
    '                   created. It does the following:
    '
    '                   1) Initialize user settable properties to safe defaults.
    '
    '                   2) Initialize read only properties to known values.
    '
    '                   3) Initialize other internal class variables that have
    '                      nonzero values.
    '
    '                   Beginning with version 1.10, this class uses the
    '                   ThisWorkbook property, rather than ActiveWorkbook, so
    '                   that the code is usable in an add-in.
    '
    '   Created:        Monday, 11 August 2008
    '
    '   ------------------------------------------------------------------------
    '   Revision History
    '   ------------------------------------------------------------------------
    '
    '   Date       Version Author Synopsis
    '   ---------- ------- ------ ----------------------------------------------
    '   2008/08/11 1.00    DAG/WW Add code to save after update.
    '
    '   2008/08/13 1.01    DAG/WW Move all the data that will or may change from
    '                             project to project into named ranges in sheet
    '                             wsWWXLAppErrors, and make them available as
    '                             read only properties.
    '
    '   2008/08/21 1.03    DAG/WW Expose the error number and message as
    '                             properties, so that I can call the Raise
    '                             method on the regular Err object, to force an
    '                             exception to bubble up the call stack.
    '
    '   2008/08/23 1.04    DAG/WW Initialize the ErrorListSheet property to its
    '                             default value. Move the initialization code
    '                             that gets data from the worksheet, to which
    '                             this property points, into a new routine that
    '                             can be called from the method that sets the
    '                             new property.
    '
    '   2008/08/27 1.06    DAG/WW Handle the exception caused by attempting to
    '                             set property ErrorListSheet to an invalid
    '                             worksheet.
    '
    '   2008/08/30 1.07    DAG/WW Correct a logic error in the error handler
    '                             that would have allowed some errors to go
    '                             unreported.
    '
    '   2012/10/13 1.10    DAG/WW Change reference to ActiveWorkbook property to
    '                             ThisWorkbook, so that this code looks in the
    '                             workbook that contains it for its reference
    '                             worksheets. See Remarks.
    '
    '   ========================================================================

    Const DEFAULT_ERR_MSG_WKS_NM As String = "wsWWXLAppErrors"

    Dim fTrapInvalidSheetName As Boolean

    On Error GoTo Class_Initialize_Err:

    '   ------------------------------------------------------------------------
    '   Set a reference to the default error message worksheet.
    '   ------------------------------------------------------------------------

    fTrapInvalidSheetName = True
    Set mwksErrorData = ThisWorkbook.Worksheets(DEFAULT_ERR_MSG_WKS_NM)
    fTrapInvalidSheetName = False
    mstrErrorListSheet = DEFAULT_ERR_MSG_WKS_NM

    '   ------------------------------------------------------------------------
    '   Read property values from the sheet.
    '   ------------------------------------------------------------------------

    Call Read_Props_From_Sheet

    '   ------------------------------------------------------------------------
    '   Assign defaults to other properties.
    '   ------------------------------------------------------------------------

    mlngErrorNumber = NO_ERROR
    mstrErrorMessage = vbNullString
    mstrTargetSite = UNDEFINED

    mblnBubbleUp = False

    GoTo Class_Initialize_End

Class_Initialize_Err:

    If fTrapInvalidSheetName Then
        If Err.Number = ERR_SUBSCR_OUT_OF_RANGE Then
            mlngErrorNumber = ERR_NBR_SHEETNAME_IS_INVALID
            mstrErrorMessage = "Specified worksheet name, " _
                & DEFAULT_ERR_MSG_WKS_NM & _
                ", is invalid."
            Err.Raise mlngErrorNumber Or vbObjectError, _
                      DEFAULT_PROJECT_NAME & ".ErrorListSheet_Property", _
                      mstrErrorMessage
        Else
            mlngErrorNumber = Err.Number
            mstrErrorMessage = "Error # " & Err.Number & " - " & Err.Description
            Err.Raise mlngErrorNumber, _
                      DEFAULT_PROJECT_NAME & ".WWXLAppExceptions", _
                      mstrErrorMessage
        End If
    Else
        mlngErrorNumber = Err.Number
        mstrErrorMessage = "Error # " & Err.Number & " - " & Err.Description
        Err.Raise mlngErrorNumber, _
            DEFAULT_PROJECT_NAME & ".WWXLAppExceptions", _
            mstrErrorMessage
    End If
    Resume Class_Initialize_End

Class_Initialize_End:

    Exit Sub

End Sub


    '   +----------------------------------------------------------------------+
    '   |                                                                      |
    '   |                         D E S T R U C T O R                          |
    '   |                                                                      |
    '   +----------------------------------------------------------------------+

Private Sub Class_Terminate()

    '   ========================================================================
    '
    '   Name:           Class_Terminate, a VBA object destructor.
    '
    '   Synopsis:       Destructor for a VBA class.
    '
    '   Arguments:      None.
    '
    '   Returns:        Nothing. It's a subroutine (a VOID function).
    '
    '   Remarks:        This routine does proper cleanup to enable Windows to
    '                   discard data structures that represent objects and
    '                   memory blocks assigned from the process heap. After all,
    '                   we want to be good Windows citizens. --:)
    '
    '   Created:        Monday, 11 August 2008
    '
    '   ------------------------------------------------------------------------
    '   Revision History
    '   ------------------------------------------------------------------------
    '
    '   Date       Version Author Synopsis
    '   ---------- ------- ------ ----------------------------------------------
    '   2008/08/11 1.00    DAG/WW Add code to save after update.
    '
    '   2008/08/13 1.01    DAG/WW Move all the data that will or may change from
    '                             project to project into named ranges in sheet
    '                             wsWWXLAppErrors, and make them available as
    '                             read only properties.
    '
    '   2008/08/21 1.03    DAG/WW Expose the error number and message as
    '                             properties, so that I can call the Raise
    '                             method on the regular Err object, to force an
    '                             exception to bubble up the call stack.
    '
    '   2012/10/13 1.10    DAG/WW Clean up the internal documentation, adding
    '                             omitted arguments and correcting errors due to
    '                             copying of comment blocks between routines.
    '
    '   ========================================================================

    On Error GoTo Class_Terminate_Err

    Set mrngWWXLAppExceptionList = Nothing
    Set mwksErrorData = Nothing

    GoTo Class_Terminate_End

Class_Terminate_Err:

    mlngErrorNumber = Err.Number
    mstrErrorMessage = "Error # " & Err.Number & " - " & Err.Description
    Err.Raise mlngErrorNumber, _
              DEFAULT_PROJECT_NAME & ".WWXLAppExceptions", _
              mstrErrorMessage
    Resume Class_Terminate_End

Class_Terminate_End:

    Exit Sub

End Sub


    '   +----------------------------------------------------------------------+
    '   |                                                                      |
    '   |                            M E T H O D S                             |
    '   |                                                                      |
    '   +----------------------------------------------------------------------+


Public Sub Show(Optional ByRef pstrTargetSite As String = vbNullString)

    '   ========================================================================
    '
    '   Name:           Show
    '
    '   Synopsis:       Construct and display an error message from the
    '                   properties of the intirinsic VBA Err object and the name
    '                   of the active workbook.
    '
    '   Arguments:      pstrTargetSite  = Optional string containing the name of
    '                                     the calling function. If the argument
    '                                     is prsent, its value is saved into the
    '                                     TargetSite property.
    '
    '   Returns:        Nothing. This is a Sub.
    '
    '   Remarks:        If this method gets called because a message is working
    '                   its way up the call stack, we don't want to display it
    '                   again, nor do we want to tamper with the message.
    '
    '                   However, if this method is called directly by the error
    '                   handler for a routine, to report a VBA run-time error,
    '                   the Source property of the Err object will always differ
    '                   from the internal Source property of this instance of
    '                   the WWXLAppExceptions class.
    '
    '                   When this method is called to handle a VBA run-time
    '                   error, by the routine in which the error happened, the
    '                   TargetSite property of the instance will be equal to
    '                   "Undefined." This is a more dependable way to distin-
    '                   guish a first-chance exception from one that is being
    '                   passed up an unwinding call stack.
    '
    '                   Beginning with version 1.10, this class uses the
    '                   ThisWorkbook property, rather than ActiveWorkbook, so
    '                   that the code is usable in an add-in.
    '
    '   References:     "The VBA Err Object," Microsoft Office XP Developer,
    '                   MSDN Library, Office Developer Center
    '                   http://msdn.microsoft.com/en-us/library/aa164768(office.10).aspx
    '
    '   Created:        Saturday, 09 August 2008
    '
    '   ------------------------------------------------------------------------
    '   Revision History
    '   ------------------------------------------------------------------------
    '
    '   Date       Version Author Synopsis
    '   ---------- ------- ------ ----------------------------------------------
    '   2008/08/11 1.00    DAG/WW Add code to save after update.
    '
    '   2008/08/13 1.01    DAG/WW Move all the data that will or may change from
    '                             project to project into named ranges in sheet
    '                             wsWWXLAppErrors, and make them available as
    '                             read only properties.
    '
    '   2008/08/21 1.03    DAG/WW Expose the error number and message as
    '                             properties, so that I can call the Raise
    '                             method on the regular Err object, to force an
    '                             exception to bubble up the call stack.
    '
    '   2008/08/23 1.04    DAG/WW 1) Add code to identify exceptions that were
    '                                reported by a routine below the current one
    '                                in the call stack, and do the following.
    '
    '                                A) Avoid reporting the same error twice.
    '
    '                                B) Preserve the error number and message
    '                                   text, so that routines above the current
    '                                   one can use use them to log the error.
    '
    '                             2) Add the name of the calling routine to the
    '                                message template.
    '
    '   2008/08/24 1.05    DAG/WW Add code to the Show method to detect when it
    '                             is being called directly by the error handler
    '                             for a VBA routine to handle a run-time error
    '                             that was generated at the current level of the
    '                             call stack.
    '
    '   2012/10/13 1.10    DAG/WW Change reference to ActiveWorkbook property to
    '                             ThisWorkbook, so that this code looks in the
    '                             workbook that contains it for its reference
    '                             worksheets. See Remarks.
    '
    '   ========================================================================

    Const ERR_MSG_TPL As String = "A run-time exception occurred in function $$RTN_NAME$$ of module $$MOD_NAME$$." & vbLf & "$$ERR_TYP$$ Error # $$ERR_NBR$$: $$ERR_MSG$$"
    Const ERR_NBR_TOKEN As String = "$$ERR_NBR$$"
    Const ERR_MSG_TOKEN As String = "$$ERR_MSG$$"
    Const ERR_TYP_TOKEN As String = "$$ERR_TYP$$"
    Const ERR_MOD_NAME_TOKEN As String = "$$MOD_NAME$$"
    Const ERR_RTN_NAME_TOKEN As String = "$$RTN_NAME$$"

    Const ERR_TYP_APP As String = "Application"
    Const ERR_TYP_VBA As String = "VBA"

    Dim blnFirstChance As Boolean

    blnFirstChance = False                      ' Assume nothing about the error.

    If mstrTargetSite = UNDEFINED Then
        blnFirstChance = True                   ' It's a first-chance exception.
    End If

    '   ------------------------------------------------------------------------
    '   If present, update property TargetSite from the optional pstrTargetSite.
    '   ------------------------------------------------------------------------

    If pstrTargetSite <> vbNullString Then
        mstrTargetSite = pstrTargetSite
    End If

    '   ------------------------------------------------------------------------
    '   If another member of the class that created me already reported this
    '   exception, stop here.
    '   ------------------------------------------------------------------------

    If mblnBubbleUp Then
        Exit Sub
    End If

    '   ------------------------------------------------------------------------
    '   If a member of another class already reported this exception, stop here.
    '   If we've already determined that it's a first-chance exception, keep
    '   going.
    '   ------------------------------------------------------------------------

    If blnFirstChance Then
        mblnBubbleUp = True
    Else
        If Err.Source <> Me.Source Then
            mblnBubbleUp = True
            Exit Sub
        End If
    End If

    '   ------------------------------------------------------------------------
    '   Mark this exception as "reported," so that error handlers for routines
    '   above the routine that called this method know that the exception has
    '   been reported, and don't deluge the user with duplicate error messages.
    '   ------------------------------------------------------------------------

    mblnBubbleUp = True

    '   ------------------------------------------------------------------------
    '   Before the identifying bits are stripped from the error code, use them
    '   to differentiate run-time and application errors in the message text.
    '   ------------------------------------------------------------------------

    mlngErrorNumber = Err.Number

    If Err.Number And vbObjectError Then
        mstrErrorMessage = Replace( _
            Replace( _
                ERR_MSG_TPL, _
                ERR_NBR_TOKEN, _
                Err.Number Xor vbObjectError), _
            ERR_TYP_TOKEN, _
            ERR_TYP_APP)
    Else
        mstrErrorMessage = Replace( _
            Replace( _
                ERR_MSG_TPL, _
                ERR_NBR_TOKEN, _
                Err.Number), _
            ERR_TYP_TOKEN, _
            ERR_TYP_APP)
    End If

    '   ------------------------------------------------------------------------
    '   Replace the remaining tokens one at a time, to keep the code simple.
    '   ------------------------------------------------------------------------

    mstrErrorMessage = Replace( _
            mstrErrorMessage, _
            ERR_MSG_TOKEN, _
            Err.Description)

    mstrErrorMessage = Replace( _
        mstrErrorMessage, _
        ERR_MOD_NAME_TOKEN, _
        mstrModuleName)

    mstrErrorMessage = Replace( _
        mstrErrorMessage, _
        ERR_RTN_NAME_TOKEN, _
        mstrTargetSite)

    '   ------------------------------------------------------------------------
    '   Show the error message.
    '
    '   The text is already saved in the Description and Mesaage properties,
    '   which are synonymns.
    '
    '   The error number is already saved in the Number property, and the name
    '   of the routine that raised the exception is in the TargetSite property.
    '   ------------------------------------------------------------------------

    MsgBox mstrErrorMessage, _
        vbExclamation, _
        ThisWorkbook.Name

    Exit Sub

End Sub


Public Sub Throw( _
        ByVal plngAppErrNbr As Long, _
        Optional ByRef pstrExtraInfo As String = vbNullString, _
        Optional ByVal RuntimeError As Boolean = False)

    '   ========================================================================
    '
    '   Name:           Throw
    '
    '   Synopsis:       Construct and display an error message from the
    '                   properties of the intirinsic VBA Err object and the name
    '
    '   Arguments:      plngAppErrNbr = A Long integer that contains the number
    '                                   to look up in the error message table,
    '                                   and to assign to the Number property of
    '                                   the Err object.
    '
    '                   pstrExtraInfo = An optional string literal, to append to
    '                                   the error message.
    '
    '                   RuntimeError  = An optional Boolean, which alters the
    '                                   way the routine behaves when the error
    '                                   being processed is a VBA runtime error.
    '
    '   Remarks:        None.
    '
    '   References:     "The VBA Err Object," Microsoft Office XP Developer,
    '                   MSDN Library, Office Developer Center
    '                   http://msdn.microsoft.com/en-us/library/aa164768(office.10).aspx
    '
    '   Created:        Saturday, 09 August 2008
    '
    '   ------------------------------------------------------------------------
    '   Revision History
    '   ------------------------------------------------------------------------
    '
    '   Date       Version Author Synopsis
    '   ---------- ------- ------ ----------------------------------------------
    '   2008/08/11 1.00    DAG/WW Add code to save after update.
    '
    '   2008/08/13 1.01    DAG/WW Move all the data that will or may change from
    '                             project to project into named ranges in sheet
    '                             wsWWXLAppErrors, and make them available as
    '                             read only properties.
    '
    '   2008/08/21 1.03    DAG/WW Expose the error number and message as
    '                             properties, so that I can call the Raise
    '                             method on the regular Err object, to force an
    '                             exception to bubble up the call stack.
    '
    '   2008/08/23 1.04    DAG/WW 1) Store the error message in the Description
    '                                property, and eliminate the local string.
    '
    '                             2) Save the name of the function that threw
    '                                the exception into a new mstrTargetSite
    '                                property.
    '
    '                             3) Incorporate the new TargetSite property
    '                                into the error message.
    '
    '                             4) Amend the format of the output message, so
    '                                that pstrExtraInfo, if prseent, is preceded
    '                                by a linefeed character, so that it starts
    '                                a new line in the message box display.
    '
    '   2009/11/01 1.06    DAG/WW Add an optional Boolean argument RuntimeError,
    '                             with a default value of False.
    '
    '   2012/10/13 1.10    DAG/WW Clean up the internal documentation, adding
    '                             omitted arguments and correcting errors due to
    '                             copying of comment blocks between routines.
    '
    '   ========================================================================

    Const INDEX_COL = 1
    Const NUMBER_COL = 2
    Const FUNCTION_NAME_COL = 3
    Const MESSAGE_TEMPLATE_COL = 4

    Dim lngMsgRow As Long
    Dim lngMsgIndex As Long
    Dim lngMsgNumber As Long

    If RuntimeError = True Then
        Err.Raise plngAppErrNbr, _
                  Me.Source, _
                  pstrExtraInfo
    Else
        If plngAppErrNbr > mrngWWXLAppExceptionList.Rows.Count Then
            lngMsgRow = mrngWWXLAppExceptionList.Rows.Count
        Else
            lngMsgRow = plngAppErrNbr
        End If

        lngMsgIndex = mrngWWXLAppExceptionList.Cells(lngMsgRow, INDEX_COL)
        lngMsgNumber = mrngWWXLAppExceptionList.Cells(lngMsgRow, NUMBER_COL)

        mstrTargetSite = mrngWWXLAppExceptionList.Cells( _
            lngMsgRow, _
            FUNCTION_NAME_COL)

        If pstrExtraInfo = vbNullString Then
            mstrErrorMessage = mrngWWXLAppExceptionList.Cells( _
                lngMsgRow, _
                MESSAGE_TEMPLATE_COL)
        Else
            mstrErrorMessage = mrngWWXLAppExceptionList.Cells( _
                    lngMsgRow, _
                    MESSAGE_TEMPLATE_COL) & vbLf _
                    & pstrExtraInfo
        End If

        mlngErrorNumber = lngMsgNumber Or vbObjectError
        Err.Raise mlngErrorNumber, _
                  Me.ErrorSource, _
                  mstrErrorMessage
    End If

End Sub


    '   +----------------------------------------------------------------------+
    '   |                                                                      |
    '   |                    P R I V A T E   M E T H O D S                     |
    '   |                                                                      |
    '   +----------------------------------------------------------------------+


Private Sub Read_Props_From_Sheet()

    '   ========================================================================
    '
    '   Name:           Read_Props_From_Sheet
    '
    '   Synopsis:       Initialize object properties.
    '
    '   Remarks:        I moved this code into its own routine because some or
    '                   all of the properties whose values are stored in the
    '                   error list worksheet may change when a new sheet is
    '                   attached to the object by changing the value of property
    '                   ErrorListSheet.
    '
    '   References:     "The VBA Err Object," Microsoft Office XP Developer,
    '                   MSDN Library, Office Developer Center
    '                   http://msdn.microsoft.com/en-us/library/aa164768(office.10).aspx
    '
    '   Created:        Friday, 22 August 2008
    '
    '   ------------------------------------------------------------------------
    '   Revision History
    '   ------------------------------------------------------------------------
    '
    '   Date       Version Author Synopsis
    '   ---------- ------- ------ ----------------------------------------------
    '   2008/08/23 1.04    DAG/WW Move this code out of the constructor, because
    '                             the method that sets property ErrorListSheet
    '                             must also call it.
    '
    '   2008/08/30 1.07    DAG/WW Add comprehenisve error trapping.
    '
    '   2012/10/13 1.10    DAG/WW Clean up the internal documentation, adding
    '                             omitted arguments and correcting errors due to
    '                             copying of comment blocks between routines.
    '
    '   ========================================================================

    Const DEFAULT_MODULE_RNG_NM As String = "rngWWXLAppDefaultModuleName"
    Const ERROR_MESSAGE_RNG_NM As String = "rngWWXLAppExceptionList"
    Const PROJECT_NAME_RNG_NM As String = "rngWWXLAppProjectName"
    Const ORIGIN_CELL_INDEX As Long = 1
    Const STARTING_ERROR_NUMBER_RNG_NM As String = "rngWWXLAppStartingErrorNumber"

    Dim rngThisProperty As Range
    Dim strRangeName As String

    strRangeName = ERROR_MESSAGE_RNG_NM
    Set mrngWWXLAppExceptionList = mwksErrorData.Range(ERROR_MESSAGE_RNG_NM)

    strRangeName = PROJECT_NAME_RNG_NM
    Set rngThisProperty = mwksErrorData.Range(PROJECT_NAME_RNG_NM)
    mstrProjectName = rngThisProperty.Cells(ORIGIN_CELL_INDEX, ORIGIN_CELL_INDEX)
    Set rngThisProperty = Nothing

    strRangeName = DEFAULT_MODULE_RNG_NM
    Set rngThisProperty = mwksErrorData.Range(DEFAULT_MODULE_RNG_NM)
    mstrModuleName = rngThisProperty.Cells(ORIGIN_CELL_INDEX, ORIGIN_CELL_INDEX)
    Set rngThisProperty = Nothing

    strRangeName = STARTING_ERROR_NUMBER_RNG_NM
    Set rngThisProperty = mwksErrorData.Range(STARTING_ERROR_NUMBER_RNG_NM)
    mlngStartingErrorNumber = rngThisProperty.Cells(ORIGIN_CELL_INDEX, ORIGIN_CELL_INDEX)
    Set rngThisProperty = Nothing

Read_Props_From_Sheet_End:

    Exit Sub

Read_Props_From_Sheet_Err:

    If Err.Number = ERR_OBJ_DEFINED_ERROR Then
        mstrErrorMessage = _
            "Specified range name, " & strRangeName _
            & ", is undefined in worksheet " & mwksErrorData.Name
        mlngErrorNumber = ERR_NBR_RANGENAME_IS_UNDEFINED
        Err.Raise mlngErrorNumber Or vbObjectError, _
            DEFAULT_PROJECT_NAME & ".ErrorListSheet_Property", _
            mstrErrorMessage
    Else
        mlngErrorNumber = Err.Number
        mstrErrorMessage = "Error # " & Err.Number & " - " & Err.Description
        Err.Raise mlngErrorNumber, _
                  DEFAULT_PROJECT_NAME & ".WWXLAppExceptions", _
                  mstrErrorMessage
        Resume Read_Props_From_Sheet_End
    End If

End Sub
