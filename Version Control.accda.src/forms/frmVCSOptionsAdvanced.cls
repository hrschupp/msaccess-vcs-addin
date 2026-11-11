Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsAdvanced
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Advanced options section: debugging, hashing, export tweaks, and logging.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load option values into controls on initial display.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim opt As clsOptions
    Dim intCILevel As eConnIdentity

    Set opt = Form_frmVCSOptions.OptionsInstance

    ' Populate connection identity string combo
    With Me.cboConnectionIdentity
        .RowSource = vbNullString
        For intCILevel = eciDbUnderscore To eConnIdentity.[_Last] - 1
            .AddItem intCILevel & ";" & opt.GetConnectionIdentityMethodName(intCILevel)
        Next intCILevel
        .ControlTipText = T("The identity defaults to the DB name in version 5, non-ASCII characters are replaced by underscores. " & _
            "It's possible to extend this identity to the server name und use escape sequences to avoid clashes. Or use the hashing.")
    End With

    LoadOptions
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IOptionsSection_LoadOptions / IOptionsSection_SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : IOptionsSection interface implementation.
'---------------------------------------------------------------------------------------
'
Private Sub IOptionsSection_LoadOptions()
    LoadOptions
End Sub

Private Sub IOptionsSection_SaveOptions()
    SaveOptions
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load option values from the working options instance into form controls.
'---------------------------------------------------------------------------------------
'
Private Sub LoadOptions()

    With Form_frmVCSOptions.OptionsInstance
        ' Debugging
        Me.chkShowDebug = .ShowDebug
        Me.chkBreakOnError = .BreakOnError
        Me.chkShowVCSLegacy = .ShowVCSLegacy

        ' Hashing
        Me.cboHashAlgorithm = .HashAlgorithm
        Me.chkUseShortHash = .UseShortHash

        ' Export Tweaks
        Me.chkStripPublishOption = .StripPublishOption
        Me.chkPreserveRubberDuckID = .PreserveRubberDuckID
        Me.chkSaveAllDocumentProperties = .SaveAllDocumentProperties
        Me.chkSaveLinkedTableDef = .SaveLinkedTableDef
        Me.cboConnectionIdentity = .ConnectionIdentity

        ' System
        Me.chkUseGitIntegration = .UseGitIntegration

        ' Logging
        Me.txtMaxLogFiles = .MaxLogFiles

        ' Automated Testing
        Me.txtDefaultTestFilter = .DefaultTestFilter
        Me.chkUseWebTestRunner = .UseWebTestRunner
        Me.chkExportTestResultsJUnit = .ExportTestResultsJUnit
        Me.chkExportTestResultsHtml = .ExportTestResultsHtml
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Save form control values back to the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()

    With Form_frmVCSOptions.OptionsInstance
        ' Debugging
        .ShowDebug = Nz(Me.chkShowDebug)
        .BreakOnError = Nz(Me.chkBreakOnError)
        .ShowVCSLegacy = Nz(Me.chkShowVCSLegacy)

        ' Hashing
        .HashAlgorithm = Nz(Me.cboHashAlgorithm)
        .UseShortHash = Nz(Me.chkUseShortHash)

        ' Export Tweaks
        .StripPublishOption = Nz(Me.chkStripPublishOption)
        .PreserveRubberDuckID = Nz(Me.chkPreserveRubberDuckID)
        .SaveAllDocumentProperties = Nz(Me.chkSaveAllDocumentProperties)
        .SaveLinkedTableDef = Nz(Me.chkSaveLinkedTableDef)
        .ConnectionIdentity = Nz(Me.cboConnectionIdentity, 0)

        ' System
        .UseGitIntegration = Nz(Me.chkUseGitIntegration, 0)

        ' Logging
        .MaxLogFiles = Nz(Me.txtMaxLogFiles, 0)

        ' Automated Testing
        .DefaultTestFilter = Nz(Me.txtDefaultTestFilter, vbNullString)
        .UseWebTestRunner = Nz(Me.chkUseWebTestRunner, 0)
        .ExportTestResultsJUnit = Nz(Me.chkExportTestResultsJUnit, 0)
        .ExportTestResultsHtml = Nz(Me.chkExportTestResultsHtml, 0)
    End With

End Sub
