Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: frmFilterFormWizard
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Wizard form for filter code import
' </summary>
' <remarks></remarks>
'**/
'------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_OpenModules As Long
Private m_CodeModulImporter As AppFileCodeModulTransfer

Private Const SubFormApplyFilterOption As Long = 2
Private Const FILTERFORMWIZARD_TEXT_CREATENEW As String = "<Create new>"

Private WithEvents m_L10nDict As L10nDict
Attribute m_L10nDict.VB_VarHelpID = -1

Private Property Get SourceDb() As DAO.Database
   Set SourceDb = CurrentDb
End Property

Private Property Get CodeModulImporter() As AppFileCodeModulTransfer
   If m_CodeModulImporter Is Nothing Then
      Set m_CodeModulImporter = New AppFileCodeModulTransfer
      m_CodeModulImporter.UseVbComponentsImport = APPLICATION_FILTERCODEMODULE_USEVBCOMPONENTSIMPORT
   End If
   Set CodeModulImporter = m_CodeModulImporter
End Property

Private Sub cbxFormName_AfterUpdate()

   Dim FormName As String
   Dim FormSelected As Boolean
   Dim FormRef As Form

   RefreshAddFilterButtonEnableMode

   FormName = Nz(Me.cbxFormName.Value, vbNullString)
   FormSelected = (Len(FormName) > 0)

   Me.sfrFilterControls.Enabled = FormSelected
   Me.cbxApplyFilterCtl.Enabled = FormSelected
   Me.cbxRemoveFilterCtl.Enabled = FormSelected
   Me.cbxAutoFilterCtl.Enabled = FormSelected

   Me.ApplyFilterMethodOptions.Enabled = FormSelected
   With Me.cbxApplyFilterSubForm
      .Enabled = FormSelected And (Me.ApplyFilterMethodOptions = SubFormApplyFilterOption)
      .Value = Null
   End With

   With Me.FormCodeOptions
      .Enabled = FormSelected
      Me.cbUseFilterControlTagConverter.Enabled = FormSelected And (.Value = 1)
   End With

   If Not FormSelected Then
      Me.cbxFormName.RowSource = vbNullString
      Me.sfrFilterControls.Form.InitFilterForm Nothing
      Me.cbxApplyFilterSubForm.RowSource = vbNullString
      Exit Sub
   End If

   Set FormRef = GetFilterFormRef

   LoadControlsFromForm FormRef

End Sub

Private Sub RefreshAddFilterButtonEnableMode()

   Dim EnableBtn As Boolean

   EnableBtn = (Len(Nz(Me.cbxFormName.Value, vbNullString)) > 0)

   If EnableBtn Then
      If Me.ApplyFilterMethodOptions.Value = 2 Then
         EnableBtn = (Len(Nz(Me.cbxApplyFilterSubForm.Value, vbNullString)) > 0)
      End If
   End If

   Me.cmdAddFilterCodeToForm.Enabled = EnableBtn

End Sub

Private Function GetFilterFormRef() As Form

   Dim FormName As String

   FormName = Nz(Me.cbxFormName.Value, vbNullString)
   If Len(FormName) = 0 Then
      Set GetFilterFormRef = Nothing
      Exit Function
   End If

   OpenFormIfClosed FormName
   Set GetFilterFormRef = Application.Forms(FormName)

End Function

Private Sub OpenFormIfClosed(ByVal FormName As String)
   If Not CurrentProject.AllForms(FormName).IsLoaded Then
      DoCmd.OpenForm FormName, acDesign, , , , acHidden
   End If
End Sub

Private Sub cbxFormName_GotFocus()
   LoadForms
End Sub

Private Sub LoadForms()
   Me.cbxFormName.RowSource = "Select Name FROM [" & SourceDb.Name & "].MSysObjects where Type = -32768 Order By Name"
End Sub

Private Sub cbxLangCode_AfterUpdate()
   With Me.cbxLangCode
      If IsNull(.Value) Then
         .Value = "EN"
      End If
      L10n.LangCode = .Value
   End With
End Sub

Private Sub cbxApplyFilterSubForm_AfterUpdate()
   SetFillFilterControlsFromDataSourceMode
   RefreshAddFilterButtonEnableMode
End Sub

Private Sub cbxApplyFilterSubForm_GotFocus()
   RefreshSubFormList
End Sub

Private Sub RefreshSubFormList()

   Dim frm As Form
   Dim ctl As Control

   Me.cbxApplyFilterSubForm.RowSource = vbNullString

   Set frm = GetFilterFormRef
   If frm Is Nothing Then
      Exit Sub
   End If

   For Each ctl In frm.Controls
      Select Case ctl.ControlType
         Case AcControlType.acSubform
            Me.cbxApplyFilterSubForm.AddItem ctl.Name
      End Select
   Next

End Sub

Private Sub AddFilterCodeToForm(ByVal ImporterRef As IFilterFormCodeBuilder)

   Dim FormName As String
   Dim ApplyFilterCtlName As String
   Dim RemoveFilterCtlName As String
   Dim AutoFilterCtlName As String
   Dim FilterControlTagConverterMode As Long

   FormName = Nz(Me.cbxFormName.Value, vbNullString)
   If Len(FormName) = 0 Then
      MsgBox "Bitte zuerst Formular auswählen!", , CurrentApplication.ApplicationName
      Exit Sub
   End If

   ApplyFilterCtlName = Nz(Me.cbxApplyFilterCtl.Value, vbNullString)
   RemoveFilterCtlName = Nz(Me.cbxRemoveFilterCtl.Value, vbNullString)
   AutoFilterCtlName = Nz(Me.cbxAutoFilterCtl.Value, vbNullString)

   AddChangeFilterControlsToForm FormName, ApplyFilterCtlName, RemoveFilterCtlName, AutoFilterCtlName

   If Me.FormCodeOptions.Value = 1 Then
      FilterControlTagConverterMode = Abs(Nz(Me.cbUseFilterControlTagConverter.Value, 0))
   Else
      FilterControlTagConverterMode = 0
   End If

   With New FilterFormCodeImporter

      Set .Importer = ImporterRef

      If AddControlsToImporter(.Self, FormName, FilterControlTagConverterMode <> 0) Then ' New Controls added to Header
         If FilterControlTagConverterMode = 1 Then
            FilterControlTagConverterMode = 2 'use Me.Section(acHeader).Controls
         End If
      End If

      .WriteToForm FormName, _
                   Nz(Me.ApplyFilterMethodOptions.Value, 0), _
                   Nz(Me.cbxApplyFilterSubForm.Value, vbNullString), _
                   Nz(Me.cbxSqlLang.Column(1), vbNullString), _
                   Nz(Me.cbxSqlLang.Column(2), vbNullString), _
                   Nz(Me.cbxSqlLang.Column(3), vbNullString), _
                   GetCheckedControlName(ApplyFilterCtlName), _
                   GetCheckedControlName(RemoveFilterCtlName), _
                   GetCheckedControlName(AutoFilterCtlName), _
                   FilterControlTagConverterMode

   End With

   CheckModules

   Me.SetFocus

End Sub

Private Property Get FilterFormWizardText_CreateNew() As String
   FilterFormWizardText_CreateNew = L10n.Text(FILTERFORMWIZARD_TEXT_CREATENEW)
End Property


Private Sub AddChangeFilterControlsToForm(ByVal FilterFormName As String, ByRef ApplyFilterCtlName As String, ByRef RemoveFilterCtlName As String, ByRef AutoFilterCtlName As String)

   Dim SaveForm As Boolean

   If ApplyFilterCtlName = FilterFormWizardText_CreateNew Then
      ApplyFilterCtlName = "cmdApplyFilter"
      AddChangeFilterControlToForm FilterFormName, ApplyFilterCtlName, acCommandButton, L10n.Text("Filter anwenden")
      SaveForm = True
   End If

   If RemoveFilterCtlName = FilterFormWizardText_CreateNew Then
      RemoveFilterCtlName = "cmdRemoveFilter"
      AddChangeFilterControlToForm FilterFormName, RemoveFilterCtlName, acCommandButton, L10n.Text("Filter löschen")
      SaveForm = True
   End If

   If AutoFilterCtlName = FilterFormWizardText_CreateNew Then
      AutoFilterCtlName = "cbAutoFilter"
      AddChangeFilterControlToForm FilterFormName, AutoFilterCtlName, acCheckBox, L10n.Text("Autofilter"), True
      SaveForm = True
   End If

   If SaveForm Then
      DoCmd.Save acForm, FilterFormName
   End If

End Sub

Private Sub AddChangeFilterControlToForm(ByVal FilterFormName As String, _
                                         ByRef ControlName As String, _
                                         ByVal ControlType As AcControlType, _
                                         ByVal ControlCaption As String, _
                                Optional ByVal DefaultValue As Variant)

   Const DistanceBetweenControls As Long = 72

   Dim InsertLabelControl As Boolean

   Select Case ControlType
      Case AcControlType.acCommandButton, AcControlType.acOptionButton, AcControlType.acLabel
         InsertLabelControl = False
      Case Else
         InsertLabelControl = True
   End Select

   With New FormDesigner
      .FormName = FilterFormName
      With .AddControl(ControlType, ControlName, acHeader, , , , , DistanceBetweenControls, InsertLabelControl, ControlCaption)
         If Not InsertLabelControl Then
            .Caption = ControlCaption
         End If
         ControlName = .Name
         If Not IsMissing(DefaultValue) Then
            .DefaultValue = DefaultValue
         End If
      End With
   End With

End Sub

Private Function AddControlsToImporter(ByVal Importer As FilterFormCodeImporter, ByVal FilterFormName As String, ByVal UseFilterDefinitionInTagProperty As Boolean) As Boolean
   AddControlsToImporter = Me.sfrFilterControls.Form.AddControlsToImporter(Importer, FilterFormName, UseFilterDefinitionInTagProperty)
End Function

Private Sub cmdAddFilterCodeToForm_Click()

   Select Case Me.FormCodeOptions.Value
      Case 1
         AddFilterCodeToForm New FilterControlManagerCodeBuilder
      Case 2
         AddFilterCodeToForm New FilterStringBuilderCodeBuilder
   End Select

End Sub

Private Sub cmdFillFilterControlsFromDataSource_Click()

   Dim FilterFormName As String
   Dim SubFormName As String

   FilterFormName = Nz(Me.cbxFormName.Value, vbNullString)
   If Len(FilterFormName) = 0 Then Exit Sub

   Select Case Me.ApplyFilterMethodOptions.Value
      Case 1
         ' keine weiteren Anpassungen
      Case 2
         SubFormName = Nz(Me.cbxApplyFilterSubForm.Value, vbNullString)
         If Len(SubFormName) = 0 Then Exit Sub
      Case Else
         Exit Sub
   End Select

   If Not CurrentProject.AllForms(FilterFormName).IsLoaded Then
      DoCmd.OpenForm FilterFormName, acDesign, , , , acWindowNormal
   End If

   With Application.Forms(FilterFormName)
      Select Case Me.ApplyFilterMethodOptions.Value
         Case 1
            FillFilterControlsFromDataSource .RecordSource
         Case 2
            FillFilterControlsFromDataSource .Controls(SubFormName).Form.RecordSource
      End Select

   End With

End Sub

Private Sub FillFilterControlsFromDataSource(ByVal DataSource As String)

   Dim fld As DAO.Field

   If Len(DataSource) = 0 Then
      MsgBox "Es ist keine Datenquelle eingestellt!", vbExclamation, "Filter aus Datenfelder erzeugen"
      Exit Sub
   End If

   With CurrentDb.OpenRecordset(DataSource)
      For Each fld In .Fields
         AddFilterControlFromDataField fld
      Next
      .Close
   End With

End Sub

Private Sub AddFilterControlFromDataField(ByVal DataField As DAO.Field)
   Me.sfrFilterControls.Form.InsertFilterRecord DataField.Name, DataField.Type
End Sub

Private Sub Form_Current()
   CheckModules True, True
End Sub

Private Sub OpenReplaceModulesContextMenu(ByRef MenuControlRef As Control, ByVal X As Single, ByVal Y As Single)

   Dim mnu As WinApiShortcutMenu
   Dim SelectedMenuItem As Long

   Set mnu = New WinApiShortcutMenu

   With mnu
      .ControlSection = acDetail
      Set .AccessForm = Me
      Set .MenuControl = MenuControlRef

      .AddMenuItem 3, L10n.Text("Update code modules in the add-in from the SCC repository (master).")
      .AddMenuItem 4, L10n.Text("Update code modules in the add-in from the SCC repository (draft).")
      .AddMenuItem 0, "", MF_SEPARATOR
      .AddMenuItem 1, L10n.Text("Replace all code modules (existing modules are overwritten)")
      .AddMenuItem 2, L10n.Text("Replace all code modules except 'SqlTools' (SqlTools remains unchanged)")

   End With

   SelectedMenuItem = mnu.OpenMenu(X, Y)
   Select Case SelectedMenuItem
      Case 1
         ReplaceAllModules True
      Case 2
         ReplaceAllModules False
      Case 3
         RefreshCodeModulesFromSccRepo False
      Case 4
         RefreshCodeModulesFromSccRepo True
      Case Else
         '
   End Select

   Set mnu = Nothing

End Sub

Private Sub ReplaceAllModules(Optional ByVal WithSqlTools As Boolean = False)

   Dim ModulesReplaced As Boolean

On Error Resume Next

   With CodeModulImporter
      If WithSqlTools Then
         ModulesReplaced = .ReplaceCodeModules("SqlTools", "StringCollection", "FilterStringBuilder", "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager", "FilterControlTagConverter")
      Else
         ModulesReplaced = .ReplaceCodeModules("StringCollection", "FilterStringBuilder", "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager", "FilterControlTagConverter")
      End If
   End With

   If ModulesReplaced Then

      SetModulCopyControls "CopyModulSqlTools", True
      SetModulCopyControls "CopyModulFilterStringBuilder", True
      SetModulCopyControls "CopyModulFilterControlManager", True
      SetModulCopyControls "CopyModules", True

      If m_OpenModules = 0 Then
         MsgBox "Codemodule wurden aktualisiert.", , CurrentApplication.ApplicationName
      End If

      m_OpenModules = 0

   End If

End Sub

Private Sub Form_Open(ByRef Cancel As Integer)
   Me.cbxLangCode.Value = L10n.LangCode
   L10n.TranslateControls Me.Controls
   Set m_L10nDict = L10n
   SetFormCaption
End Sub

Private Sub SetFormCaption()
   Me.Caption = "ACLib FilterForm Wizard  " & ChrW(&H25AA) & "  Version " & CurrentApplication.Version & " (" & L10n.Text("Filter Classes") & ": SCC-Rev. " & SccRev & ")"
End Sub

Private Sub FormCodeOptions_AfterUpdate()

   With Me.cbUseFilterControlTagConverter
      .Enabled = (Me.FormCodeOptions.Value = 1)
      If Not .Enabled Then
         .Value = False
      End If
   End With

End Sub

Private Sub labCheckVersion_Click()

   Select Case L10n.ControlTag(Me.labCheckVersion.Tag)
      Case "unchecked"
         SetNewVersionInfo
      Case "NewVersionExists"
         OpenDownloadSource
      Case Else
         '
   End Select

End Sub

Private Sub labCheckVersion_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   If L10n.ControlTag(Me.labCheckVersion.Tag) <> "VersionUptodate" Then
      modWinApi_Mouse.MouseCursor IDC_HAND
   End If
End Sub

Private Sub labCopyCaption_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

   If Button = 1 Or Button = 2 Then
      OpenReplaceModulesContextMenu Me.labCopyCaption, X, Y
   End If

End Sub

Private Sub CheckModules(Optional ByVal WithControlManager As Boolean = True, Optional ByVal WithTagConverter As Boolean = False)

   Dim CheckOk As Boolean

On Error Resume Next

   m_OpenModules = 0

   With CodeModulImporter
      CheckOk = .CheckCodeModules("SqlTools")
      m_OpenModules = m_OpenModules + 1 + CheckOk
      SetModulCopyControls "CopyModulSqlTools", CheckOk

      CheckOk = .CheckCodeModules("StringCollection", "FilterStringBuilder")
      m_OpenModules = m_OpenModules + 1 + CheckOk
      SetModulCopyControls "CopyModulFilterStringBuilder", CheckOk

      If WithControlManager Then
         CheckOk = .CheckCodeModules("FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager")
         m_OpenModules = m_OpenModules + 1 + CheckOk
         SetModulCopyControls "CopyModulFilterControlManager", CheckOk

         If WithTagConverter Then
            CheckOk = .CheckCodeModules("FilterControlTagConverter")
            m_OpenModules = m_OpenModules + 1 + CheckOk
            SetModulCopyControls "CopyModulFilterControlTagConverter", CheckOk
         End If

      End If

      CheckOk = (m_OpenModules = 0)
      SetModulCopyControls "CopyModules", CheckOk
   End With

End Sub

Private Sub cmdCopyModulSqlTools_Click()

   CodeModulImporter.TransferCodeModules "SqlTools"
   m_OpenModules = m_OpenModules - 1
   SetModulCopyControls "CopyModulSqlTools", True

End Sub

Private Sub cmdCopyModulFilterControlManager_Click()

   With CodeModulImporter
      .TransferCodeModules "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager"
      m_OpenModules = m_OpenModules - 1
      If Not .CheckCodeModules("StringCollection", "FilterStringBuilder") Then
         cmdCopyModulFilterStringBuilder_Click
      End If
      SetModulCopyControls "CopyModulFilterControlManager", True
   End With

End Sub

Private Sub cmdCopyModulFilterControlTagConverter_Click()

   With CodeModulImporter
      .TransferCodeModules "FilterControlTagConverter"
      m_OpenModules = m_OpenModules - 1
      If Not .CheckCodeModules("FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager") Then
         cmdCopyModulFilterControlManager_Click
      End If
      SetModulCopyControls "CopyModulFilterControlTagConverter", True
   End With

End Sub

Private Sub cmdCopyModulFilterStringBuilder_Click()

   With CodeModulImporter
      .TransferCodeModules "StringCollection", "FilterStringBuilder"
      m_OpenModules = m_OpenModules - 1
      If Not .CheckCodeModules("SqlTools") Then
         cmdCopyModulSqlTools_Click
      End If
      SetModulCopyControls "CopyModulFilterStringBuilder", True
   End With

End Sub

Private Sub cmdCopyModules_Click()

   CodeModulImporter.TransferCodeModules "SqlTools", "StringCollection", "FilterStringBuilder", "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager", "FilterControlTagConverter"
   SetModulCopyControls "CopyModulSqlTools", True
   SetModulCopyControls "CopyModulFilterStringBuilder", True
   SetModulCopyControls "CopyModulFilterControlManager", True
   SetModulCopyControls "CopyModulFilterControlTagConverter", True
   SetModulCopyControls "CopyModules", True
   m_OpenModules = 0

End Sub

Private Sub SetModulCopyControls(ByVal ModulCodeName As String, ByVal Exists As Boolean)

   Dim ModuleCodeCaption As String
   Dim BorderStyle As Long

On Error Resume Next

   If Exists Then Me.sysFirst.SetFocus
   Me.Controls("cmd" & ModulCodeName).Enabled = Not Exists
   If Exists Then
      ModuleCodeCaption = "b"
      BorderStyle = 0
   Else
      ModuleCodeCaption = vbNullString
      BorderStyle = 1
   End If
   With Me.Controls("lab" & ModulCodeName)
      .Caption = ModuleCodeCaption
      .BorderStyle = BorderStyle
   End With

   If ModulCodeName <> "CopyModules" And m_OpenModules = 0 Then
      SetModulCopyControls "CopyModules", True
   End If

End Sub

Private Sub labCopyCaption_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   modWinApi_Mouse.MouseCursor IDC_HAND
End Sub

Private Sub m_L10nDict_LanguageChanged()
   If Nz(Me.cbxLangCode.Value, vbNullString) <> L10n.LangCode Then
      Me.cbxLangCode.Value = L10n.LangCode
   End If
   SetFormCaption
   L10n.TranslateControls Me.Controls

   If Len(Me.cbxFormName.Value) > 0 Then
      LoadControlsFromForm GetFilterFormRef
   End If
End Sub

Private Sub ApplyFilterMethodOptions_AfterUpdate()

   With Me.cbxApplyFilterSubForm
      .Enabled = (Me.ApplyFilterMethodOptions = SubFormApplyFilterOption)
      If .Enabled = False Then .Value = Null
   End With

   SetFillFilterControlsFromDataSourceMode

   RefreshAddFilterButtonEnableMode

End Sub

Private Sub SetFillFilterControlsFromDataSourceMode()

   With Me.cmdFillFilterControlsFromDataSource
      Select Case Me.ApplyFilterMethodOptions
         Case 1
            .Enabled = True
         Case 2
            .Enabled = (Len(Nz(Me.cbxApplyFilterSubForm.Value, vbNullString)) > 0)
         Case Else
            .Enabled = False
      End Select

      If .Enabled Then
         .ControlTipText = vbNullString
      Else
         .ControlTipText = L10n.Text("Please set the data form in the 'ApplyFilter method' option first!")
      End If
   End With

End Sub

Private Sub LoadControlsFromForm(ByVal FormRef As Form)

   Dim ctl As Control

   ClearFilterCtlCombobox Me.cbxApplyFilterCtl
   ClearFilterCtlCombobox Me.cbxRemoveFilterCtl
   ClearFilterCtlCombobox Me.cbxAutoFilterCtl

   If FormRef Is Nothing Then
      Exit Sub
   End If

   Me.cbxApplyFilterCtl.AddItem FilterFormWizardText_CreateNew
   Me.cbxRemoveFilterCtl.AddItem FilterFormWizardText_CreateNew
   Me.cbxAutoFilterCtl.AddItem FilterFormWizardText_CreateNew

   For Each ctl In FormRef.Controls
      Select Case ctl.ControlType
         Case AcControlType.acCommandButton

            Me.cbxApplyFilterCtl.AddItem ctl.Name
            Me.cbxRemoveFilterCtl.AddItem ctl.Name

         Case AcControlType.acCheckBox, AcControlType.acToggleButton

            Me.cbxAutoFilterCtl.AddItem ctl.Name

      End Select
   Next

   Me.sfrFilterControls.Form.InitFilterForm FormRef

End Sub

Private Sub ClearFilterCtlCombobox(ByVal CbxRef As ComboBox)
   With CbxRef
      .RowSource = vbNullString
      .Value = Null
   End With
End Sub

Private Sub SetNewVersionInfo()

   Dim AhProcParams(1) As Variant
   Dim AhResumeMode As ApplicationHandlerResumeModes
   Dim AhResumeMessage As Variant

   Call CurrentApplication.CallExtensionProcedure("AppVersion", "NewVersionExists", AhResumeMode, AhResumeMessage, AhProcParams)

   If AhResumeMode <> AppResumeMode_Completed Then
      If AhResumeMode = AppResumeMode_Error Then
      If Len(AhResumeMessage) > 0 Then
         MsgBox AhResumeMessage, vbExclamation, "Error during version check"
      End If
      End If
      Exit Sub
   End If

   With Me.labCheckVersion
   If AhResumeMessage = False Then
      .Caption = L10n.Text("Current version is installed")
      .Tag = "VersionUptodate" '& L10n.LangSeparatorChar '& "LANG:Wizard-Version prüfen"
      .ForeColor = 0
      .FontUnderline = False
   ElseIf AhResumeMessage = True Then
      .Caption = FormatText("Neue Version ({0}) verfügbar", AhProcParams(0))
      .ControlTipText = AhProcParams(1)
      .Tag = "NewVersionExists" '& L10n.LangSeparatorChar & "LANG:Wizard-Version prüfen"
   Else

   End If
   End With

End Sub

Private Sub OpenDownloadSource()
   Application.FollowHyperlink APPLICATION_DOWNLOADSOURCE, , True, False
End Sub

Private Function GetCheckedControlName(ByVal ControlNameToCheck As String) As String
   GetCheckedControlName = Replace(ControlNameToCheck, " ", "_")
End Function

Private Sub RefreshCodeModulesFromSccRepo(ByVal UseDraft As Boolean)

   With New ACLibGitHubImporter

      If UseDraft Then
         .BranchName = "draft"
      End If

      If .RevisionString(True) = SccRevMin Then
         If MsgBox("All modules are already up to date." & vbNewLine & "Still re-insert?", vbYesNo + vbDefaultButton2, CurrentApplication.ApplicationName) <> vbYes Then
            Exit Sub
         End If
      End If

      .UpdateCodeModules

   End With

   MsgBox "Code has been updated", , CurrentApplication.ApplicationName

   SetFormCaption

End Sub
