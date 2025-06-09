Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Form: frmFilterFormWizard_SF_Controls
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Wizard form subform for FilterControls import
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const ADODB_EARLYBINDING = 0

#If ADODB_EARLYBINDING = 1 Then

#Else
   Public Enum CursorTypeEnum
      adOpenKeyset = 1
   End Enum

   Public Enum CursorLocationEnum
      adUseClient = 3
   End Enum

   Public Enum LockTypeEnum
      adLockOptimistic = 3
   End Enum
#End If

Private WithEvents m_LngDict As L10nDict
Attribute m_LngDict.VB_VarHelpID = -1

Public Sub InitFilterForm(ByVal FrmRef As Form)

   InitDataSource
   LoadControlsFromForm FrmRef

End Sub

Private Sub InitDataSource()

#If ADODB_EARLYBINDING = 1 Then
   Dim rst As ADODB.Recordset
#Else
   Dim rst As Object
#End If

   Set rst = CreateObject("ADODB.Recordset")
   rst.CursorLocation = adUseClient
   rst.Open "select * from tabFilterControls", CodeProject.Connection, adOpenKeyset, adLockOptimistic
   Set rst.ActiveConnection = Nothing

   Set Me.Recordset = rst

End Sub

Private Sub LoadControlsFromForm(ByVal FrmRef As Form)

   Dim ctl As Control

   Me.fcControl.RowSource = vbNullString

   If FrmRef Is Nothing Then
      Exit Sub
   End If

   For Each ctl In FrmRef.Controls
      Select Case ctl.ControlType
         Case AcControlType.acTextBox, AcControlType.acComboBox, _
              AcControlType.acCheckBox, AcControlType.acListBox, _
              AcControlType.acOptionButton, AcControlType.acOptionGroup, _
              AcControlType.acToggleButton

            Me.fcControl.AddItem ctl.Name

      End Select
   Next

End Sub

Public Function AddControlsToImporter(ByVal Importer As FilterFormCodeImporter, ByVal FilterFormName As String, ByVal UseFilterDefinitionInTagProperty As Boolean) As Boolean

#If ADODB_EARLYBINDING = 1 Then
   Dim rst As ADODB.Recordset
#Else
   Dim rst As Object
#End If

   Dim RelOp As String
   Dim UseBetweenRelOp As Boolean
   Dim FilterCtl As Control
   Dim FrmDsgnr As FormDesigner

   Set rst = Me.Recordset.Clone
   If rst.RecordCount = 0 Then
      Exit Function
   End If

   Dim DataFieldName As String
   Dim ControlName As String, ControlName2 As String

   Dim SaveForm As Boolean

   If UseFilterDefinitionInTagProperty Then
      Set FrmDsgnr = New FormDesigner
      FrmDsgnr.FormName = FilterFormName
   End If

   With rst
      Do While Not .EOF

         RelOp = Nz(.Fields("RelationalOperator"), vbNullString)
         UseBetweenRelOp = (RelOp = "SQL_Between")

         If UseFilterDefinitionInTagProperty Then
            RelOp = GetRelationalOperatorKey(RelOp)
            If .Fields("WildCardSuffix") = True Then
               RelOp = RelOp & "*"
            End If
            If .Fields("WildCardPrefix") = True Then
               RelOp = "*" & RelOp
            End If
            If .Fields("RelationalOperatorNot") = True Then
               RelOp = "Not " & RelOp
            End If
         Else
            If .Fields("WildCardSuffix") = True Then
               RelOp = RelOp & " + SQL_Add_WildCardSuffix"
            End If
            If .Fields("WildCardPrefix") = True Then
               RelOp = RelOp & " + SQL_Add_WildCardPrefix"
            End If
            If .Fields("RelationalOperatorNot") = True Then
               RelOp = "SQL_Not + " & RelOp
            End If
         End If

         DataFieldName = Nz(.Fields("DataField").Value, vbNullString)

         If .Fields("CreateControl").Value = True Then
            Set FilterCtl = InsertFormControl(FilterFormName, Nz(.Fields("ControlType").Value, acTextBox), DataFieldName, UseBetweenRelOp, ControlName, ControlName2, UseFilterDefinitionInTagProperty)
            SaveForm = True
            AddControlsToImporter = True
         Else
            ControlName = Nz(.Fields("Control").Value, vbNullString)
            ControlName2 = Nz(.Fields("Control2"), vbNullString)
            If UseFilterDefinitionInTagProperty Then
               Set FilterCtl = FrmDsgnr.Form.Controls(ControlName)
            End If
         End If

         If UseFilterDefinitionInTagProperty Then
            FilterCtl.Tag = GetFilterControlTag(Nz(.Fields("DataType").Value, vbNullString), RelOp, DataFieldName, ControlName2)
            SaveForm = True
         Else
            ControlName = GetCheckedControlName(ControlName)
            ControlName2 = GetCheckedControlName(ControlName2)
            Importer.AddFilterControlDefinition DataFieldName, _
                                                Nz(.Fields("DataType"), vbNullString), _
                                                RelOp, _
                                                ControlName, _
                                                ControlName2
         End If

         Set FilterCtl = Nothing
         ControlName = vbNullString
         ControlName2 = vbNullString
         .MoveNext

      Loop
   End With

   If UseFilterDefinitionInTagProperty Then
      Set FrmDsgnr = Nothing
   End If

   If SaveForm Then
      DoCmd.Save acForm, FilterFormName
   End If

End Function

Private Function GetFilterControlTag(ByVal DataTypeText As String, ByVal RelOpText As String, ByVal DataFieldName As String, ByVal ControlName2 As String) As String

   Dim ControlTag As String
   Const SeparatorChar As String = "|"

   ControlTag = Replace(DataTypeText, "SQL_", vbNullString) & SeparatorChar & RelOpText & SeparatorChar & DataFieldName & SeparatorChar & ControlName2

   Do While Right(ControlTag, 1) = SeparatorChar
      ControlTag = Left(ControlTag, Len(ControlTag) - 1)
   Loop

   GetFilterControlTag = ControlTag

End Function

Private Function GetRelationalOperatorKey(ByVal RelationalOperatorCode As String) As String
   With CodeProject.Connection.Execute("select RelationalOperator from tabRelationalOperators where RelationalOperatorCode = '" & Replace(RelationalOperatorCode, "'", "''") & "'")
      If Not .EOF Then
         GetRelationalOperatorKey = .Fields(0).Value
      End If
      .Close
   End With
End Function

Private Function InsertFormControl(ByVal FilterFormName As String, ByVal ControlType As AcControlType, ByVal DataFieldName As String, ByVal UseBetweenRelOp As Boolean, ByRef ControlName As String, ByRef Control2Name As String, ByVal UseFilterDefinitionInTagProperty As Boolean) As Control

   Const DistanceBetweenControls As Long = 72

   Dim FilterCtl As Control
   Dim LabCtl As Control

   Dim ControlCaption As String, Control2Caption As String
   Dim CreateControl2 As Boolean

   ControlCaption = DataFieldName
   If StrComp(Left(DataFieldName, 1), UCase(Left(DataFieldName, 1)), vbBinaryCompare) <> 0 Then
      DataFieldName = UCase(Left(DataFieldName, 1)) & Mid(DataFieldName, 2)
   End If
   ControlName = GetCheckedControlName("fctl" & DataFieldName)

   With New FormDesigner

      .FormName = FilterFormName

      If UseBetweenRelOp Then
         CreateControl2 = True
         Control2Caption = ControlCaption & " bis"
         If UseFilterDefinitionInTagProperty Then
            Control2Name = "x" & ControlName & "Max"
         Else
            Control2Name = ControlName & "Max"
            ControlName = ControlName & "Min"
         End If
         ControlCaption = ControlCaption & " von"
      End If

      Set FilterCtl = .AddControl(ControlType, ControlName, acHeader, , , , , DistanceBetweenControls, True, ControlCaption)
      ControlName = FilterCtl.Name
      If CreateControl2 Then
         Set LabCtl = FilterCtl.Controls(0)
         Control2Name = .AddControl(ControlType, Control2Name, acHeader, LabCtl.Left + LabCtl.Width + DistanceBetweenControls, LabCtl.Top, , , DistanceBetweenControls, True, Control2Caption).Name
      End If

   End With

   Set InsertFormControl = FilterCtl

End Function

Private Sub cbCreateControl_AfterUpdate()

   Dim CreateNewControl As Boolean

   CreateNewControl = Me.cbCreateControl.Value

   With Me.ogNewControlType
      .TabStop = Nz(.Value, False)
      If CreateNewControl Then
         .Value = acTextBox
      Else
         .Value = Null
      End If

   End With

End Sub

Private Sub Form_Current()
   Me.ogNewControlType.TabStop = Nz(Me.cbCreateControl.Value)
End Sub

Private Sub Form_Load()
   L10n.TranslateControls Me.Controls
   Set m_LngDict = L10n
End Sub

Private Sub m_LngDict_DictionaryRefreshed()
   L10n.TranslateControls Me.Controls
End Sub

Private Sub ogNewControlType_Enter()
   Me.ogNewControlType.Locked = (Me.cbCreateControl.Value = False)
End Sub

Private Function GetCheckedControlName(ByVal ControlNameToCheck As String) As String
   GetCheckedControlName = Replace(ControlNameToCheck, " ", "_")
End Function

Public Sub InsertFilterRecord(ByVal DataFieldName As String, ByVal DataFieldType As DAO.DataTypeEnum)

   Dim FieldDataType As SqlFieldDataType

   Select Case DataFieldType
      Case DAO.DataTypeEnum.dbBinary, DAO.DataTypeEnum.dbGUID, DAO.DataTypeEnum.dbVarBinary
         Exit Sub
      Case 101 ' 101 = DAO.DataTypeEnum.dbAttachment
         Exit Sub
   End Select

   FieldDataType = ConvertDaoDataTypeToSqlFieldDataType(DataFieldType)

   With Me.Recordset
      .AddNew
      .Fields("DataField").Value = DataFieldName
      .Fields("DataType").Value = ConvertSqlFieldDataTypeToString(FieldDataType)
      If FieldDataType = SQL_Text Then
         .Fields("RelationalOperator").Value = "SQL_Like"
         .Fields("WildCardSuffix").Value = True
      Else
         .Fields("RelationalOperator").Value = "SQL_Equal"
         .Fields("WildCardSuffix").Value = False
      End If
      .Fields("CreateControl").Value = True
      .Fields("ControlType").Value = acTextBox
      .Fields("WildCardPrefix").Value = False
      .Fields("RelationalOperatorNot").Value = False
      .Update
   End With

End Sub

Private Function ConvertDaoDataTypeToSqlFieldDataType(ByVal DT As DAO.DataTypeEnum) As SqlFieldDataType

   Select Case DT
      Case DAO.DataTypeEnum.dbBigInt, DAO.DataTypeEnum.dbByte, DAO.DataTypeEnum.dbDecimal, DAO.DataTypeEnum.dbCurrency, DAO.DataTypeEnum.dbDouble, DAO.DataTypeEnum.dbFloat, DAO.DataTypeEnum.dbInteger, DAO.DataTypeEnum.dbLong, DAO.DataTypeEnum.dbNumeric, DAO.DataTypeEnum.dbSingle
         ConvertDaoDataTypeToSqlFieldDataType = SqlFieldDataType.SQL_Numeric
      Case DAO.DataTypeEnum.dbDate, DAO.DataTypeEnum.dbTime
         ConvertDaoDataTypeToSqlFieldDataType = SqlFieldDataType.SQL_Date
      Case DAO.DataTypeEnum.dbBoolean
         ConvertDaoDataTypeToSqlFieldDataType = SqlFieldDataType.SQL_Boolean
      Case Else
         ConvertDaoDataTypeToSqlFieldDataType = SqlFieldDataType.SQL_Text
   End Select

End Function

Private Function ConvertSqlFieldDataTypeToString(ByVal DT As SqlFieldDataType) As String
   Select Case DT
      Case SqlFieldDataType.SQL_Text
         ConvertSqlFieldDataTypeToString = "SQL_Text"
      Case SqlFieldDataType.SQL_Numeric
         ConvertSqlFieldDataTypeToString = "SQL_Numeric"
      Case SqlFieldDataType.SQL_Date
         ConvertSqlFieldDataTypeToString = "SQL_Date"
      Case SqlFieldDataType.SQL_Boolean
         ConvertSqlFieldDataTypeToString = "SQL_Boolean"
   End Select
End Function
