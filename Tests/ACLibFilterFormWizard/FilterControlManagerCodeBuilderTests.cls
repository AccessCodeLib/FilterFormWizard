VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControlManagerCodeBuilderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private m_OriginalLangCode As String

'--------------------------------------------------------------------
' Test Preparation / Cleanup
'--------------------------------------------------------------------
Public Sub Setup()
   m_OriginalLangCode = L10n.LangCode
End Sub

Public Sub TearDown()
   L10n.LangCode = m_OriginalLangCode
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row("EN", "' show 'apply filter' is required").Name("EN apply filter signal")
'AccUnit:Row("EN", "'or with checkbox: = Nz(").Name("EN auto filter checkbox")
'AccUnit:Row("EN", "' 1=0 .. Don't show records").Name("EN remove filter 1=0")
'AccUnit:Row("EN", "' without parameter transfer all records are displayed").Name("EN remove filter all")
'AccUnit:Row("DE", "' Notwendigen Klick auf ""Filter anwenden"" signalisieren").Name("DE apply filter signal")
'AccUnit:Row("DE", "'oder z. B. mittels Checkbox: = Nz(").Name("DE auto filter checkbox")
'AccUnit:Row("DE", "' 1=0 .. Anzeige leeren").Name("DE remove filter 1=0")
'AccUnit:Row("DE", "' ohne Parameterübergabe werden alle Datensätze angezeigt").Name("DE remove filter all")
Public Sub ClassInitalisationCode_LangCode_ContainsLocalizedComment(ByVal LangCode As String, ByVal ExpectedComment As String)
   ' Arrange
   L10n.LangCode = LangCode
   ' Act
   Dim Actual As String
   Actual = BuildClassInitalisationCode()
   ' Assert
   Assert.IsTrue InStr(1, Actual, ExpectedComment, vbBinaryCompare) > 0, "Missing: " & ExpectedComment & vbNewLine & Actual
End Sub

'AccUnit:Row("Notwendigen").Name("apply filter signal")
'AccUnit:Row("oder z. B.").Name("auto filter checkbox")
'AccUnit:Row("Datensätze").Name("remove filter")
'AccUnit:Row("Parameterübergabe").Name("remove filter all")
Public Sub ClassInitalisationCode_LangCodeEN_ContainsNoGermanText(ByVal GermanText As String)
   ' Arrange
   L10n.LangCode = "EN"
   ' Act
   Dim Actual As String
   Actual = BuildClassInitalisationCode()
   ' Assert
   Assert.IsTrue InStr(1, Actual, GermanText, vbTextCompare) = 0, "Unexpected: " & GermanText & vbNewLine & Actual
End Sub

Private Function BuildClassInitalisationCode() As String
   Dim Builder As IFilterFormCodeBuilder
   Set Builder = New FilterControlManagerCodeBuilder
   BuildClassInitalisationCode = Builder.ClassInitalisationCode( _
                                    "Private Sub ApplyFilter(ByVal NewFilterString As String)" & vbNewLine & "End Sub", _
                                    vbNullString, New StringCollection, _
                                    ApplyFilterCtlName:="cmdApplyFilter", _
                                    RemoveFilterCtlName:="cmdRemoveFilter")
End Function
