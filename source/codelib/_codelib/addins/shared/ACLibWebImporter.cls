VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACLibWebImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ACLibWebImporter
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsklasse zum Importieren von Codemodulen aus der Code-Bibliothek in die usys_Appfiles-Tabelle
' </summary>
' <remarks>
' </remarks>
'\ingroup addins_shared
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/ACLibWebImporter.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Const SvnBaseUrl As String = "https://svn.access-codelib.net/svn/codelib/"

Private m_SvnRev As String
Private m_UseDraftBranch As Boolean

#If VBA7 Then
Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Public Property Get UseDraftBranch() As Boolean
   UseDraftBranch = m_UseDraftBranch
End Property

Public Property Let UseDraftBranch(ByVal NewValue As Boolean)
   m_UseDraftBranch = NewValue
End Property

Public Property Get CurrentSvnRev(Optional ByVal RequerySvnRev As Boolean = False) As String
   If Len(m_SvnRev) = 0 Or RequerySvnRev Then
      m_SvnRev = GetRevisionFromWeb
      If UseDraftBranch Then
         m_SvnRev = m_SvnRev & "-draft"
      End If
   End If
   CurrentSvnRev = m_SvnRev
End Property

Public Sub UpdateCodeModules()
   
   Dim SelectSql As String
   Dim IsFirstRecord As Boolean
   
   SelectSql = "select id, url from usys_Appfiles where url > ''"
   
   With CreateObject("ADODB.Recordset")
      .CursorLocation = 3 'adUseClient
      .Open SelectSql, CodeProject.Connection, 1, 1 ' 1 = adOpenKeyset, 1 = adLockReadOnly
      Set .ActiveConnection = Nothing
   
      IsFirstRecord = True
      Do While Not .EOF
         UpdateCodeModuleInTable .Fields(0).Value, .Fields(1).Value, IsFirstRecord
         If IsFirstRecord Then IsFirstRecord = False
         .MoveNext
      Loop
      
      .Close

   End With

End Sub

Private Sub UpdateCodeModuleInTable(ByVal ModuleName As String, ByVal ACLibPath As String, Optional ByVal RequerySvnRev As Boolean = False)

   Dim TempFile As String
   Dim DownLoadUrl As String
   
   TempFile = FileTools.TempPath & ModuleName & ".cls"
   
   DownLoadUrl = SvnBaseUrl
   If UseDraftBranch Then
      DownLoadUrl = DownLoadUrl & "branches/draft"
   Else
      DownLoadUrl = DownLoadUrl & "trunk/codelib"
   End If
   DownLoadUrl = DownLoadUrl & ACLibPath
   
   DownloadFileFromWeb DownLoadUrl, TempFile
   CurrentApplication.SaveAppFile ModuleName, TempFile, True, "SvnRev", Me.CurrentSvnRev(RequerySvnRev)
   Kill TempFile

End Sub

Private Function GetRevisionFromWeb() As Long

   Const RevisionTag As String = "Revision "

   Dim IE As Object           ' SHDocVw.InternetExplorer
   Dim HtmlDoc As Object      ' MSHTML.HTMLDocument
   Dim HtmlElements As Object ' MSHTML.IHTMLElementCollection
   Dim RevText As String
      
   OpenIEandLoadHtmlDoc SvnBaseUrl, IE, HtmlDoc
   
   Set HtmlElements = HtmlDoc.getElementsByTagName("H2")
   
   If HtmlElements.Length > 0 Then
      RevText = HtmlElements.Item.innerText
   Else
      Err.Raise vbObjectError, "GetRevisionFromWeb", "Text mit Revisionsnummer fehlt in '" & SvnBaseUrl & "'"
   End If

   Set HtmlElements = Nothing
   HtmlDoc.Close
   Set HtmlDoc = Nothing

   IE.Quit
   Set IE = Nothing
   
   Dim RevPos As Long
   
   RevPos = InStr(1, RevText, RevisionTag)
   If RevPos = 0 Then
      Err.Raise vbObjectError, "GetRevisionFromWeb", "Text mit Revisionsnummer fehlt in '" & SvnBaseUrl & "'"
   End If
   RevText = Mid(RevText, RevPos + Len(RevisionTag))
   
   GetRevisionFromWeb = Val(RevText)

End Function

Private Sub OpenIEandLoadHtmlDoc(ByVal Url As String, ByRef IE As Object, ByRef HtmlDoc As Object)

   Dim TimeOut As Long
   Dim RunInTimeOut As Boolean
   Dim ErrHdlCnt As Long
   
   Dim ErrNumber As Long
   Dim ErrDescription As String

On Error Resume Next
   Set IE = CreateObject("InternetExplorer.Application")
   Do While Err.Number = -2147023706 And ErrHdlCnt < 10
      Err.Clear
      ErrHdlCnt = ErrHdlCnt + 1
      Set IE = CreateObject("InternetExplorer.Application")
   Loop

   If Err.Number <> 0 Then
      ErrNumber = Err.Number
      ErrDescription = Err.Description
   On Error GoTo 0
      Err.Raise ErrNumber, "ACLibWebImporter.OpenIEandLoadHtmlDoc", ErrDescription
   End If
   
On Error GoTo 0
   
   With IE
      TimeOut = Timer + 10
      Do While .Busy And (Not RunInTimeOut)
         DoEvents
         If Timer > TimeOut Then RunInTimeOut = True
      Loop
      
      If Not RunInTimeOut Then
         .Visible = 0
         .navigate Url
         TimeOut = Timer + 10
         Do Until .ReadyState = 4 Or RunInTimeOut
            DoEvents
            If Timer > TimeOut Then RunInTimeOut = True
         Loop
      End If
      
      If RunInTimeOut Then
         On Error Resume Next
         IE.Quit
         Set IE = Nothing
         On Error GoTo 0
         Err.Raise vbObjectError, "OpenIEandLoadHtmlDoc", "Time-Out beim Laden von '" & Url & "'"
      End If
      
      Set HtmlDoc = IE.Document
             
   End With

End Sub

Private Sub DownloadFileFromWeb(ByVal Url As String, ByVal TargetPath As String)
   If FileExists(TargetPath) Then Kill TargetPath
   DeleteUrlCacheEntry Url
   URLDownloadToFile 0, Url, TargetPath, 0, 0
End Sub
