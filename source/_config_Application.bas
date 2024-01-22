Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
' Modul: _initApplication
'---------------------------------------------------------------------------------------
'
' Application configuration
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/_config_Application.bas</file>
'  <replace>base/_config_Application.bas</replace>
'  <license>_codelib/license.bas</license>
'  <use>%AppFolder%/source/defGlobal_ACLibFilterFormWizard.bas</use>
'  <use>base/_initApplication.bas</use>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/ApplicationHandler_AppFile.cls</use>
'  <use>_codelib/addins/shared/AppFileCodeModulTransfer.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>localization/L10nTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Don't forget: set USELOCALIZATION = 1 as Compiler argument in Project properties
'
'
Option Compare Database
Option Explicit
Option Private Module

'Version
Private Const APPLICATION_VERSION As String = "1.8.4" '2024-01

#Const USE_CLASS_APPLICATIONHANDLER_APPFILE = 1
#Const USE_CLASS_APPLICATIONHANDLER_VERSION = 1

Public Const APPLICATION_NAME As String = "ACLib FilterForm Wizard"
Private Const APPLICATION_FULLNAME As String = "Access Code Library - FilterForm Wizard"
Private Const APPLICATION_TITLE As String = APPLICATION_FULLNAME
Private Const APPLICATION_ICONFILE As String = "ACLib.ico"

Public Const APPLICATION_DOWNLOADSOURCE As String = "https://wiki.access-codelib.net/ACLib-FilterForm-Wizard"
Private Const APPLICATION_DOWNLOAD_FOLDER As String = "https://access-codelib.net/download/addins/"
Private Const APPLICATION_DOWNLOAD_VERSIONXMLFILE As String = APPLICATION_DOWNLOAD_FOLDER & "ACLibFilterFormWizard.xml"

Private Const ApplicationStartFormName As String = "frmFilterFormWizard"

Public Const APPLICATION_FILTERCODEMODULE_USEVBCOMPONENTSIMPORT As Boolean = True

Private m_Extensions As ApplicationHandler_ExtensionCollection
'

'---------------------------------------------------------------------------------------
' Sub: InitConfig
'---------------------------------------------------------------------------------------
'
' Init application configuration
'
' Parameters:
'     CurrentAppHandlerRef - Possibility of a reference transfer so that CurrentApplication does not have to be used</param>
'
'---------------------------------------------------------------------------------------
Public Sub InitConfig(Optional ByRef CurrentAppHandlerRef As ApplicationHandler = Nothing)

On Error GoTo HandleErr

'----------------------------------------------------------------------------
' Error handler
'

   modErrorHandler.DefaultErrorHandlerMode = DefaultErrorHandlerMode

'----------------------------------------------------------------------------
' Application instance
'
   If CurrentAppHandlerRef Is Nothing Then
      Set CurrentAppHandlerRef = CurrentApplication
   End If

   With CurrentAppHandlerRef
   
      'To be on the safe side, set AccDb
      Set .AppDb = CodeDb 'must point to CodeDb,
                          'as this application is used as an add-in
   
      'Application name
      .ApplicationName = APPLICATION_NAME
      .ApplicationFullName = APPLICATION_FULLNAME
      .ApplicationTitle = APPLICATION_TITLE
      
      'Version
      .Version = APPLICATION_VERSION
      
      'Form called at the end of CurrentApplication.Start
      .ApplicationStartFormName = ApplicationStartFormName
    
   End With
   
'----------------------------------------------------------------------------
' Extensions:
'
   Set m_Extensions = New ApplicationHandler_ExtensionCollection
   With m_Extensions
      Set .ApplicationHandler = CurrentAppHandlerRef
      .Add New ApplicationHandler_AppFile
   
#If USE_CLASS_APPLICATIONHANDLER_VERSION = 1 Then
      Dim AppHdlVersion As ApplicationHandler_Version
      Set AppHdlVersion = New ApplicationHandler_Version
      .Add AppHdlVersion
      AppHdlVersion.XmlVersionCheckFile = APPLICATION_DOWNLOAD_VERSIONXMLFILE
#End If

   End With

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "InitConfig", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub


'############################################################################
'
' Functions for application maintenance
' (only needed in the application design)
'
'----------------------------------------------------------------------------
' Auxiliary function for saving files to the local AppFile table
'----------------------------------------------------------------------------
Private Sub SetAppFiles()
   Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)
End Sub
