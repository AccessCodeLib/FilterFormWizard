Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>%AppFolder%/source/_config_Application.bas</file>
'  <replace>base/_config_Application.bas</replace> 'dieses Modul ersetzt base/_config_Application.bas
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
' Nicht vergessen: USELOCALIZATION = 1 als Compiler-Arg. in Projekteigenschaft einstellen
'
'
Option Compare Database
Option Explicit
Option Private Module

'Versionsnummer
Private Const APPLICATION_VERSION As String = "1.7.1" '2023-03-01

#Const USE_CLASS_APPLICATIONHANDLER_APPFILE = 1
#Const USE_CLASS_APPLICATIONHANDLER_VERSION = 1

Private Const APPLICATION_NAME As String = "ACLib FilterForm Wizard"
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
'/**
' <summary>
' Konfigurationseinstellungen initialisieren
' </summary>
' <param name="oCurrentAppHandler">M�glichkeit einer Referenz�bergabe, damit nicht CurrentApplication genutzt werden muss</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitConfig(Optional ByRef CurrentAppHandlerRef As ApplicationHandler = Nothing)

On Error GoTo HandleErr

'----------------------------------------------------------------------------
' Fehlerbehandlung
'

   modErrorHandler.DefaultErrorHandlerMode = DefaultErrorHandlerMode

   
'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   'defGlobal_FilterFormWizard.ACLibIconFileName = m_ApplicationIconFile


'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If CurrentAppHandlerRef Is Nothing Then
      Set CurrentAppHandlerRef = CurrentApplication
   End If

   With CurrentAppHandlerRef
   
      'Zur Sicherheit AccDb einstellen
      Set .AppDb = CodeDb 'muss auf CodeDb zeigen,
                          'da diese Anwendung als Add-In verwendet wird
   
      'Anwendungsname
      .ApplicationName = APPLICATION_NAME
      .ApplicationFullName = APPLICATION_FULLNAME
      .ApplicationTitle = APPLICATION_TITLE
      
      'Version
      .Version = APPLICATION_VERSION
      
      ' Formular, das am Ende von CurrentApplication.Start aufgerufen wird
      .ApplicationStartFormName = ApplicationStartFormName
   
    
   End With
   
'----------------------------------------------------------------------------
' Erweiterungen:
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

'----------------------------------------------------------------------------
' Konfiguration nach Erweiterungen
'

   'AppIcon
   'oCurrentAppHandler.SetAppIcon CodeProject.Path & "\" & m_ApplicationIconFile, True

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
' Funktionen f�r die Anwendungswartung
' (werden nur im Anwendungsentwurf ben�tigt)
'
'----------------------------------------------------------------------------
' Hilfsfunktion zum Speichern von Dateien in die lokale AppFile-Tabelle
'----------------------------------------------------------------------------
Private Sub SetAppFiles()
   Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)
End Sub
