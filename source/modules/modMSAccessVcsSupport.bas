Attribute VB_Name = "modMSAccessVcsSupport"
Option Compare Database
Option Explicit

Public Sub VcsRunBeforeExport()

   With New ACLibGitHubImporter
      .BranchName = "master"
      .UpdateCodeModules
   End With

End Sub
