Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Public gFSO As FileSystemObject

Sub Main()

  Dim CurrentDoc As ModelDoc2
  Dim CurrentPath As String
  Dim OldWorkDir As String

  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  
  Set CurrentDoc = swApp.ActiveDoc
  OldWorkDir = ""
  If Not CurrentDoc Is Nothing Then
    OldWorkDir = swApp.GetCurrentWorkingDirectory
    CurrentPath = gFSO.GetParentFolderName(CurrentDoc.GetPathName)
    If gFSO.FolderExists(CurrentPath) Then
      swApp.SetCurrentWorkingDirectory CurrentPath
    End If
  End If
  swApp.Command swFileOpen, Nothing
  If OldWorkDir <> "" Then
    swApp.SetCurrentWorkingDirectory OldWorkDir
  End If

End Sub
