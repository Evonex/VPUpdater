﻿'╔════════════════════════════════════════════════════════════╗
'║ Updater for Virtual Paradise (virtualparadise.org)         ║
'║ Written in 2015 by Christopher Daxon                       ║
'║ Open-source. Modify as needed but keep original credit.    ║
'╚════════════════════════════════════════════════════════════╝

'Future Ideas:      Use command string for args or an INI file.
'You can add        Download and use 7z version instead of zip.
'                   Better GUI with everything in one window.
'                   Support VPUpdater being included with the downloaded update
'                   Self-update

Imports System.Net
Imports System.Io

Public Class frmMain
    '--- Configure main options here ---
    Private Const VersionURL As String = "http://virtualparadise.org/version.txt" 'URL to get the new version string. Length of this string will also be used to find the offset.
    Private Const DownloadURL As String = "http://virtualparadise.org/downloads/virtualparadise_{newVersion}_windows_x86.zip" 'URL to download new version zip. {newVersion} will be replaced with the version number.
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Download the file from the webserver containing latest version
        Dim curVersion As String = Nothing, newVersion As String = Nothing

        Try
            Dim myWebClient As New WebClient()
            newVersion = myWebClient.DownloadString(VersionURL)
            myWebClient.Dispose()
        Catch ex3 As Exception 'Problems downloading? Run installed version
            GoTo RunInstalledVersion
        End Try

        'Get current version
        curVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.StartupPath & "\VirtualParadise.exe").FileVersion.ToString

        'Alert the user of a new version, allow them the option of bypassing update
        If curVersion.Contains(newVersion) Then
            GoTo RunInstalledVersion
        Else
            If MsgBox("New version available. Update to latest version " & newVersion & "?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "VP Updater") <> vbYes Then GoTo RunInstalledVersion
        End If

        Try 'Delete existing update.zip
            My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\update.zip")
        Catch ex As Exception
        End Try

        'Download the file to the VP folder
        Try
            My.Computer.Network.DownloadFile(DownloadURL.Replace("{newVersion}", newVersion), Application.StartupPath & "\update.zip", vbNullString, vbNullString, True, 5000, True)
        Catch ex4 As Exception
            MsgBox("Error downloading update, running existing version.", vbCritical + vbOK, "Error")
            GoTo RunInstalledVersion
        End Try

        'Extract the zip
        UnZip(Application.StartupPath & "\update.zip", Application.StartupPath)

        Try 'Delete existing update.zip
            My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\update.zip")
        Catch ex As Exception
        End Try

RunInstalledVersion:
        If System.Environment.CommandLine.Contains(" ") Then
            'Check it is not any understood commands, else pass it to the VP executable
            Dim CmdLine() As String
            CmdLine = Split(System.Environment.CommandLine, " ", 2)

            If CmdLine.Length > 1 Then
                Process.Start(Application.StartupPath & "\VirtualParadise.exe", CmdLine(1)) 'Run VP with commandline
            Else
                Process.Start(Application.StartupPath & "\VirtualParadise.exe") 'Run VP
            End If
        Else
            Process.Start(Application.StartupPath & "\VirtualParadise.exe") 'Run VP
        End If

        Me.Dispose()
        End

    End Sub

    Private Sub UnZip(ByVal inputZip As String, ByVal outputFolder As String)
        'Requires shell32.dll reference in %windir%\system32\ 
        'Note: Shell32 zip extraction may create files in temp folder that are not deleted

        Dim sc As New Shell32.Shell()
        'Create directory in which you will unzip your files .
        IO.Directory.CreateDirectory(outputFolder)
        'Declare the folder where the files will be extracted
        Dim output As Shell32.Folder = sc.NameSpace(outputFolder)
        'Declare your input zip file as folder  .
        Dim input As Shell32.Folder = sc.NameSpace(inputZip)
        'Extract the files from the zip file using the CopyHere command .
        output.CopyHere(input.Items, 16) 'Copy files. 16 = Yes to All (on replacing files)
    End Sub
End Class
