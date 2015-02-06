'╔════════════════════════════════════════════════════════════╗
'║ Updater for Virtual Paradise (virtualparadise.org)         ║
'║ Written in 2015 by Christopher Daxon                       ║
'║ Open-source. Modify as needed but keep original credit.    ║
'╚════════════════════════════════════════════════════════════╝

'Future Ideas:      Use command string for args.
'                   Download and use 7z version instead of zip.
'                   Better GUI with everything in one window.
'                   Support VPUpdater being included with the downloaded update

Imports System.Net
Imports System.Io

Public Class frmMain
    '--- Configure main options here ---
    Private Const VersionURL As String = "http://virtualparadise.org/version.txt" 'URL to get the new version string. Length of this string will also be used to find the offset.
    Private Const DownloadURL As String = "http://virtualparadise.org/downloads/virtualparadise_{newVersion}_windows_x86.zip" 'URL to download new version zip. {newVersion} will be replaced with the same string.
    Private Const ExeFindString As String = "Virtual Paradise " 'String to find in EXE where the version string is expected following it

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
        If File.Exists(Application.StartupPath & "\vp.version") = True Then
            'Exising file from previous checks
            curVersion = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "\vp.version")
        Else
            'New VP Updater installation, extract the version string from the VP exe
            Dim vpExe As String = File.ReadAllText(Application.StartupPath & "\VirtualParadise.exe")
            Dim index As Integer = vpExe.IndexOf(ExeFindString)
            If index > 0 Then
                curVersion = vpExe.Substring(index + ExeFindString.Length, newVersion.Length)
            Else
                MsgBox("Failed to find version string in VirtualParadise.exe, running existing version. Check for a new release of VPUpdater.", vbCritical + vbOK, "Error")
                Me.Dispose()
                End
            End If
        End If

        'Alert the user of a new version, allow them the option of bypassing update
        If curVersion <> newVersion Then
            If MsgBox("New version available. Update to latest version " & newVersion & "?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "VP Updater") <> vbYes Then GoTo RunInstalledVersion
        Else
            GoTo RunInstalledVersion
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

        'Write new version to file
        WriteTextFile(Application.StartupPath & "\vp.version", newVersion)

RunInstalledVersion:
        Process.Start(Application.StartupPath & "\VirtualParadise.exe") 'Run VP
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

    Private Sub WriteTextFile(ByVal path As String, ByVal value As String)
        Dim sr As StreamWriter
        sr = New StreamWriter(path)
        sr.Write(value)
        sr.Close()
    End Sub
End Class
