Set WshShell = CreateObject("WScript.Shell")
Dim cmd, i, arg
cmd = ""
For i = 0 To WScript.Arguments.Count - 1
    arg = WScript.Arguments(i)
    If InStr(arg, " ") > 0 And Left(arg, 1) <> """" Then
        arg = """" & arg & """"
    End If
    If i > 0 Then cmd = cmd & " "
    cmd = cmd & arg
Next
WshShell.Run cmd, 0, False
