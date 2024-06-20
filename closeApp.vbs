Option Explicit

Dim appName
appName = InputBox("Введите название приложения для закрытия:")

If appName <> "" Then
    Dim objWMIService, colProcesses, objProcess

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & appName & "'")

    If colProcesses.Count = 0 Then
        MsgBox "Приложение с названием '" & appName & "' не найдено.", vbExclamation, "Ошибка"
    Else
        For Each objProcess in colProcesses
            objProcess.Terminate()
        Next

        MsgBox "Приложение '" & appName & "' успешно закрыто.", vbInformation, "Успех"
    End If
End If
