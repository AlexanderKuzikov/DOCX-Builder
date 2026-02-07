Set WshShell = CreateObject("WScript.Shell") 

' Проверяем порт 5555 через PowerShell
' Если соединение есть (порт занят), вернет exit 1
' Если нет (порт свободен), вернет exit 0
command = "powershell -Command ""If (Get-NetTCPConnection -LocalPort 5555 -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"""
returnCode = WshShell.Run(command, 0, True)

If returnCode = 1 Then
    ' Порт занят -> Сервер уже работает, просто открываем вкладку
    WshShell.Run "http://localhost:5555"
Else
    ' Порт свободен -> Запускаем сервер скрыто (в окне cmd)
    ' node server.js будет висеть в процессах, но без окна
    WshShell.Run "cmd /c node server.js", 0
End If
