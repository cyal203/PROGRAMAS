' Mensagem de aviso - Servidor 24/7
' Criado para exibir alerta no login

Dim objShell, strMessage

' Mensagem completa
strMessage = "SUPORTE FENOX INFORMA " & vbCrLf & vbCrLf
strMessage = strMessage & "SERVIDOR LIGADO 24h" & vbCrLf & vbCrLf
strMessage = strMessage & "Prezados," & vbCrLf & vbCrLf
strMessage = strMessage & "Para garantir o funcionamento eficiente do sistema, " & vbCrLf
strMessage = strMessage & "pedimos que o computador servidor permaneca" & vbCrLf
strMessage = strMessage & " ligado 24 horas por dia, 7 dias por semana." & vbCrLf & vbCrLf
strMessage = strMessage & "O desligamento do servidor pode causar" & vbCrLf
strMessage = strMessage & "interrupcoes no sistema e impactar o atendimento." & vbCrLf & vbCrLf
strMessage = strMessage & "Caso seja necessario desligar o computador, certifique-se de " & vbCrLf
strMessage = strMessage & "liga-la novamente." & vbCrLf & vbCrLf
strMessage = strMessage & "____________________________________________________________" & vbCrLf
strMessage = strMessage & "Quaisquer duvida, entrar em contato com o suporte" & vbCrLf
strMessage = strMessage & "     pelo canal de atendimento (12) 99110-0298."

' Exibe a mensagem com botão OK
Set objShell = CreateObject("WScript.Shell")
objShell.Popup strMessage, 0, "AVISO - FENOX", 64 + 4096