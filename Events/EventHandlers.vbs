Option Explicit

Private Const ADMIN = "Administrator"
Private Const PASSWORD = "supersecretpassword"

Function Include(sInstFile)
   Dim f, s, oFSO
   Set oFSO = CreateObject("Scripting.FileSystemObject")
   On Error Resume Next
   If oFSO.FileExists(sInstFile) Then
      Set f = oFSO.OpenTextFile(sInstFile)
      s = f.ReadAll
      f.Close
      ExecuteGlobal s
   End If
   On Error Goto 0
End Function

'   Sub OnClientConnect(oClient)
'   End Sub

'   Sub OnHELO(oClient)
'   End Sub

Sub OnAcceptMessage(oClient, oMessage)

	REM *****  BEGIN LIST SERVER  *****
	Include("C:\Program Files (x86)\hMailServer\Events\listserv.vbs") '<-- CHANGE THIS!!!
	write_log(d1 & " OnAcceptMessage Event " & d1)
	Set obApp = CreateObject("hMailServer.Application")
	Call obApp.Authenticate(ADMIN, PASSWORD)
	add_client_info oClient, oMessage 	'prepare_mailconfiguration oClient, oMessage
	REM *****  END LIST SERVER  *****

End Sub

Sub OnDeliveryStart(oMessage)

	REM *****  BEGIN LIST SERVER  *****
	Include("C:\Program Files (x86)\hMailServer\Events\listserv.vbs") '<-- CHANGE THIS!!!
	write_log (d1 & " OnDeliveryStart Event " & d1)
	Set obApp = CreateObject("hMailServer.Application")
	Call obApp.Authenticate(ADMIN, PASSWORD)
	write_log (s0 & "Mail FromAddress: " & oMessage.FromAddress)
	write_log (s0 & "Mail From: " & oMessage.From)
	write_log (s0 & "Mail To: " & oMessage.To)
	write_log (s0 & "Subject: " & oMessage.Subject)
	If mail_configuration_active then
		write_log (s0 & a1 & "Starting mail configuration check")
		Result.Value = process_mailconfiguration (oMessage)
	End If
	REM *****  END LIST SERVER  *****

End Sub

Sub OnDeliverMessage(oMessage)

	REM *****  BEGIN LIST SERVER  *****
	Include("C:\Program Files (x86)\hMailServer\Events\listserv.vbs") '<-- CHANGE THIS!!!
	write_log (d1 & " OnDeliverMessage Event " & d1)
	write_log (s1 & "Message Delivered")
	REM *****  END LIST SERVER  *****

End Sub

'   Sub OnBackupFailed(sReason)
'   End Sub

'   Sub OnBackupCompleted()
'   End Sub

'   Sub OnError(iSeverity, iCode, sSource, sDescription)
'   End Sub

'   Sub OnDeliveryFailed(oMessage, sRecipient, sErrorMessage)
'   End Sub

'   Sub OnExternalAccountDownload(oMessage, sRemoteUID)
'   End Sub
