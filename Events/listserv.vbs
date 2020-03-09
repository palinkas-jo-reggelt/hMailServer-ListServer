'******************************************************************************************************************************
'********** List Server Settings                                                                                     **********
'******************************************************************************************************************************

'--------History------------------------------------------
'ListServer Script Version 2.4 - Sunday, March 03, 2013
'See accompanying documentation at http://www.sspxusa.org/goodies/hMailServer/listserver_manual/listserver_manual.htm

'Original Listserver Script written by AndyP
'modifications by Peter Hyde and Brother Gabriel-Marie


'Note that if your list config contains email_subscription=false then it won't check to see whether the poster is valid.

'------------------------------------------------------------------
' Global variables and settings
'Make sure to adjust these to match your configuration
'The rest of the script should not need any adjustment
'------------------------------------------------------------------

Public obApp
Public domain_buffer
Public configpaths
Public procode
Public configbuffer(2,50)
Public configbuffer_nr
Public logspath
public configfile_standard_settings
Public generalposterslist
Public general_allowed_list
Public SMTP_log_position	' internal storage file
Public msg_from 
Public msg_fromaddress

Public Const RootDir = "C:\Program Files (x86)\hMailServer\"	'ends with backslash
Public Const ipslocalhost = "127.0.0.1#192.168.1.2"  'separated by #
Public Const using_v5 = true
Public Const serveradmin = "admin@mydomain.com"
Public Const write_log_active = true
Public Const mail_configuration_active = true
Public Const apply_std_sett_lists_without_cf = false
Public Const apply_std_sett_accounts_without_cf = false
Public Const apply_standard_settings_file = ""		'if empty or not found, standards will be for all. Each address in a new line.
Public Const apply_standard_settings_file_negate = false
Public Const subject_help="help" ' Commands for subscription via email
Public Const subject_list="list"
Public Const subject_subscription="subscribe"
Public Const subject_subscription_address="subscribeaddress"
Public Const subject_subscription_list="subscribelist"
Public Const subject_unsubscription="unsubscribe"
Public Const subject_unsubscription_address="unsubscribeaddress"
Public Const subject_unsubscription_list="unsubscribelist"
Public Const delta_pos_log = 100000 ' 100000 for small installations, 1000000 for large installations
Public Const smtp_log_search_string = "SENT: 220 mail.anotherplace.org ESMTP"
public const configfileextension = ".hms"	'don't leave off the dot!


msg_from = "ListServer <listServ@%domain%>" 	' PeterWeb defined these two for messages sent FROM the listserver; BGM substituted the domain tokens
msg_fromaddress = "listServ@%domain%"  		' they were left empty in original script!

'Spacers and tokens for logging
Public Const s0		= "~"							'logging base level
Public Const s1 	= "     "						'logging level 1
Public Const s2 	= "          "					'logging level 2
public const a1		= "-->"							'notice.  Use this when an integral action is performed
Public Const e1		= "===ERROR===>"				'error
Public const w1		= "===WARNING=>"				'warning
Public Const d1		= "-------------------------"	'divider



'addresses who can post to ALL domains go here.
Public Const globally_allowed_list = "admin@domain.com#admin@lists.domain.net#me@mydomain.tld"	 ' separated by #
general_allowed_list = globally_allowed_list

'file containing list of generally allowed posters for specified domain (should be in the domain's directory)  
'It is okay if the file doesn't exist.
'the domain token will get replaced in the script below
generalposterslist = RootDir & "Config\%domain%\domainposters.txt"	

logspath = RootDir & "Logs\"   'ends with a backslash
configpaths = RootDir & "Config\#" & RootDir & "Config\%domain%\" 	' separated by #, ending with \, %domain% will be replaced by the domain of the list
configfile_standard_settings = RootDir & "Config\config.txt"		'if empty, standard settings in function get_config_string
SMTP_log_position = RootDir & "Events\smtplogposition.txt" 

'******************************************************************************************************************************
'**********  LIST SERVER CODE                                                                                        **********
'******************************************************************************************************************************

'------------------------------------------------------------------
' Check mail configuration / ListServer
'------------------------------------------------------------------


Function add_client_info(oClient, oMessage)
	oMessage.HeaderValue("oclient") = oClient.IPAddress & "#" & oClient.Username  & "#" & oClient.Port
	write_log (s1 & "Adding oclient header: " & oClient.IPAddress & "#" & oClient.Username  & "#" & oClient.Port)
	oMessage.save
End Function

'PROCESS MAIL CONFIGURATION
'return 1 for failure and 0 for success
function process_mailconfiguration(oMessage)
	process_mailconfiguration = 0
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	Dim fs 
	Set fs = CreateObject("scripting.filesystemobject")
	
	Dim mailfrom
	Dim mailto
	Dim rgh(500)
	Dim addr(500)
	Dim lcl(500)
	Dim cmd(500)
	Dim fnd
	Dim arr
	Dim add_admin
	Dim recipientnr
	Dim ipaddr
	dim usern
	Dim smtprec					'smtp recipient
	Dim smtprecarr				'smtp recipient array
	Dim noneconfigured		'flag for whether the list is configured [I think]
	Dim headerto
	Dim tmpMessage
	Dim msgcontent
	Dim temp
	Dim configfile
	
	
	smtprec = ""
	If oMessage.HeaderValue("oclient") <> "" then
		arr = Split(oMessage.HeaderValue("oclient"),"#")
		If UBound(arr) = 2 then
			ipaddr = arr(0)
			usern = arr(1)
			smtprec = LCase(get_smtp_recipient(oMessage, ipaddr))
		Else
			write_log (s1 & "Cannot read oclient header in email.")
		End If
	Else
		write_log (s1 & "No oclient header in email. Internal Mail!")
		' I had to remove these lines of Peter's because they were causing an infinite loop in the message cycle.  I had to kill hMailserver.exe to clear the queue.
		' Here was his proposal (I am not concerned with version 4)
		' write_log ("    No oclient header in email. Internal Mail! Adding workaround defaults") ' workaround text PeterWeb, plus 3 lines below
		' ipaddr = "127.0.0.1"
		' usern = ""
		' smtprec=LCase(get_smtp_recipient(oMessage, ipaddr)) ' in v5 the ipaddr isn't used in this call anyway, but maybe below...
	End If
	
	If smtprec <> "" then
		mailfrom = oMessage.FromAddress
		smtprecarr = Split(smtprec,"#")
		noneconfigured = True
		configbuffer_nr = 0
		
        Dim k
		For k = 1 To UBound(smtprecarr) + 1
			addr(k) = smtprecarr(k - 1)
			rgh(k) = True
			cmd(k) = false
			mailto = addr(k)
			
			write_log (s1 & "Recipient: " & addr(k))
			
			If update_config_buffer(mailto) Then
				print_config_string mailto, "write"
				If get_config_string(mailto, "email_subscription") = "true" Then
					write_log (s1 & "Subscription via email activated.")
					cmd(k) = do_emailsubscription (mailto, mailfrom, oMessage)
				End If
				
				If Not cmd(k) then
					If has_client_authenticated_man(ipaddr,usern) And get_config_string(mailto, "allow_authenticated_users") = "true" Then
						rgh(k) = True
					Else
						rgh(k) = check_recipient(mailto, mailfrom, oMessage)
					End If
				Else
					rgh(k) = False
				End If
				
				noneconfigured = false
			Else
				write_log (s2 & a1 & "No config file found; Either this isn't a mailing list, or it isn't configured.")
			End If
		Next
		
		If noneconfigured = false then
			recipientnr = UBound(smtprecarr) + 1
            Dim g
			For g = 0 To oMessage.headers.count -1  ' added "- 1" PeterWeb
				If oMessage.headers(g).name = "oclient" Then
					oMessage.headers(g).delete
					g = oMessage.headers.count + 1
				End if
			Next
			headerto = oMessage.to
			oMessage.ClearRecipients
			msgcontent = msg_file_content_read(oMessage)
			write_log (s1 & "Changing recipients")
			
			dim thistodo 'fetch the todo setting from the config file
			thistodo =	get_config_string(mailto, "todo")

		    Dim i
			For i = 1 To recipientnr
				add_admin = False
				If rgh(i) = true And cmd(k) = False Then
					write_log (s1 & "Recipient " & addr(i) & " is accepted")
					If is_local_list(addr(i)) then   'the recipient is the mailing list itself	
						write_log (s2 & "Generating new mail for " & addr(i))
						Set tmpMessage = CreateObject("hMailServer.Message")
						msg_file_content_write tmpMessage, msgcontent
						write_log (s2 & "Copying finished.")
						
						tmpMessage.FromAddress = addr(i)
						tmpMessage.AddRecipient addr(i), addr(i)
						
						write_log (s1 & "Adding available header and footer")
						attach_header_footer tmpMessage, addr(i), mailfrom
						write_log (s1 & "Modifiying message")
						modify_msg tmpMessage, addr(i), mailfrom
						
						tmpMessage.HeaderValue("To") = headerto
						tmpMessage.HeaderValue("oclient") = ""
						write_log (s2 & a1 & "Generation finished. Now saving to queue")
						tmpMessage.Save
					else
						oMessage.AddRecipient addr(i), addr(i)
					End If
					
					write_notification True, mailfrom, mailto   'BGM add
					
				ElseIf cmd(i) = True Then
					write_log (s1 & "Subscription via email command " & addr(i) & " detected. Deleting recipient.")
					if get_config_string(mailto, "email_subscription_admin_notification") = "true" then
						add_admin = true
						write_log (s2 & "Sending a copy to admin.")
					end if
				ElseIf thistodo = "delete" Then
					write_notification False, mailfrom, mailto  'BGM add
					write_log (s1 & "Recipient " & addr(i) & " has been deleted")
				ElseIf thistodo = "nothing" Then
					write_log (s1 & "Recipient " & addr(i) & " has todo nothing. Adding the recipient.")
					oMessage.AddRecipient addr(i), addr(i)
				ElseIf thistodo = "redirect" then
					add_admin = true
					write_log (s1 & "Recipient " & addr(i) & " has been redirected to admin")
				Else
					write_log (e1 & "Recipient " & addr(i) & " has not been processed. No valid command found. Script error.")
				End If
				
				check_list_setting(addr(i))
				
				dim thisadmin	'fetch the admin setting from the config file
				thisadmin = get_config_string(mailto, "admin")
				thisadmin = ReplaceTokens(temp, mailto, mailfrom)	
				' thisadmin = ReplaceTokens(thisadmin, mailto, mailfrom)	'http://hmailserver.com/forum/viewtopic.php?p=155148#p155148
			    Dim n
				'If admin is already in the list of recipients, don't add it again
				For n = 0 to omessage.recipients.count - 1
					if omessage.recipients(n).Address = thisadmin then
						add_admin = false
						n = omessage.recipients.count
						write_log(s2 & "Admin is already a member.")
					end if
				next
				'if admin isn't already a recipient, and the setting is true, add him to the list
				If add_admin = True then
					write_log (s1 & "Adding admin email address.")
					oMessage.AddRecipient thisadmin, thisadmin
				End If
			Next
			
			oMessage.HeaderValue("To") = headerto
			If oMessage.Recipients.count = 0 Then
				process_mailconfiguration = 1
			End If
		Else
			write_log (s1 & "No recipient configured")
		End If
		oMessage.Subject = Replace(oMessage.Subject, get_pw(oMessage.Subject, mailto),"")
		oMessage.save
	End If
End function



'MODIFY THE MESSAGE HEADERS AND ATTACHMENTS
Sub modify_msg(oMessage, mailto, mailfrom)
	write_log(s0 & "MODIFYING MESSAGE")
	Dim temp
	Dim temp1
	Dim i
	Dim ts
	Dim del
	write_log(s1 & "Removing possible pws")
	oMessage.Subject = Replace(oMessage.Subject, get_pw(oMessage.Subject, mailto),"")
	oMessage.Subject = Replace(oMessage.Subject, create_pw(mailto),"")

	dim thisadmin	'fetch the admin setting from the config file
	thisadmin = get_config_string(mailto, "admin")
	thisadmin = ReplaceTokens(temp, mailto, mailfrom)		
	
	'ADD SENDER HEADER
	temp = get_config_string(mailto, "addsenderheader")
	temp = ReplaceTokens(temp, mailto, mailfrom)	'BGM Friday, March 01, 2013
	If temp = "list" Then
		oMessage.HeaderValue("Sender") = mailto
		write_log (s2 & "Changing header Sender to " & mailto)
	elseIf temp <> "" Then
		oMessage.HeaderValue("Sender") = temp
		write_log (s2 & "Changing header Sender to " & temp)
	End If
	
	'ADD REPLY-TO HEADER
	temp = get_config_string(mailto, "addreplytoheader")
	temp = ReplaceTokens(temp, mailto, mailfrom)	'BGM Friday, March 01, 2013
	If temp = "list" Then
		oMessage.HeaderValue("Reply-To") = mailto
		write_log (s2 & "Changing header Reply-To to " & mailto)
	elseIf temp <> "" Then
		oMessage.HeaderValue("Reply-To") = temp
		write_log (s2 & "Changing header Reply-To to " & temp)
	elseIf Len(oMessage.HeaderValue("Reply-To")) < 3 Then
		oMessage.HeaderValue("Reply-To") = mailfrom
		write_log (s2 & "Changing empty header Reply-To to " & mailfrom)
		'write_log (s1 & a1 & "Header Reply-To is empty")
	End If
	
	'ADD RETURN-PATH HEADER
	temp = get_config_string(mailto, "addreturnpathheader")
	temp = ReplaceTokens(temp, mailto, mailfrom)	'BGM Friday, March 01, 2013
	If temp = "list" Then
		oMessage.HeaderValue("Return-Path") = mailto
		write_log (s2 & "Changing header Return-Path to " & mailto)
	elseIf temp <> "" Then
		oMessage.HeaderValue("Return-Path") = temp
		write_log (s2 & "Changing header Return-Path to " & temp)
	End If
	
	'ADD SMTP-FROM HEADER
	temp = get_config_string(mailto, "smtpmailfrom")
	temp = ReplaceTokens(temp, mailto, mailfrom)	'BGM Friday, March 01, 2013	
	If temp = "admin" then
		oMessage.FromAddress = thisadmin
		write_log (s2 & "Changing mail from in smtp session to " & thisadmin)
	ElseIf temp = "list" then
		oMessage.FromAddress = mailto
		write_log (s2 & "Changing mail from in smtp session to " & temp = "list")
	ElseIf Len(temp) > 5 then
		oMessage.FromAddress = temp
		write_log (s2 & "Changing mail from in smtp session to " & temp)
	ElseIf Len(temp1) > 5 then
		oMessage.FromAddress = thisadmin
		write_log (s2 & "Changing mail from in smtp session to " & thisadmin)
	End If
	
	'ADD SUBJECT PREFIX
	temp = get_config_string(mailto, "subjectprefix")
	temp = ReplaceTokens(temp, mailto, mailfrom)	
	If temp <> "" Then
		If InStr(1, oMessage.Subject, temp) = 0 Then
			oMessage.Subject = temp & oMessage.Subject
		End If
		write_log (s2 & "Adding prefix to subject " & temp)
	End If
	
	'ADJUST THE ATTACHMENTS
	write_log (s1 & a1 & "Checking attachments")
	Dim thisext
	Dim att_temp
	Dim ats_temp
	Dim atws_temp
	dim atn_temp
	att_temp = get_config_string(mailto, "allowed_attachments")
	att_temp = ReplaceTokens(att_temp, mailto, mailfrom)	
	
	atn_temp = get_config_string(mailto, "disallowed_attachments")
	atn_temp = ReplaceTokens(atn_temp, mailto, mailfrom)		
	
	ats_temp = get_config_string(mailto, "allowed_attachment_size")			'this must be numeric, so don't do token replacement
	atws_temp = get_config_string(mailto, "allowed_attachments_total_size")	'this must be numeric, so don't do token replacement
	
	If att_temp = "none" Then
		write_log (s2 & a1 & "Removing all attachments")
		oMessage.Attachments.clear
	Else
		ts = CLng(0)
		i = oMessage.Attachments.count
		write_log (s1 & "There are " & i & " attachments.")
		Do While i > 0
			i = i - 1
			del = false
			thisext = get_file_suffix(oMessage.Attachments(i).Filename)
			
			'then see if this file is amoungst the disallowed_attachments
			'Once we've removed all the disallowed ones, we can rest easy.
			If InStr(1,"   #" & atn_temp & "#", "#" & thisext & "#") > 0 Then		'if the extension IS found, delete it
				write_log (s2 & a1 & "Removing disallowed extension: " & thisext)
				del = true
			End If			
			
			'Are all attachments allowed?
			If att_temp = "*" Or att_temp = "" Then
				'don't do anything!
			'Is this file amoungst the allowed_attachments? If not, delete it.
			ElseIf InStr(1,"   #" & att_temp & "#", "#" & thisext & "#") = 0 Then
				write_log(s2 & a1 & "Removing unapproved extension: " & thisext)
				del = true
			End If
			
			'Now lets work on the size of the attachments.
			If ats_temp = "" then
			ElseIf Not IsNumeric(ats_temp) Then
					write_log (e1 & "allowed_attachment_size is not numeric, cannot be converted to long")
			Else
				If oMessage.Attachments(i).size > CLng(Trim(ats_temp)) * CLng(1000) then
					write_log (s2 & a1 & "Removing too large file: " & thisext & " Size: " & oMessage.Attachments(i).size)
					del = true
				End If
			End If
			If del Then
				oMessage.Attachments(i).delete
			Else
				ts = ts + oMessage.Attachments(i).size
			End if
		loop
		If atws_temp = "" then
		ElseIf Not IsNumeric(atws_temp) Then
				write_log (e1 & "allowed_attachments_total_size is not numeric, cannot be converted to long")
		Else
			If CLng(ts) > CLng(Trim(atws_temp)) * CLng(1000) then
				write_log (s1 & a1 & "Total size of: " & ts & " exceeded.")
				write_log (s2 & a1 & "Removing all attachments.")
				oMessage.Attachments.clear
			End If
		End If
	End If
End sub

'erg (true or false) decides whether to send a success message or a failure message
'mailfrom is the address who should receive the notifications
'mailto is the address that is to be checked
sub write_notification(erg, mailfrom, mailto)
	write_log(s0 & "Sending Notification Message to " & mailfrom & " (mailfrom=" & mailfrom & " mailto=" & mailto & ")")  'BGM add
	Dim fs
	Set fs = CreateObject("scripting.filesystemobject")
	Dim f
	Dim temp
	Dim success_notification_text
	Dim failure_notification_text
	dim nMessage
	success_notification_text = "Hello %mailfrom%" & nl & nl & "your mail has been successfully deliverd to %mailto%." & nl & nl & _
		"Regards" & nl & serveradmin
	failure_notification_text = "Hello %mailfrom%" & nl & nl & "you are not permitted to deliver an email to %mailto%" & nl & nl & _
		"Please contact the administrator." & nl & nl & "Regards" & nl & serveradmin
	
	dim thissuccesssubject
	dim thisfailuresubject
	dim thisfailurefile
	dim thissuccessfile
	dim thisadmin
	
	thisfailuresubject = ReplaceTokens(get_config_string(mailto, "failure_notification_subject"), mailto, mailfrom)
	thissuccesssubject = ReplaceTokens(get_config_string(mailto, "success_notification_subject"), mailto, mailfrom)
	thisfailurefile = ReplaceTokens(get_config_string(mailto, "failure_notification_file"), mailto, mailfrom)
	thissuccessfile = ReplaceTokens(get_config_string(mailto, "success_notification_file"), mailto, mailfrom)
	thisadmin = ReplaceTokens(get_config_string(mailto, "admin"), mailto, mailfrom)

	If get_config_string(mailto, "failure_notification") = "true" And erg = False Then		'why erg=false?
		failure_notification_text = ""
		temp = get_configpath(mailto) & thisfailurefile
		If fs.FileExists(temp) Then
			Set f = fs.opentextfile(temp)
			failure_notification_text = f.readall
			failure_notification_text = ReplaceTokens(failure_notification_text, mailto, mailfrom)
			f.Close
		End If
		write_log (s1 & a1 & "Sending failure notification to sender: " & mailfrom)
		Set nMessage = CreateObject("hMailServer.Message")
			nMessage.From = thisadmin
			nMessage.FromAddress = thisadmin
			nMessage.AddRecipient mailfrom, mailfrom
			nMessage.Subject = thisfailuresubject
			nMessage.Body = failure_notification_text
			nMessage.Save
		Set nMessage = nothing
	'else
		'write_log("****missed first if statement in write_notification****")
	End If
	
	
	If get_config_string(mailto, "success_notification") = "true" And erg = True Then	'BGM change (originally, it was config_string instead of mailto in the first parameter)
		success_notification_text = ""
		temp = get_configpath(mailto) & thissuccessfile
		If fs.FileExists(temp) Then
			Set f = fs.opentextfile(temp)
			success_notification_text = f.readall
			success_notification_text = ReplaceTokens(success_notification_text, mailto, mailfrom)
			f.Close
		End If
		write_log (s1 & a1 & "Sending success notification to sender: " & mailfrom)
		Set nMessage = CreateObject("hMailServer.Message")
			nMessage.From = thisadmin
			nMessage.FromAddress = thisadmin 
			nMessage.AddRecipient mailfrom, mailfrom
			nMessage.Subject = thissuccesssubject
			nMessage.Body = success_notification_text
			nMessage.Save
		Set nMessage = nothing
	'else
		'write_log("****missed second if statement in write_notification****")
	End If
End sub



'ADD HEADER TO MESSAGE
Sub attach_header_footer(oMessage, mailto, mailfrom)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fs
	Dim f
	Set fs = CreateObject("scripting.filesystemobject")
	Dim path
	Dim header
	Dim footer
	Dim headerhtml
	Dim footerhtml
	dim footerlist
	dim temp
	Dim fn
	Dim pos
	
	path = get_configpath(mailto)
	header = ""
	footer = ""
	headerhtml = ""
	footerhtml = ""
	footerlist = ""
	
	fn = path & get_config_string(mailto, "header_file")
	If fs.FileExists(fn) Then
		Set f = fs.opentextfile(fn,ForReading)
		header = ReplaceTokens(f.ReadAll, mailto, mailfrom)
		f.Close
	End If
	fn = path & get_config_string(mailto, "footer_file")
	If fs.FileExists(fn) Then
		Set f = fs.opentextfile(fn,ForReading)
		footer = ReplaceTokens(f.ReadAll, mailto, mailfrom)
		f.Close
	End If
	fn = path & get_config_string(mailto, "header_file_html")
	If fs.FileExists(fn) Then
		Set f = fs.opentextfile(fn,ForReading)
		headerhtml = ReplaceTokens(f.ReadAll, mailto, mailfrom)
		f.Close
	End If
	fn = path & get_config_string(mailto, "footer_file_html")
	If fs.FileExists(fn) Then
		Set f = fs.opentextfile(fn,ForReading)
		footerhtml = ReplaceTokens(f.ReadAll, mailto, mailfrom)
		f.Close
	End If
	if get_config_string(mailto, "attach_recipients_to_footer") = "true" and is_local_list(mailto) then
		footerlist = get_recp_of_list(mailto, get_config_string(mailto, "attach_recipients_delimiter"))
		temp = get_config_string(mailto, "attach_recipients_html")
		if instr("  " & temp, "%recipients%") > 0 then
			temp = replace(temp, "%recipients%", replace(footerlist, nl ,"<br>"))
		else
			temp = replace(footerlist, nl ,"<br>")
		end if
		footerhtml = footerhtml & temp
		footer = footer & footerlist
	end if
	
	
	If header <> "" Then
		write_log (s2 & "Adding text header")
		oMessage.Body = header & nl & oMessage.Body
	End If
	If footer <> "" Then
		write_log (s2 & "Adding text footer")
		oMessage.Body = oMessage.Body & nl & footer
	End If
	If Len(oMessage.HTMLBody) > 5 then
		If headerhtml <> "" Then
			write_log (s2 & "Adding html header")
			pos = InStr(1, oMessage.HTMLBody, "<div")
			pos = InStr(pos + 5, oMessage.HTMLBody, ">")
			If pos < 1 then
				pos = InStr(1, oMessage.HTMLBody, "<body")
				pos = InStr(pos + 5, oMessage.HTMLBody, ">")
			End if
			If pos < 1 Then
				pos = 1
			End if
			oMessage.HTMLBody = Mid(oMessage.HTMLBody, 1, pos) & nl & headerhtml & Mid(oMessage.HTMLBody, pos + 1)
		End If
		If footerhtml <> "" Then
			write_log (s2 & "Adding html footer")
			pos = InStrRev(oMessage.HTMLBody, "</div")
			If pos < 1 then
				pos = InStr(1, oMessage.HTMLBody, "</body")
		End If
			If pos < 2 Then
				pos = 2
			End if
			oMessage.HTMLBody = Mid(oMessage.HTMLBody, 1, pos - 1) & footerhtml & nl & Mid(oMessage.HTMLBody, pos)
		End If
	End if
End sub


sub check_list_setting(mailto)
	Dim fnd
	Dim is_list
	Dim doms
	Dim obList
	Dim obDomain
	fnd = False
	is_list = false
	dim ok
	ok = true
	Set doms = obApp.Domains
	
    Dim i
	For i = 0 To doms.Count - 1
		If LCase(doms.Item(i).Name) = LCase(Mid(mailto, InStr(1,mailto,"@") + 1)) Then
			fnd = True
		End if
	Next
	If fnd = True Then
		Set obDomain = doms.ItemByName(Mid(mailto, InStr(1,mailto,"@") + 1))
		For i = 0 To obDomain.DistributionLists.Count - 1
			If LCase(mailto) = LCase(obDomain.DistributionLists.Item(i).Address) Then
				is_list = True
			End if
		Next
	End If
	If is_list then
		Set obList = obDomain.DistributionLists.ItemByAddress(mailto)
		
		if oblist.Mode <> 0 then
			ok = false
		end if
		if oblist.RequireSMTPAuth = true then
			ok = false
		end if
		
		if not ok then
			write_log (w1 & "Warning: The list is not set to public or requires SMTP authentication. This might interfer with the config settings!")
		end if
	End if
end sub


function update_config_buffer(mailto)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	update_config_buffer = false
	Dim fs
	Dim f
	Set fs = CreateObject("scripting.filesystemobject")
	
	Dim fn
	Dim path
	Dim domain
	domain = Mid(mailto, InStr(1,mailto,"@") + 1)
	Dim content
	content = ""
	dim usefile
	usefile = "currently none"
	Dim arr
	arr = Split(configpaths,"#")
	Dim na
	dim fok
	fok = false
	dim go
	na = Replace(mailto,"@","_at_")
	na = Replace(na,".","_")
	
    Dim i
	For i = 0 To UBound(arr)
		path = Replace(arr(i),"%domain%", domain)
		fn = path & na & configfileextension
		If fs.FileExists(fn) Then
			write_log (s0 & a1 & "Reading config file: " & fn)
			Set f = fs.opentextfile(fn,ForReading)
			content = f.ReadAll
			f.Close
			fok = true
			i = UBound(arr)
		End If
	Next
	
	if not fok and address_in_appdeffile(mailto) then
		go = false
		if apply_std_sett_lists_without_cf and is_local_list(mailto) then
			go = true
			write_log (s1 & a1 & "Applying standard settings for lists activated! This is a list.")
		elseif apply_std_sett_accounts_without_cf and is_local_account(mailto) then
			go = true
			write_log (s1 & a1 & "Applying standard settings for accounts activated! This is an account.")
		end if
		if go then
			If fs.FileExists(configfile_standard_settings) Then
				write_log (s1 & "Reading standard config file: " & configfile_standard_settings)
				Set f = fs.opentextfile(configfile_standard_settings,ForReading)
				content = f.ReadAll
				f.Close
			else
				write_log (s1 & a1 & "Standard config file does not exist: " & configfile_standard_settings)
			End If
		end if
	end if
	
	do while usefile <> ""
		usefile = get_use_in_config(content, path)
		if usefile <> "" then
			write_log (s1 & "Use in config file found: " & usefile)
			If fs.FileExists(usefile) Then
				write_log (s1 & "Reading config file: " & usefile)
				Set f = fs.opentextfile(usefile,ForReading)
				content = f.ReadAll
				f.Close
			else
				write_log (w1 & "Config file does not exist: " & usefile)
				usefile = ""
			End If
		end if
	loop
	
	If content <> "" Then
		update_config_buffer = true
		configbuffer(0,configbuffer_nr) = mailto
		configbuffer(1,configbuffer_nr) = content
		configbuffer(2,configbuffer_nr) = path
		configbuffer_nr = configbuffer_nr + 1
	End If
End function

function get_use_in_config(configcontent, path)
	get_use_in_config = ""
	dim lns
	dim ln
	dim erg
	lns = Split(configcontent, nl)
	
    Dim i
	For i = 0 to ubound(lns)
		ln = replace(lns(i), " ", "")
		if lcase(mid(ln,1,3)) = "use" then
			erg = mid(ln, 5)
			i = ubound(lns) + 10
		end if
	next
	
	if erg <> "" then
		if instr("  " & erg, ":") > 0 and instr("  " & erg, "\") > 0 then
				get_use_in_config = erg
		else
			if mid(path, len(path),1) = "\" then
				get_use_in_config = path & erg
			else
				get_use_in_config = path & "\" & erg
			end if
		end if
	end if
end function

function address_in_appdeffile(mailto)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	address_in_appdeffile = false
	Dim fs
	Dim f
	dim content
	Set fs = CreateObject("scripting.filesystemobject")
	
	if apply_standard_settings_file <> "" then
		If fs.FileExists(apply_standard_settings_file) Then
			Set f = fs.opentextfile(apply_standard_settings_file,ForReading)
			content = f.ReadAll
			content = replace(content, " ", "")
			content = "##" & replace(content, nl, "##") & "##"
			f.Close
			if instr("  " & content, "##" & mailto & "##") > 0 then
				address_in_appdeffile = true
				write_log (s1 & "Address " & mailto & " exists in apply_standard_settings_file -> applying standard setting")	
			else
				write_log (s1 & "Address " & mailto & " does not exist in apply_standard_settings_file -> not applying standard setting")
			end if
		else
			write_log (s1 & "Standard address file does not exist " & apply_standard_settings_file)
			address_in_appdeffile = true
		End If
	else
		'write_log (s1 & "Standard address file is not configured")
		address_in_appdeffile = true
	end if
	if apply_standard_settings_file_negate then
		address_in_appdeffile = not address_in_appdeffile
	end if
end function

'FROM THE LIST'S ADDRESS, FIGURE WHICH CONFIG FILE TO USE
Function get_configpath(mailto)
	get_configpath = ""
    Dim i
	For i = 0 To configbuffer_nr - 1
		If configbuffer(0,i) = mailto Then
			get_configpath = configbuffer(2,i)
			i = configbuffer_nr
		End If
	Next
End Function

'GET THE SETTINGS FROM THE CONFIG FILE
function get_config_string(mailto, setting_name)
	get_config_string = ""
	
	Dim erg
	erg = ""
	Dim source
	source = ""
	Dim ln
	Dim sourcecut
	Dim isarry
	Dim arr
	Dim cmd
	Dim val
	Dim lnnr
	lnnr = 1
	
    Dim i
	For i = 0 To configbuffer_nr - 1
		If configbuffer(0,i) = mailto Then
			source = configbuffer(1,i)
		End If
	Next
	
	'Defaults for the settings in case there is no config file
	Select Case LCase(setting_name)
	   	Case "admin" erg = serveradmin
	   	Case "todo" erg = "delete"
	   	Case "operator" erg = "or"
	   	Case "todo" erg = "delete"
	   	Case "pw_enclosure" erg = "#"
	   	Case "pw" erg = ""
	   	Case "allowmembers" erg = "true"
	   	Case "addsenderheader" erg = "list"
	   	Case "addreplytoheader" erg = ""
	   	Case "addreturnpathheader" erg = ""
	   	Case "smtpmailfrom" erg = "admin"
	   	Case "allowed_attachments" erg = "*"
	   	Case "disallowed_attachments" erg = ""
	   	Case "subjectprefix" erg = ""
	   	Case "success_notification" erg = "false"
	   	Case "success_notification_subject" erg = "Delivery successful"
	   	Case "failure_notification" erg = "false"
	   	Case "failure_notification_subject" erg = "Delivery failure"
	   	Case "attach_recipients_to_footer" erg = "false"
	   	Case "attach_recipients_delimiter" erg = "; "
	   	Case "attach_recipients_html" erg = "<p>%recipients%</p>"
	   	Case "email_subscription" erg = "false"
	   	Case "email_subscription_allowed_cmds" erg = "help#subscribe#unsubscribe"
	   	Case "email_subscription_admin_notification" erg = "true"
	   	Case "email_subscription_success_notification" erg = "true"
	   	Case "email_subscription_failure_notification" erg = "true"
	End Select
	
	sourcecut = source
	
	isarry = false
	Do While Len(sourcecut) > 0
		If InStr(1,sourcecut,Chr(13) & Chr(10)) > 0 then
			ln = Mid(sourcecut,1,InStr(1,sourcecut,Chr(13) & Chr(10))-1)
			sourcecut = Mid(sourcecut,InStr(1,sourcecut,Chr(13) & Chr(10))+2)
		Else
			ln = sourcecut
			sourcecut = ""
		End if
		If Mid(ln, 1, 1) <> "#" And Len(ln) > 0 Then
			If isarry = True Then
				If erg = "" Then
					erg = ln
				Else
					erg = erg & "#" & ln
				End If
			Else
				arr = Split(ln,"=")
				If UBound(arr) = 1 Then
					cmd = Trim(arr(0)) 
					val = arr(1)
				ElseIf UBound(arr) = 0 then
					cmd = Trim(arr(0)) 
					val = ""
				Else
					if instr("   " & ln, "=") > 0 then
						cmd = mid(ln, 1, instr(ln, "=") - 1)
						val = mid(ln, instr(ln, "=") + 1)
					else
						cmd = ""
						val = ""
					end if
				End If
			
				If LCase(cmd) = LCase(setting_name) Then
					If LCase(cmd) = LCase(setting_name) And val = "" Then
						erg = ""
						isarry = True
					Else
						erg = val
					End If
				End If
			End If
		Else
			isarry = False
		End If
		
		lnnr = lnnr + 1
		If lnnr > 2000 Then
			write_log(w1 & "Aborting lookup process, because number of lines reached critical limit.")
			write_log(s1 & "Rest:" & sourcecut)
			write_log(s1 & "Pos:" & InStr(1,sourcecut,Chr(13) & Chr(10)))
			sourcecut = ""
		End If
	Loop
	If setting_name = "pw" Then
		erg = Trim(erg)
	End If
	If erg <> "" Then
		get_config_string = erg
	End if
End function

function print_config_string(mailto, todo)
	print_config_string = ""
	
	Dim erg
	Dim arrval
	arrval = ""
	erg = ""
	Dim source
	source = ""
	Dim output
	output = ""
	Dim ln
	Dim sourcecut
	Dim isarry
	Dim lnnr
	lnnr = 1
	
    Dim i
	For i = 0 To configbuffer_nr - 1
		If configbuffer(0,i) = mailto Then
			source = configbuffer(1,i)
		End If
	Next
	
	sourcecut = source
	
	isarry = false
	Do While Len(sourcecut) > 0
		If InStr(1,sourcecut,Chr(13) & Chr(10)) > 0 then
			ln = Mid(sourcecut,1,InStr(1,sourcecut,Chr(13) & Chr(10))-1)
			sourcecut = Mid(sourcecut,InStr(1,sourcecut,Chr(13) & Chr(10))+2)
		Else
			ln = sourcecut
			sourcecut = ""
		End if
		If Mid(ln, 1, 1) <> "#" And Len(ln) > 0 Then
			If isarry = True Then
				erg = erg & "#" & ln
			Else
				erg = ln
			End If
			If Mid(ln,Len(ln)) = "=" Then
				isarry = True
				erg = ""
				arrval = ln
			End If
		Else
			isarry = False
			If erg <> "" And arrval <> "" Then
				erg = Mid(erg,2)
				output = output & s2 & arrval & erg & nl
				If todo = "write" Then
					write_log (s2 & arrval & erg)
				End If
				erg = ""
				arrval = ""
			End if
		End If
		
		If isarry = False And erg <> "" and Mid(ln, 1, 1) <> "#" And Len(ln) > 0 then
			output = output & s2 & erg & nl
			If todo = "write" Then
				write_log (s2 & erg)
			End If
			erg = ""
		End If
		lnnr = lnnr + 1
		If lnnr > 1000 Then
			sourcecut = ""
		End If
	Loop
	
	If output <> "" And todo = "return" Then
		print_config_string = output
	End if
End function

'GET THE LIST RECIPIENTS OF THE LIST FROM hMailServer
function get_recp_of_list(list, delimiter)
	get_recp_of_list = ""
	Dim fnd
	Dim suffix
	Dim doms
	dim olist
	dim orcps
	dim erg
	dim obDomain
	erg = ""
	suffix = Mid(list, InStr(1,list,"@") + 1)
	fnd = False
	
	Set doms = obApp.Domains
    Dim i
	For i = 0 To doms.Count - 1
		
		If LCase(doms.Item(i).Name) = LCase(suffix) Then
			fnd = True
		End if
	Next
	If fnd Then
		fnd = false
		Set obDomain = doms.ItemByName(suffix)
		For i = 0 To obDomain.DistributionLists.Count - 1
			If LCase(list) = LCase(obDomain.DistributionLists.Item(i).Address) Then
				set olist = obDomain.DistributionLists.Item(i)
				fnd = true
			End if
		Next
	End If
	if fnd then
		Set orcps = olist.Recipients
	    Dim j
		For j = 0 to orcps.count - 1
			If erg = "" Then
				erg = orcps.item(j).RecipientAddress
			elseif delimiter = "nl" then
				erg = erg & nl & orcps.item(j).RecipientAddress
			else
				erg = erg & delimiter & orcps.item(j).RecipientAddress
			End if
		Next
		
		erg = nl & nl & "Recipients of " & list & ":" & nl & erg
	end if
	if erg <> "" then
		get_recp_of_list = erg
	end if
end function

'CONSULT hMailServer AS TO WHETHER A RECIPIENT EXISTS IN THE LIST
Function check_recipient(mailto, mailfrom, omsg)
	Dim pw_right
	Dim allowed_right
	Dim allowed_cnt
	Dim is_member
	Dim is_list
	
	Dim erg
	Dim	fnd
	Dim	arr
	
	Dim obDomain
	Dim obList
	Dim obreclist
	
	write_log (s1 & "Checking recipient: " & mailto)
	pw_right = True
	If get_config_string(mailto, "pw") <> "" Then
		if InStr(1,"   " & omsg.subject, create_pw(mailto)) = 0 Then
			pw_right = false
			write_log (s1 & a1 & "Ooooo! Looky! : PW wrong")
		Else
			write_log (s1 & a1 & "PW OK")
		End If
	End If
	
	is_member = False
	is_list = is_local_list(mailto)
	If is_list Then
		write_log (s2 & "This is a distribution list")
		If is_member_of_local_list(mailfrom,mailto) Then
			is_member = True
		End if
	End If
	
	If Not is_member And is_list then
		write_log (s2 & "Not a member of the distribution list")
	Else
		write_log (s2 & "Member of the distribution list")
	End	If
	
	fnd = false
	dim theseallowedaddresses
	theseallowedaddresses = ReplaceTokens(get_config_string(mailto, "allowaddresses"), mailto, mailfrom)
	arr = Split(theseallowedaddresses,"#")
    Dim i
	For i = 0 To UBound(arr)
		If arr(i) = mailfrom Then
			fnd = True
		End If
	Next
	allowed_right = False
	If UBound(arr) >= 0 then
		write_log (s2 & UBound(arr) + 1 & " members in allowed list found")
		If fnd = true Then
			write_log (s2 & "Member of allowed list")
			allowed_right = true
		else
			write_log (s2 & "Not a member of allowed list (maybe members aren't allowed to post)")
		End If
	End if
	
	erg = false
	If get_config_string(mailto, "pw") = "" Then
		write_log (s2 & "Checking conditions: No PW defined")
		If allowed_right Then
			erg = True
			write_log (s2 & a1 & "result = true : 11 member of allowed list")
		elseIf is_list And is_member And get_config_string(mailto, "allowmembers") = "true" then
			erg = True
			write_log (s2 & a1 & "result = true : 12 member of distribution list")
		elseIf is_list And get_config_string(mailto, "allowmembers") <> "true" And Len(get_config_string(mailto, "allowaddresses")) < 5 then
			erg = True
			write_log (s2 & a1 & "result = true : 13 nothing defined to stop the mail")
		Else
			write_log (s2 & a1 & "result = false : recipient will be deleted")
		End If
	elseIf get_config_string(mailto, "operator") = "or" then
		write_log (s2 & "Checking conditions: Operator or")
		If pw_right Then
			erg = True
			write_log (s2 & a1 & "result = true : 21 PW OK")
		elseIf allowed_right Then
			erg = True
			write_log (s2 & a1 & "result = true : 22 member of allowed list")
		elseIf is_list And is_member And get_config_string(mailto, "allowmembers") = "true" then
			erg = True
			write_log (s2 & a1 & "result = true : 23 member of distribution list")
		Else
			write_log (s2 & a1 & "result = false : recipient will be deleted")
		End If
	Else
		write_log (s2 & "Checking conditions: Operator and")
		If allowed_right And pw_right Then
			erg = True
			write_log (s2 & a1 & "result = true : 31 PW OK and member of allowed list")
		elseIf is_list and is_member And get_config_string(mailto, "allowmembers") = "true" And pw_right then
			erg = True
			write_log (s2 & a1 & "result = true : 32 PW OK and member of distribution list")
		Else
			write_log (s2 & a1 & "result = false : recipient will be deleted")
		End If
	End If
	
	general_allowed_list = GetGeneralAdmins(mailto, mailfrom)
	write_log (s1 & "Global + Domain Posters are:" & general_allowed_list)	
	
	If Not erg Then
		If InStr(1,"  " & general_allowed_list, mailfrom) > 1 Then
			erg = True
			write_log (s2 & "But: Recipient is on the global allowed list.")
		End If
	End If
	
	If erg = True then
		write_log (s2 & "Recipient is OK.")
		check_recipient = true
	Else
		write_log (s2 & a1 & "Recipient will be deleted.")
		check_recipient = false
	End if
End Function

'SUBSCRIPTION VIA EMAIL
Function do_emailsubscription (mailto, mailfrom, omsg)
	do_emailsubscription = False
	Dim fnd
	Dim obDomain
	Dim obList
	Dim sRecipients
	Dim SRecipient
	Dim sTempRecipient
   	Dim	sRecipientID
	Dim allrcpts
	Dim txt
	Dim tmp
	Dim arr
	Dim tmpMessage
	Dim cmd
	cmd = ""
	Dim cmdald
	cmdald = get_config_string(mailto, "email_subscription_allowed_cmds")
	Dim allcmds
	allcmds = "#"
	allcmds = allcmds & subject_help & "#"
	allcmds = allcmds & subject_list & "#"
	allcmds = allcmds & subject_subscription & "#"
	allcmds = allcmds & subject_subscription_address & "#"
	allcmds = allcmds & subject_subscription_list & "#"
	allcmds = allcmds & subject_unsubscription & "#"
	allcmds = allcmds & subject_unsubscription_address & "#"
	allcmds = allcmds & subject_unsubscription_list & "#"
	Dim thisadmin
	thisadmin = ReplaceTokens(get_config_string(mailto, "admin"), mailto, mailfrom)
	
	
	'WHAT?????
	msg_from = ReplaceTokens(msg_from, mailto, mailfrom)	'BGM add
	msg_fromaddress = ReplaceTokens(msg_fromaddress, mailto, mailfrom)	'BGM add
	
	write_log(s1 & "CONSIDERING SUBSCRIPTIONS COMMANDS")
	
	If is_local_list(mailto) Then
		Set obDomain = obApp.Domains.ItemByName(Mid(mailto, InStr(1,mailto,"@") + 1))
		Set obList = obDomain.DistributionLists.ItemByAddress(mailto)
		Set sRecipients = obList.Recipients
		
		allrcpts = ""
        Dim i
		For i = 0 To sRecipients.count - 1
			allrcpts = allrcpts & "#" & sRecipients(i).RecipientAddress
		Next
		
		dim thisfailurenotification
		dim thissuccessnotification
		thisfailurenotification = get_config_string(mailto, "email_subscription_failure_notification")
		thissuccessnotification = get_config_string(mailto, "email_subscription_success_notification")
		
		
		cmd = Trim(omsg.subject)
		If InStr(1, "  #" & cmdald & "#", "#" & cmd & "#") < 1 And InStr(1, "  " & allcmds, "#" & cmd & "#") > 0 Then
			write_log (s2 & "The command " & cmd & " is not in allowed")
			If thisfailurenotification = "true" Then
				Set tmpMessage = CreateObject("hMailServer.Message")
				tmpMessage.From = msg_from
				tmpMessage.FromAddress = msg_fromaddress
				tmpMessage.AddRecipient mailfrom, mailfrom
				tmpMessage.Subject = "Command is not allowed"
				txt = "Hello " & mailfrom & nl & nl
				txt = txt & "the command """ & cmd & """ is not not allowed by the admin," & nl
				txt = txt & "please send help to the list for further information." & nl & nl
				txt = txt & "Regards" & nl & thisadmin
				tmpMessage.body = txt
				tmpMessage.Save
				write_log (s1 & "Sending notification")
			End If
			do_emailsubscription = true
		elseIf cmd = LCase(subject_help) Then
			write_log (s2 & "Sending help to " & mailfrom)
			Set tmpMessage = CreateObject("hMailServer.Message")
			tmpMessage.From = msg_from
			tmpMessage.FromAddress = msg_fromaddress
			tmpMessage.AddRecipient mailfrom, mailfrom
			tmpMessage.Subject = "Help and available commands for " & mailto
			txt = "Hello " & mailfrom & nl & nl
			txt = txt & "This email describes the available commands for the list " & mailto & "." & nl & nl
			txt = txt & "List of all possible commands:" & nl
			txt = txt & "  " & subject_help & ":  send you this help" & nl
			txt = txt & "  " & subject_list & ":  return the members of the list" & nl
			txt = txt & "  " & subject_subscription & ":  add your email address to the list" & nl
			txt = txt & "  " & subject_subscription_address & ":  add the single email address given in the message body" & nl
			txt = txt & "  " & subject_subscription_list & ":  add many email addresses, each given as single line in the message body" & nl
			txt = txt & "  " & subject_unsubscription & ":  delete your email address from the list" & nl
			txt = txt & "  " & subject_unsubscription_address & ":  delete the single email address in the message body" & nl
			txt = txt & "  " & subject_unsubscription_list & ":  delete many email addresses, each given as a single line in the message body" & nl & nl
			txt = txt & "The command is given in the subject," & nl & "and any required address(es) in the body of the email." & nl
			txt = txt & "Adding and deleting large lists may take some time -- lists above 100 members are not recommended!" & nl & nl
			txt = txt & "Allowed commands for list " & mailto & ":" & nl & "  " & Replace(cmdald,"#", nl & "  ") & nl
			txt = txt & "Other commands are forbidden by the admin." & nl & nl
			txt = txt & "Regards" & nl & thisadmin
			tmpMessage.body = txt
			tmpMessage.Save
			do_emailsubscription = true
		elseIf cmd = LCase(subject_list) Then
			write_log (s2 & "Sending list to " & mailfrom)
			Set tmpMessage = CreateObject("hMailServer.Message")
			tmpMessage.From = msg_from
			tmpMessage.FromAddress = msg_fromaddress
			tmpMessage.AddRecipient mailfrom, mailfrom
			tmpMessage.Subject = "Members of list " & mailto
			txt = "Hello " & mailfrom & nl & nl
			txt = txt & "this emails contains the members of list " & mailto & "." & nl & nl
			txt = txt & "Regards" & nl & thisadmin & nl & nl & nl
			txt = txt & "All members:" & nl
			txt = txt & Replace(allrcpts,"#", nl) 
			tmpMessage.body = txt
			tmpMessage.Save
			do_emailsubscription = true
		elseIf cmd = LCase(subject_subscription) Then
			if InStr(1, "  " & allrcpts, mailfrom) < 1  then 
				Set SRecipient = sRecipients.Add  
				SRecipient.RecipientAddress = LCase(mailfrom)
				SRecipient.Save 
				write_log (s2 & a1 & mailfrom & " has been added to list " & mailto)
				If thissuccessnotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = "Command " & cmd & " to list " & mailto & " successful"
					tmpMessage.Save
					write_log (S2 & "Sending notification")
				End If
			else
				write_log (s2 & mailfrom & " is already member of list " & mailto)
				If thisfailurenotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = "Already member of list " & mailto
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
			end if 
			do_emailsubscription = true
		elseIf cmd = LCase(subject_subscription_address) Then
			arr = Split(omsg.body,nl)
			tmp = Trim(LCase(arr(0)))
			If tmp = "" And UBound(arr) >= 1 then ' PeterWeb - empty first line? Try second (some mail senders have extra blank line)
				write_log (s2 & "Tried to get address from second line")
  				tmp = Trim(LCase(arr(1)))
			End if
			
			If Not obApp.Utilities.IsValidEmailAddress(tmp) then
				write_log (s1 & tmp & " is not a valid emailaddress, not added to " & mailto)
				If thisfailurenotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = tmp & " is not a valid emailaddress"
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
			ElseIf InStr(1, "  " & allrcpts, tmp) < 1 then 
				Set SRecipient = sRecipients.Add  
				SRecipient.RecipientAddress = tmp 
				SRecipient.Save 
				write_log (s1 & tmp & " has been added to list " & mailto)
				If thissuccessnotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = "Command " & cmd & " to list " & mailto & " successful"
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
			else
				write_log (s1 & tmp & " is already member of list " & mailto)
				If thisfailurenotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = tmp & " already member of list " & mailto
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
			end if 
			do_emailsubscription = true
		elseIf cmd = LCase(subject_subscription_list) Then
			arr = Split(omsg.body,nl)
			
			txt = ""
			For i = 0 To UBound(arr)
				tmp = Trim(LCase(arr(i)))
				If Len(tmp) = 0 Then
					write_log (s2 & "Ignoring empty line")
				elseIf Len(tmp) < 5 Or Mid(tmp,1,1) = "#" then
					write_log (s2 & "Ignoring line " & tmp)
					txt = txt & spc_fill_up_string(tmp & " has been ignored",50) & "ERROR" & nl
				elseIf Not obApp.Utilities.IsValidEmailAddress(tmp) Then
					write_log (s2 & a1 & tmp & "  is not a valid emailaddress")
					txt = txt & spc_fill_up_string(tmp & " is not a valid emailaddress", 50) & "ERROR" & nl
				elseIf InStr(1, "  " & allrcpts, tmp) > 0 Then
					write_log (s2 & tmp & " is already member of list " & mailto)
					txt = txt & spc_fill_up_string(tmp & " is already member", 50) & "ERROR" & nl
				Else
					Set SRecipient = sRecipients.Add  
					SRecipient.RecipientAddress = tmp 
					SRecipient.Save
					allrcpts = allrcpts & "#" & tmp
					write_log (s2 & tmp & " has been added to list " & mailto)
					txt = txt & tmp & " has been added" & nl
				End if
			Next
			
			If thissuccessnotification = "true" Then
				Set tmpMessage = CreateObject("hMailServer.Message")
				tmpMessage.From = msg_from
				tmpMessage.FromAddress = msg_fromaddress
				tmpMessage.AddRecipient mailfrom, mailfrom
				tmpMessage.Subject = "Report of command " & cmd & " to list " & mailto
				tmpMessage.Body = txt
				tmpMessage.Save
				write_log (s2 & "Sending notification")
			End If
			do_emailsubscription = true
		ElseIf LCase(omsg.subject) = LCase(subject_unsubscription) Then
   			sRecipientID = 0
		    Dim j
      		For j = 0 to sRecipients.Count - 1
      			Set sTempRecipient = sRecipients.Item(j)
      			write_log(s2 & j)
      			write_log(s2 & sTempRecipient.RecipientAddress & " = " & mailfrom)
            	if LCase(sTempRecipient.RecipientAddress) = LCase(mailfrom) then 
               		sRecipientID = sTempRecipient.ID 
         			j = sRecipients.Count 
       			end if 
   			next 
    
   			if sRecipientID <> 0 then 
      			obList.Recipients.DeleteByDBID(sRecipientID) 
				write_log (s2 & mailfrom & " has been deleted from list " & mailto)
				If thissuccessnotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = "Command " & cmd & " to list " & mailto & " successful"
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
   			else    
				write_log (s2 & mailfrom & " is not a member of list " & mailto)
				If thisfailurenotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress 
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = "Not a member of list " & mailto
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
   			end if 
			do_emailsubscription = true
		elseIf cmd = LCase(subject_unsubscription_address) Then
			arr = Split(omsg.body,nl)
			tmp = Trim(LCase(arr(0)))
			If tmp = "" And UBound(arr) >= 1 then ' PeterWeb - empty first line? Try second (some mail senders have extra blank line)
				write_log (s2 & "Tried to get address from second line")
  				tmp = Trim(LCase(arr(1)))
			End if
			
   			sRecipientID = 0 
      		For j = 0 to sRecipients.Count  - 1
      			Set sTempRecipient = sRecipients.Item(j) 
            	if LCase(sTempRecipient.RecipientAddress) = LCase(tmp) then 
               		sRecipientID = sTempRecipient.ID 
         			j = sRecipients.Count 
       			end if 
   			next 
			
 			If Not obApp.Utilities.IsValidEmailAddress(tmp) then
				write_log (s2 & tmp & " is not a valid emailaddress, not deleted from " & mailto)
				If thisfailurenotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = tmp & " is not a valid emailaddress"
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
  			elseif sRecipientID <> 0 then 
      			obList.Recipients.DeleteByDBID(sRecipientID) 
				write_log (s2 & tmp & " has been deleted from list " & mailto)
				If thissuccessnotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = "Command " & cmd & " to list " & mailto & " successful"
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
   			else    
				write_log (s2 & tmp & " is not a member of list " & mailto)
				If thisfailurenotification = "true" Then
					Set tmpMessage = CreateObject("hMailServer.Message")
					tmpMessage.From = msg_from
					tmpMessage.FromAddress = msg_fromaddress 
					tmpMessage.AddRecipient mailfrom, mailfrom
					tmpMessage.Subject = tmp & " is not a member of list " & mailto
					tmpMessage.Save
					write_log (s2 & "Sending notification")
				End If
   			end if 
			do_emailsubscription = true
		elseIf cmd = LCase(subject_unsubscription_list) Then
			arr = Split(omsg.body,nl)
			txt = ""
			For i = 0 To UBound(arr)
				tmp = Trim(LCase(arr(i)))
				If Len(tmp) = 0 Then
					write_log (s2 & "Ignoring empty line")
				elseIf Len(tmp) < 5 Or Mid(tmp,1,1) = "#" then
					write_log (s2 & "Ignoring line " & tmp)
					txt = txt & spc_fill_up_string(tmp & " has been ignored",50) & "ERROR" & nl
				elseIf Not obApp.Utilities.IsValidEmailAddress(tmp) Then
					write_log (s2 &  tmp & "  is not a valid emailaddress")
					txt = txt & spc_fill_up_string(tmp & " is not a valid emailaddress",50) & "ERROR" & nl
				elseIf InStr(1, "  " & allrcpts, tmp) < 1 Then
					write_log (s2 &  tmp & " is not in list " & mailto)
					txt = txt & spc_fill_up_string(tmp & " is not a member",50) & "ERROR" & nl
				Else
		   			sRecipientID = 0 
		      		for j = 0 to sRecipients.Count  - 1
		      			Set sTempRecipient = sRecipients.Item(j) 
		            	if LCase(sTempRecipient.RecipientAddress) = LCase(tmp) then 
		               		sRecipientID = sTempRecipient.ID 
		         			j = sRecipients.Count 
		       			end if 
		   			next 
					allrcpts = Replace(allrcpts, tmp, "")
					if sRecipientID <> 0 then 
      					obList.Recipients.DeleteByDBID(sRecipientID) 
						write_log (s2 &  tmp & " has been deleted from list " & mailto)
						txt = txt & tmp & " has been deleted" & nl
					else
						write_log (s2 &  tmp & " could not be found in list " & mailto)
						txt = txt & spc_fill_up_string(tmp & " could not be found",50) & "ERROR" & nl
					End if
				End if
			Next
			
			If thissuccessnotification = "true" Then
				Set tmpMessage = CreateObject("hMailServer.Message")
				tmpMessage.From = msg_from
				tmpMessage.FromAddress = msg_fromaddress
				tmpMessage.AddRecipient mailfrom, mailfrom
				tmpMessage.Subject = "Report of command " & cmd & " to list " & mailto
				tmpMessage.Body = txt
				tmpMessage.Save
				write_log (s2 & "Sending notification")
			End If
			do_emailsubscription = true
		else
			write_log (s1 & " Email subject does not contain a valid command.")
		End if
	Else
		write_log (s1 & "This is not a distribution list!")
	End If
End function




Function get_pw(str, mailto)
	Dim inpw
	Dim pw
	inpw = False
	pw = ""
	dim thispwe
	thispwe = get_config_string(mailto, "pw_enclosure")
	
    Dim i
	For i = 1 To Len(str)
		If Mid(str,i,1) = thispwe And inpw = true Then
			inpw = False
		ElseIf Mid(str,i,1) = thispwe And inpw = False Then
			inpw = True
		elseif inpw = True Then 
			pw = pw & Mid(str,i,1)
		End if
	Next
	If pw <> "" Then 
		get_pw = thispwe & get_config_string(mailto, "pw") & thispwe
	End If
End Function

Function create_pw(mailto)
	dim thispwe
	thispwe = get_config_string(mailto, "pw_enclosure")
	create_pw = thispwe & get_config_string(mailto, "pw") & thispwe
End Function

'------------------------------------------------------------------
' General functions of all scripts
'------------------------------------------------------------------

Sub write_log(txt)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fs
	Dim f
	Dim fn
	Dim fnl
	Dim tmp 
	
	If write_log_active then
		Set fs = CreateObject("scripting.filesystemobject")
		
		fn = logspath & "listserv_event_" & get_date & ".log"
		fnl = logspath & "listserv_event_" & get_date & ".lock"
		
		If fs.FileExists(fnl) then
		    Dim i
			For i = 0 To 25
				log_delay
				If Not fs.FileExists(fnl) Then
					i = 10000
				End If
			next
		End If
		
		If Not fs.FileExists(fnl) then
			Set f = fs.opentextfile(fnl, ForWriting, true)
			f.Write("")
			f.Close
			Set f = fs.opentextfile(fn, ForAppending, true)
			'It would be cleaner without the quotes, but they seem to be used as discovery tokens in SMTP_log_position
			tmp = """" & FormatDateTime(Date + time,0) & """" & Chr(9) & """" & txt & """" & nl	
			'tmp = FormatDateTime(Date + time,0) & "   " & txt & nl
			on error resume next
			f.Write(tmp)
			on error goto 0
			f.Close
			fs.DeleteFile(fnl)
		End If
	End if
End Sub

Sub log_delay()
	Dim wscr
	Set wscr = CreateObject("WScript.Shell")
	wscr.Popup "Waiting...", 1,"Waiting..."
End sub

Function get_date
	Dim tmp
	Dim erg
	tmp = Year(Date)
	erg = CStr(tmp)
	
	If Month(Date) < 10 Then
		tmp = "0" & Month(Date)
	Else
		tmp = Month(Date)
	End If
	erg = erg & "-" & tmp
	
	If day(Date) < 10 Then
		tmp = "0" & day(Date)
	Else
		tmp = day(Date)
	End If
	erg = erg & "-" & tmp
	
	get_date = erg
End Function

Function nl
	nl = Chr(13) & Chr(10)
End function

Function spc_fill_up_string(input, pos)
	If pos - Len(input) > 0 then
		spc_fill_up_string = input & Space(pos - Len(input))
	Else
		spc_fill_up_string = input & " "
	End if
End function

Function is_local_domain(domain_or_email)
	is_local_domain = False
	Dim domain
	Dim doms
	Dim alss
	Dim i
	Dim j
	Dim dom
	Dim als
	
	If InStr(1,"  " & domain_or_email,"@") > 0 Then
		domain = Mid(domain_or_email, InStr(1,domain_or_email,"@") + 1)
	Else
		domain = domain_or_email
	End If
	
	If domain_buffer = "" then
		i = 0
		Set doms = obapp.Domains
		Do While i <= doms.Count - 1
			Set dom = doms.Item(i)
			domain_buffer = domain_buffer & "#" & dom.Name
			j = 0
			Set alss = dom.DomainAliases
			Do While j <= alss.Count - 1
				Set als = alss.item(j)
				domain_buffer = domain_buffer & "#" & als.AliasName
				j = j + 1
			Loop
			i = i + 1
		Loop
	End If
	
	If InStr(1, "  " & domain_buffer, domain) > 0 Then
		is_local_domain = True
	End If
End Function

Function is_local_account(emailaddress)
	is_local_account = False
	
	Dim domain
	Dim doms
	dim dom
	dim als
	dim fnd
	dim accs
	dim alss
	dim j
	dim i
	
	If InStr(1,"  " & emailaddress,"@") > 0 Then
		domain = Mid(emailaddress, InStr(1,emailaddress,"@") + 1)
	Else
		domain = emailaddress
	End If
	
	if is_local_domain(domain) then
		Set doms = obapp.Domains
		i = 0
		Do While i <= doms.Count - 1
			Set dom = doms.Item(i)
			if lcase(dom.name) = lcase(domain) then
				fnd = true
			end if
			j = 0
			Set alss = dom.DomainAliases
			Do While j <= alss.Count - 1
				Set als = alss.item(j)
				if lcase(als.AliasName) = lcase(domain) then
					fnd = true
				end if
				j = j + 1
			Loop
			if fnd then
				i = doms.Count + 10
			end if
			i = i + 1
		Loop
		i = 0
		set accs = dom.accounts
		Do While i <= accs.count - 1
			if lcase(accs(i).Address) = lcase(emailaddress) then
				is_local_account = true
				i = accs.count + 10
			end if
			i = i + 1
		Loop
	end if
End function

Function is_local_list(emailaddress)
	is_local_list = False
	Dim fnd
	Dim suffix
	Dim doms
	Dim obDomain
	suffix = Mid(emailaddress, InStr(1,emailaddress,"@") + 1)
	fnd = False
	Set doms = obApp.Domains
    Dim i
	For i = 0 To doms.Count - 1
		If LCase(doms.Item(i).Name) = LCase(suffix) Then
			fnd = True
		End if
	Next
	If fnd = True Then
		Set obDomain = doms.ItemByName(suffix)
		For i = 0 To obDomain.DistributionLists.Count - 1
			If LCase(emailaddress) = LCase(obDomain.DistributionLists.Item(i).Address) Then
				is_local_list = True
			End if
		Next
	End If
End function

Function is_member_of_local_list(member,list)
	is_member_of_local_list = False
	Dim fnd
	Dim is_list
	Dim doms
	Dim obDomain
	Dim obList
	Dim obrecList
	fnd = False
	is_list = false
	Set doms = obApp.Domains
    Dim i
	For i = 0 To doms.Count - 1
		If LCase(doms.Item(i).Name) = LCase(Mid(list, InStr(1,list,"@") + 1)) Then
			fnd = True
		End if
	Next
	If fnd = True Then
		Set obDomain = doms.ItemByName(Mid(list, InStr(1,list,"@") + 1))
		For i = 0 To obDomain.DistributionLists.Count - 1
			If LCase(list) = LCase(obDomain.DistributionLists.Item(i).Address) Then
				is_list = True
			End if
		Next
	End If
	If is_list then
		Set obList = obDomain.DistributionLists.ItemByAddress(list)
		Set obreclist = obList.Recipients
		For i = 0 to obreclist.count - 1
			If LCase(obreclist.item(i).RecipientAddress) = LCase(member) Then
				is_member_of_local_list = True
			End if
		Next
	End if
End Function

Function has_client_authenticated(oclient)
	has_client_authenticated = false
	If oCLient.username <> "" Or InStr(1,"  " & ipslocalhost, oClient.IPAddress) > 0 Then
		has_client_authenticated = true
	End if
End Function

Function has_client_authenticated_man(ipaddr,usern)
	has_client_authenticated_man = false
	If usern <> "" Or InStr(1,"  " & ipslocalhost, ipaddr) > 0 Then
		has_client_authenticated_man = true
	End if
End Function

Function get_smtp_recipient(oMessage, IPAddress)
	Dim tmp
	tmp = ""
	Dim erg
	Dim i
	erg = ""
	If using_v5 Then
		For i = 0 To oMessage.Recipients.count - 1
			tmp = oMessage.Recipients(i).OriginalAddress
			If InStr(1,"  " & erg,tmp) < 1 Then
				erg = erg & tmp & "#"
			End If
		Next
		If Len(erg) > 0 Then
			erg = Mid(erg,1,Len(erg) - 1)
		End If
		get_smtp_recipient = erg
	Else
		get_smtp_recipient = get_smtp_recipient_log(IPAddress)
	End If
End Function

Function get_smtp_recipient_log(ipaddr)
    get_smtp_recipient_log = ""
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Dim fs
    Dim f
    Set fs = CreateObject("scripting.filesystemobject")
    
    Dim fn
    Dim ln
    Dim arr
    Dim tmp
    Dim erg
    erg = ""
    Dim startwithpos
    startwithpos = 0
    Dim currpos
    currpos = 0
    Dim nextstartpos
    nextstartpos = 0
    
    If fs.FileExists(SMTP_log_position) Then
        Set f = fs.opentextfile(SMTP_log_position,ForReading)
    	If Not f.AtEndOfStream Then
    		ln = f.ReadLine
    	End If
    	If ln = CStr(CLng(Date())) And Not f.AtEndOfStream Then
    		ln =f.ReadLine
    		startwithpos = CLng(ln) - delta_pos_log
    	End If
    	f.Close
    End If
    
    If startwithpos < 0 Then
    	startwithpos = 0
    End If
    
    fn = logspath & "hmailserver_" & get_date & ".log"
    
    If fs.FileExists(fn) Then
        Set f = fs.opentextfile(fn,ForReading)
        currpos = startwithpos
        f.skip(startwithpos)
        Do While Not f.AtEndOfStream
            ln = f.ReadLine
            currpos = currpos + Len(ln) + 2
            arr = Split(ln, Chr(9))
            If UBound(arr) = 5 Then
                tmp = ""
                If arr(0) = """SMTPD""" And arr(4) = """" & ipaddr & """" And InStr(1, "  " & arr(5), "RECEIVED: EHLO") > 0 Then
                    erg = ""
                End If
                If arr(0) = """SMTPD""" And arr(4) = """" & ipaddr & """" And InStr(1, "  " & arr(5), "RECEIVED: HELO") > 0 Then
                    erg = ""
                End If
                If arr(0) = """SMTPD""" And arr(4) = """" & ipaddr & """" And InStr(1, "  " & arr(5), smtp_log_search_string) > 0 Then
                    erg = ""
                End If
                If arr(0) = """SMTPD""" And arr(4) = """" & ipaddr & """" And InStr(1, arr(5), "RCPT") > 0 And InStr(1, arr(5), "<") > 0 And InStr(1, arr(5), ">") > 0 Then
                    tmp = Mid(arr(5), InStr(1, arr(5), "<") + 1)
                    tmp = Mid(tmp, 1, InStr(1, tmp, ">") - 1)
                    nextstartpos = currpos
                End If
                If erg = "" And tmp <> "" Then
                	erg = tmp & "#"
                elseIf tmp <> "" Then
                    If InStr(1,"   " & erg,tmp) <= 1 Then
                    	erg = erg & tmp & "#"
                    End if
                End If
            End If
        Loop
        f.Close
    End If
    
    If erg <> "" Then
        get_smtp_recipient_log = Mid(erg, 1, Len(erg) - 1)
        
		Set f = fs.OpenTextFile(SMTP_log_position,ForWriting,True)
		f.Write(CLng(Date) & nl & nextstartpos)
		f.Close 
    End If
End Function

Function msg_file_content_read(oMessage)
	msg_file_content_read = ""
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fs , f
	Set fs = CreateObject("scripting.filesystemobject")
	If fs.FileExists(oMessage.Filename) then
		Set f = fs.OpenTextFile(oMessage.Filename, ForReading)
		msg_file_content_read = f.Readall
		f.Close
	End If
End Function

sub msg_file_content_write(oMessage, content)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fs , f
	Set fs = CreateObject("scripting.filesystemobject")
	Set f = fs.OpenTextFile(oMessage.Filename, ForWriting, true)
	f.Write(content)
	f.Close
	oMessage.RefreshContent
End sub

Function get_file_suffix(str)
	Dim cr
	Dim erg
	erg = ""
    Dim i
	For i = 1 To Len(str)
		cr = Mid(str,i,1)
		If cr = "." Then
			erg = ""
		Else
			erg = erg & cr
		End if
	Next
	If erg = "" Then
		erg = "???"
	ElseIf erg = str Then
		erg = "none"
	End If
	get_file_suffix = Trim(erg)
End Function



Function ReplaceTokens(str, mailto, mailfrom)
	Dim tmp
	Dim erg
	erg=""
	erg = str
	erg = Replace(erg,"%admin%", get_config_string(mailto, "admin"))
	erg = Replace(erg,"%serveradmin%", serveradmin)
	erg = Replace(erg,"%address%", mailto)
	erg = Replace(erg,"%mailto%", mailto)
	erg = Replace(erg,"%mailfrom%", mailfrom)
	tmp = Mid(mailto, 1, InStr(1, mailto,"@") - 1)
	erg = Replace(erg,"%prefix%", tmp)
	tmp = Mid(mailto, InStr(1,mailto,"@") + 1)
	erg = Replace(erg,"%domain%", tmp)
	erg = Replace(erg,"%date%", formatdatetime(Date(), 2))
	erg = Replace(erg,"%time%", formatdatetime(Time(), 3))
	erg = Replace(erg, "%list%", "list")	'BGM Friday, March 01, 2013 - the original script uses "list" as a token rather than "%list%" - but it should be standardized throughout the script.  This add allows for both.  There is already a token for "%address% which should translate also to the list address

	ReplaceTokens = erg
End Function

'----------------------------------------------------------------------------------------------------
'This function reads the contents of the generalposters list and appends it to the general-allowed posters; see top of script
Function GetGeneralAdmins(mailto, mailfrom)
	dim adminstring, templist
	adminstring = ""
	
	generalposterslist = ReplaceTokens(generalposterslist, mailto, mailfrom)	'we need to update the %domain% token
	
	Dim fs , f
	Set fs = CreateObject("scripting.filesystemobject")
	templist = generalposterslist
	If fs.FileExists(templist) Then
		Set f = fs.OpenTextFile(templist)
		adminstring = f.readall
		adminstring = Replace(adminstring, chr(013), "#")  'carriage return
		adminstring = Replace(adminstring, chr(010), "")   'new line
		f.Close
	End If
	adminstring = general_allowed_list & "#" & adminstring
	GetGeneralAdmins = adminstring
End Function	