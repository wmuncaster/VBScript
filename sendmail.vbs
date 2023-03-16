function sendmail( fromaddr, toaddr, subject, body, smtpservername, smtpserverport )
    
    Set objEmail = CreateObject( "CDO.Message" )
    With objEmail
        .From     = fromaddr
        .To       = toaddr
        .Subject  = subject
        .TextBody = body
		root = "http://schemas.microsoft.com/cdo/configuration/"
        With .Configuration.Fields
            .Item( root & "sendusing"      ) = 2
            .Item( root & "smtpserver"     ) = smtpservername
            .Item( root & "smtpserverport" ) = smtpserverport
            .Update
        End With
		on error resume next
        .Send
		'wscript.echo "result:(" & err.description
    End With
    ' Return status message
	
    If Err Then
        sendmail = err.description
    Else
        sendmail = ""
    End If

end function 