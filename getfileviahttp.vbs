
'		author: wmuncaster 
'
'		url: where to get, binaryfile = true if binaryfile, flase if text
'		if targetfile is blank and binmode = false then the requested text file results returned		
'		if targetfile is NOT blank, the requested file is saved to targetfile 
'

function getfileviahttp( url, targetfile, binaryfile )
		on error resume next

		Set objhttp = CreateObject( "WinHttp.WinHttpRequest.5.1" ) 
		objhttp.SetTimeouts 60000, 60000, 1200000, 1200000
		objhttp.open "get", url
		objhttp.Send 
		status = objhttp.Status

		if ( status <> 200 ) then
			getfileviahttp = ""
			exit function
		end if

 		if ( binaryfile = False ) then
			if ( targetfile <> "" ) then 
				writefile targetfile, objhttp.ResponseText, 2 
				exit function
			else 
				getfileviahttp = objhttp.ResponseText
				exit function
			end if 
		end if
		
		if ( binaryfile = true ) then 
			Set BinaryStream	= CreateObject("ADODB.Stream")
			BinaryStream.Type	= 1
			BinaryStream.Open
			BinaryStream.Write objhttp.ResponseBody
			BinaryStream.SaveToFile targetfile, 2
			getfileviahttp = "local file downloaded as " & targetfile
			exit function
		end if
	getfileviahttp = ""
end function

function writefile( filepath, buffer, mode ): on error resume next:with createobject("Scripting.FileSystemObject"):with .opentextfile(  filepath, mode, True ):.writeline( buffer ):.close:end with:end with:on error goto 0:end function
