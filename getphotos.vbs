'On Error Resume Next 

const HKEY_CURRENT_USER = &H80000001
 

strComputer = "."
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Photo Acquisition"
 

Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv") 



DeleteSubkeys HKEY_CURRENT_USER, strKeyPath 

deletefile "C:\Users\Administrator\AppData\Local\Microsoft\Photo Acquisition\PreviouslyAcquired.db"

Set objshell = WScript.CreateObject("WScript.Shell")
 
objshell.Run "rundll32.exe ""C:\Program Files\Windows Photo Viewer\PhotoAcq.dll"" PhotoAndVideoAcquireW"


Sub DeleteSubkeys(HKEY_CURRENT_USER, strKeyPath) 
	objRegistry.EnumKey HKEY_CURRENT_USER, strKeyPath, arrSubkeys 
	If IsArray(arrSubkeys) Then 
		For Each strSubkey In arrSubkeys 
			DeleteSubkeys HKEY_CURRENT_USER, strKeyPath & "\" & strSubkey 
		Next 
	End If 
	objRegistry.DeleteKey HKEY_CURRENT_USER, strKeyPath 
End Sub

'----------------------------------------------------------------------------------------------
'
'	delete a file ( if it does not exist then it will not croak )
'
function deletefile( target )

 	set fso = createobject("scripting.filesystemobject")
	if ( fso.fileexists( target ) ) then 
		fso.Deletefile target, True
	end if

end function