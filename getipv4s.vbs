''              
''             returns comma seperated list of all ipv4 addresses
''
function getipv4s()
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set IPConfigSet = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'")
	ips = ""
	for each ipconfig in ipconfigset
		if not isnull(ipconfig.ipaddress) then
			for i = lbound(ipconfig.ipaddress) to ubound(ipconfig.ipaddress)
				if not instr(ipconfig.ipaddress(i), ":") > 0 then
					if ips <> "" then ips = ips & ","
					ips = ips & ipconfig.ipaddress(i)
				end if
			next
		end if
	next
	getipv4s = ips
end function