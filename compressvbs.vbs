  
 
function gulpfile( filename ):dim lines:on error resume next:with createobject("scripting.filesystemobject"):with .opentextfile(filename,1):lines=.readall:.close:end with:end with:on error goto 0:gulpfile=lines:end function


	inputfile = wscript.arguments(0)
	
 	'
	'
	'	replace all end of lines with :
	'
	'
	for each line in split( gulpfile( inputfile ), vbcrlf )
		'
		'	eliminate leading/trailing spaces and tabs from line
		'
		for x = 1 to len ( line )
			c = mid(line,x,1)
			if ( c <> " " ) and ( c <> vbtab ) then exit for
		next		
		line = mid( line, x )
		
		for x = len ( line ) to 0 step -1
			c = left(line,x)
			if ( c <> " " ) then exit for
		next	
		
		line = left( line, x )

	'
	'	now, the leading and trailing spaces and tabs have been removed
	'
	'	now remove redundant spaces which are found outside of quoted strings
	'
	 	
		quotepairs = ""
		for x = 1 to len ( line )
			c = mid(line,x,1)
			if ( c = chr(34) )  then
				quotepairs = quotepairs & x & ","
				for y = x+1 to len ( line )
					c = mid(line,y,1)
					if ( c = chr(34) )  then
						quotepairs = quotepairs & y & "~"
						x = y
						exit for
					end if
				next
				
			end if
		next	
		
		'debugit "         1         2         3         4         5         6         7         8         9",""
		'debugit "123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 ",""
		'echo "DEBIG " & line
		'debugit "quotepairs",quotepairs
		
		
		replacers = split( " : ~:`" & ", ~,`( ~(` )~)` = ~=` - ~-` + ~+` & ~&`  ~ `" & vbtab & "~", "`" )
		quotepairs = split( quotepairs, "~" )
		
		for each replacer in replacers
			items = split( replacer, "~" )
			for x = 1 to len( line )
				x = instr( x, line, items(0) ) ' found the search for arg?
				if ( x = 0 ) then exit for
				instring = false
				if x then
					for each quotepair in quotepairs
						if (quotepair = "" ) then exit for
						qitems = split( quotepair, "," )
						'debugit "qitems(0)",qitems(0)
						'debugit "qitems(1)",qitems(1)
						if ( x >= qitems(0) and x <= qitems(1) ) then
							instring = true
							x = qitems(1)+1
							exit for
						end if
					next
					if ( instring = false ) then
						'
						'	we are not inside of a quoted literal so try to compress items
						'	x will be sitting on the match which is not inside a quoted string
						'
						line = replace(line, items(0), items(1) )
						x = x + len(items(1))+1
					end if
				end if
				
			next
		next
		
		'
		'	remove comments
		'
		p = instr(line, "'" )
		if ( p ) then 
			line = left( line, p-1 )
		end if
		
		if ( line <> "" ) then 
			 
			'code = code & ">" & line & "<" & vbcrlf
			code = code & line & ":"
		end if

	next
	
	echo code
 	
	 
	
	


	
