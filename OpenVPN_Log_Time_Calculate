Dim fso, InFile, buf, d, Month_s, Month, DateTimeStart, DateTimeEnd, S
Set fso = CreateObject("Scripting.FileSystemObject")
Set InFile = fso.OpenTextFile("ext_vdoshchenko.log", 1, False)
Set OutFile = fso.CreateTextFile("out.txt",True)
DateTimeStart = 0
DateTimeEnd = 0
S = 0
Do While Not InFile.AtEndOfStream
    buf = InFile.ReadLine
	Month_s = Mid(buf,5,3)
	Year_s = Mid(buf,21,4)
	Day_s = Mid(buf,9,2)	
	Month = CStr((InStr("JanFebMarAprMayJunJulAugSepOctNovDec", Month_s)-1)/3 + 1)
	Time_s = Mid(buf,12,8)
	If Mid(buf,4,1) = " " Then
	    d = CDate(Day_s & "." & Month & "." & Year_s & " " & Time_s)
	End If
	If InStr(buf,"Initialization") > 0 Then
	    DateTimeStart = d
	End If
	If InStr(buf,"exiting") > 0 Then
	    DateTimeEnd = d
		If DateTimeStart <> 0 And DateTimeEnd > DateTimeStart Then
	        i = DateDiff("n",DateTimeStart,DateTimeEnd)
			OutFile.WriteLine("Start: " & DateTimeStart & " End: " & DateTimeEnd & " Duration: " & i & " minutes")
			S = S + i
			DateTimeStart = 0
		End If
	End If
Loop
OutFile.WriteLine("Total duration: " & S \ 60 & " hours " & S Mod 60 & " minutes")
MsgBox("Total duration: " & S \ 60 & " hours " & S Mod 60 & " minutes")
InFile.Close
OutFile.Close
Set InFile = Nothing
Set fso = Nothing
