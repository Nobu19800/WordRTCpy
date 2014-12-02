set mPath="%ProgramFiles%\SEATSAT"
IF NOT EXIST %mPath% (
	set mPath="%ProgramFiles(x86)%\SEATSAT"	
)

%mPath%\SEAT.exe test.seatml
exit