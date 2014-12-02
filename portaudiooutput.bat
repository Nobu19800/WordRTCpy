set mPath="%ProgramFiles%\OpenHRIAudio"
IF NOT EXIST %mPath% (
	set mPath="%ProgramFiles(x86)%\OpenHRIAudio"	
)

cd %mPath%
portaudiooutput.exe -f rtc.conf
exit