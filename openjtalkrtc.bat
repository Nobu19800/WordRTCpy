set mPath="%ProgramFiles%\OpenHRIVoice"
IF NOT EXIST %mPath% (
	set mPath="%ProgramFiles(x86)%\OpenHRIVoice"	
)

cd %mPath%
openjtalkrtc.exe -f rtc.conf
exit