@echo off

copy MSFLXGRD.DEP %windir%\SysWOW64\
copy Msflxgrd.ocx %windir%\SysWOW64\
regsvr32 %windir%\SysWOW64\Msflxgrd.ocx /s

echo ×¢²á³É¹¦£¬ÍË³ö¡£¡£

exit