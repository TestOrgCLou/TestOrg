@echo off

copy MSFLXGRD.DEP %windir%\system32\
copy Msflxgrd.ocx %windir%\system32\
regsvr32 %windir%\system32\Msflxgrd.ocx /s

echo ע��ɹ����˳�����

exit