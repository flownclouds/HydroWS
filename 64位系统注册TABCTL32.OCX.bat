@echo 开始注册
copy TABCTL32.OCX %windir%\SysWOW64\
regsvr32 %windir%\SysWOW64\TABCTL32.OCX /s
@echo TABCTL32.OCX注册成功
@pause