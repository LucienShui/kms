@echo off
title KMS Activation - Version 2022.2.16

set kms_server=kms.example.com

goto :menu

:menu
cls

echo ---------------------------------------------------------
echo KMS Activation
echo KMS ����ű�
echo.
echo Input 1 to Activate Windows Opreating System
echo ���� 1 �Լ��� Windows ����ϵͳ
echo.
echo Input 2 to Activate Office 2010
echo ���� 2 �Լ��� Office 2010
echo.
echo Input 3 to Activate Office 2013
echo ���� 3 �Լ��� Office 2013
echo.
echo Input 4 to Activate Office 2016/2019
echo ���� 4 �Լ��� Office 2016 / 2019
echo.
echo Input 0 to exit
echo ���� 0 �˳�
echo.
echo ---------------------------------------------------------

:cho 
set choice=
set /p choice=Your choice����ѡ�񣩣�

IF NOT "%Choice%"=="" SET Choice=%Choice:~0,1%
if /i "%choice%"=="1" goto 1
if /i "%choice%"=="2" goto 2
if /i "%choice%"=="3" goto 3
if /i "%choice%"=="4" goto 4
if /i "%choice%"=="0" goto end

echo Your input is incorrect, please press any key to retype
echo ���벻��ȷ������������ԡ� & pause
goto menu 

:1
cscript C:\windows\system32\slmgr.vbs -skms %kms_server%
cscript C:\windows\system32\slmgr.vbs -ato
echo If no error is displayed, the activation is successful.
echo ���޴�����ʾ���򼤻�ɹ�. & pause
goto :menu

:2
cscript "C:\Program Files\Microsoft Office\Office14\ospp.vbs" /sethst:%kms_server%
cscript "C:\Program Files\Microsoft Office\Office14\ospp.vbs" /act
echo If no error is displayed, the activation is successful.
echo ���޴�����ʾ���򼤻�ɹ�. & pause
goto :menu

:3
cscript "C:\Program Files\Microsoft Office\Office15\ospp.vbs" /sethst:%kms_server%
cscript "C:\Program Files\Microsoft Office\Office15\ospp.vbs" /act
echo If no error is displayed, the activation is successful.
echo ���޴�����ʾ���򼤻�ɹ�. & pause
goto :menu

:4
cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /sethst:%kms_server%
cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /act
echo If no error is displayed, the activation is successful.
echo ���޴�����ʾ���򼤻�ɹ�. & pause
goto :menu
:END
