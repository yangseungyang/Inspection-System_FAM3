@echo off
rem <ドライブCの検査データをネットワークサーバにコピー>
rem xcopy "C:\Users\FAM3-ATE-J\Desktop\K号機" "E:\ATE_K\K号機\" /D /S /E /C /F /H /R /K /Y /Z
xcopy "C:\Users\FAM3-ATE-J\Desktop\K号機" "\\10.36.14.215\fam3\ATE_K\K号機\" /D /S /E /C /F /H /R /K /Y /Z
rem pause
