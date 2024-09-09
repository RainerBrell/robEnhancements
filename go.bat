@echo off
cls
rem 9. September 2024 - Create a new NVDA Add-On 
set Ord=addon\
set name=robEnhancements
set pot=%ord%locale\en\%name%.pot 
Echo stepp 1:
echo Liste.txt erstellen:
echo buildVars.py >liste.txt 
For /R %ord% %%f In (*.py) Do echo %%f >>liste.txt 
Echo POT Datei erstellen:
xgettext.exe -f liste.txt -d %name% -o %pot% -c 
Echo POT-Datei mit bestehenden PO-Dateien verschmelzen:
For /R %ord%locale\ %%f In (*.po) Do msgmerge -U --backup=none %%f %pot% 
del liste.txt 
pause 
Echo stepp 2:
del addon\*.html /s 
del addon\*.mo /s 
pause 
echo stepp 3:
copy addon\doc\en\readme.md readme.md 
pause 
echo stepp 3:
scons
pause 