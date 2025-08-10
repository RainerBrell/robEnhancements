@echo off
cls
rem 9. September 2024 - Create a new NVDA Add-On 
set Ord=addon\
set name=robEnhancements
set pot=%ord%locale\en\%name%.pot 

Echo stepp 1:
del addon\*.html /s 
del addon\*.mo /s 
pause 
echo stepp 2:
copy addon\doc\en\readme.md readme.md 
pause 
echo stepp 3:
scons
pause 