ECHO ON
REM A batch script to execute Ernobot Python script
SET PATH=%PATH%;C:\Users\User\AppData\Local\Programs\Python\Python36\

CD %~dp0

python maurubot.py

PAUSE