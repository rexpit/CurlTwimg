@echo off
powershell -ExecutionPolicy RemoteSigned -File "%~d0%~p0\curlTwimg.ps1" %1
