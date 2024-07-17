@echo off
powershell -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force"
start powershell -WindowStyle Hidden -File "run.ps1"
powershell -Command  $OutputEncoding = [System.Text.Encoding]::UTF8