@echo off
title Outlook Inbox Manager
cd /d C:\Users\dirk\OutlookSort

python outlook_manager.py
if errorlevel 1 pause
