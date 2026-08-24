@echo off
cd /d "C:\Users\sorbonne\Documents\Workspace\Risala"
start python -m http.server 8888
timeout /t 2
start http://localhost:8888/risala.html