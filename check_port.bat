@echo off
setlocal
cd /d "%~dp0"
echo Checking common dashboard ports...
for %%P in (5055 8050 8088 8765 9000 5000) do (
  echo.
  echo === PORT %%P ===
  netstat -ano | findstr ":%%P "
)
echo.
echo Windows excluded TCP port ranges:
netsh interface ipv4 show excludedportrange protocol=tcp
pause
endlocal
