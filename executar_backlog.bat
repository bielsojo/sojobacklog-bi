@echo off
cls
echo ===================================================
echo   Sincronizando Tecnicos do Backlog com o GLPI...
echo ===================================================
echo.
python process_backlog.py
echo.
echo ===================================================
echo   Processo concluido! Pressione qualquer tecla para sair.
echo ===================================================
pause > nul
