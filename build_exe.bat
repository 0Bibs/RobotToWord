@echo off
REM Gera dist\RobotToWord.exe - um unico arquivo, sem precisar de Python na
REM maquina de quem for usar. Rode este .bat uma vez, em uma maquina Windows
REM que tenha Python instalado.

cd /d "%~dp0"

echo Instalando dependencias...
py -m pip install --upgrade pyinstaller -r requirements.txt || goto :erro

echo.
echo Empacotando...
py -m PyInstaller --noconfirm --onefile --windowed --name RobotToWord painel.py || goto :erro

echo.
echo Pronto: dist\RobotToWord.exe
echo Copie esse arquivo para a pasta da rede e a equipe usa com dois cliques.
pause
exit /b 0

:erro
echo.
echo Falhou. Confira se o Python esta instalado e no PATH (py --version).
pause
exit /b 1
