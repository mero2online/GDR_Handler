pyinstaller GDR_Handler.py --onefile --windowed --add-data 'src;src' -i ".\src\gdr.ico" --splash ".\src\gdr.png" --exclude-module matplotlib