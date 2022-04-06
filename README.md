## Usage

```bash
# Run script
python GDR_Handler.py


# Compiled with Pyinstaller

# Windows
pyinstaller --onefile --windowed GDR_Handler.py
pyinstaller --onefile --windowed --add-data 'src;src' GDR_Handler.py
pyinstaller --add-data 'src;src' -i ".\src\gdr.ico" --onefile --windowed GDR_Handler.py
```

- Version: 1.0.0
- License: MIT
- Author: Mero2online
