pyinstaller --clean --onefile --windowed ^
    --distpath="dist/" ^
    --icon="icon.ico" ^
    --name="PromiseDateUtility" ^
    --add-data="icon.ico";. ^
    app.py