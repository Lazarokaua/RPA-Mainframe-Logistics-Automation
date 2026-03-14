
Write-Host "Instalando PyInstaller..."
pip install pyinstaller

Write-Host "Limpando builds anteriores..."
Remove-Item -Path "dist" -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path "build" -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path "*.spec" -ErrorAction SilentlyContinue


Write-Host "Localizando bibliotecas..."
$ctk_path = python -c "import customtkinter; import os; print(os.path.dirname(customtkinter.__file__))"
$ctk_msg_path = python -c "import CTkMessagebox; import os; print(os.path.dirname(CTkMessagebox.__file__))"

Write-Host "CustomTkinter encontrado em: $ctk_path"
Write-Host "CTkMessagebox encontrado em: $ctk_msg_path"


Write-Host "Gerando o executável..."
pyinstaller --noconfirm --clean --windowed --onedir --name "RPA - Calculo de Cubagem" --paths "src" `
    --add-data "output/base_filiais.xlsx;output" `
    --add-data "$ctk_path;customtkinter" `
    --add-data "$ctk_msg_path;CTkMessagebox" `
    src/main.py

Write-Host "Copiando arquivos de configuração..."


New-Item -Path "dist/output" -ItemType Directory -Force


Copy-Item -Path "output/base_filiais.xlsx" -Destination "dist/output/base_filiais.xlsx"

Write-Host "Build Concluído!"
Write-Host "O executável e a pasta output estão em: dist/"
