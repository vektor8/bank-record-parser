pyinstaller --noconfirm --onefile --windowed --name cec_parser `
  --add-data "data;data" `
  --add-data "core;core" `
  --add-data "forest-theme;forest-theme" `
  main.py