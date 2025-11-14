pyinstaller --noconfirm --onefile --windowed --name cec_parser `
  --add-data "rules.csv;." `
  --add-data "lib;lib" `
  --add-data "parsers;parsers" `
  --hidden-import=parsers.cec_parsers `
  --hidden-import=tika `
  --hidden-import=PyPDF2 `
  main.py