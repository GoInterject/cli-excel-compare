
For initializing project use this:
```bash
conda create -n excelcompare python=3.12
conda activate excelcompare
pip install poetry
poetry install
```



## Comparsion Excel files
For comparsion files use this commands:
```bash
excelcompare file1.xlsx file2.xlsx
```

Manualy compare saved files:
```bash
git diff -- file1.txt file2.txt
```

```bash
code --diff file1.txt file2.txt
```

## Conversion of Excel files to JSON

### Initializing
```bash
npm install
```


For convertion Excel file use command:
```bash
excelcompare tojson "FolderPaths/ExceFile.xlsx" "OutputFolderPath"
```

For saving output file in current folder:
```bash
excelcompare tojson "FolderPaths/ExceFile.xlsx" "."
```
