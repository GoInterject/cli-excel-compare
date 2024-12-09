
For initializing project use this:
```bash
conda create -n excelcompare python=3.12
conda activate excelcompare
pip install poetry
poetry install
```




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

