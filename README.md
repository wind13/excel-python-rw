# Excel python read write

## Setup Env

### Dev Env

```
mkvirtualenv excel_python_run
workon excel_python_run
pip install -r dep_dev.txt
```

### Release

```
sh zip.sh
```

### Prod Env

```
sh run.sh
```

## Install new lib

```
pip freeze > requirements.txt
```

## References

- [openpyxl - A Python library to read/write Excel 2010 xlsx/xlsm files](https://openpyxl.readthedocs.io/en/stable/)
- [openpyxl 模块](https://www.cnblogs.com/programmer-tlh/p/10461353.html)
