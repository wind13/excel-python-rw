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