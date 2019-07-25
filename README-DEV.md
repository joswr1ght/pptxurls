# Developer Notes

## Linting

`flake8 --max-line-length=120 pptxurls.py`

## Building EXE

```
C:\temp>pip install python-pptx pyinstaller
C:\temp>copy z:\dev\pptxurls\pptxurls.py .
C:\temp>pyinstaller --onefile --upx-exclude "vcruntime140.dll" --hidden-import pptx pptxurls.py
```
