# Excel extractor
Extract excel form content into structured data.

## Usage

1. SECCF extraction: supplier export control classification Form/declaration

```python
from py_excel_form_extractor import extractor, go

company_names = extractor.CompanyNameList()  # the company name which can be mentioned in the file
for company_name in ["Boeing", "Boeing Ltd"]:
    company_names.append(company_name)

extr = extractor.MakeSECCFExtractor("Example.xlsx", company_names)
extraction = extr.Extract()

# convert to JSON string
extr_json = extr.ToJson()
```

## BUILD

1. Building the go binary
```bash
go build -o gobinary ./cmd/excelExtractor
```

2. Running the program without building the binary
```bash
go build -o ./bin/excel-extrator ./cmd/excelFormExtractor/main.go
```
3. Run the binary:
```bash
./bin/excel-extrator
```

## Local Python bindings generation and installation

```bash
pip3 install pybindgen wheel
gopy build --output=py_excel_form_extractor -vm=python3 ./pkg/*
RELEASE_VERSION=YOUR_UPDATED_PACKAGE_VERSION python3 setup.py bdist_wheel --force

# install wheel file
wheel_file=$(ls dist/*.whl | head -n1); pip3 install $wheel_file
```
