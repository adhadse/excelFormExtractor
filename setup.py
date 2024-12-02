import setuptools
from setuptools import Extension
import os

version = os.environ.get('PACKAGE_VERSION', None)
if version is None:
    raise ValueError(f"version {version} is None. ENV var PACKAGE_VERSION: {os.environ.get('PACKAGE_VERSION')}")
version = version.lstrip('v')

setuptools.setup(
    name="py_excel_form_extractor",
    version=version,
    url="https://github.com/adhadse/excelFormExtractor",
    author="Anurag Dhadse",
    description="Extract excel form content into structured data.",
    long_description_content_type="text/markdown",
    packages=setuptools.find_packages(include=["py_excel_form_extractor"]),
    license="MIT",
    platforms="Linux",
    keywords=["go", "golang", "python", "excel", "xlsx", "form", "extractor"],
    ext_modules=[
        Extension(
            "py_excel_form_extractor._extractor",
            sources=["py_excel_form_extractor/extractor.c"],
            include_dirs=["py_excel_form_extractor"],
        )
    ],
    py_modules = ["py_excel_form_extractor.extractor", "py_excel_form_extractor.utils"],
    package_data={"py_excel_form_extractor": [
        "*.so",
        "*_go.py",
        "*.py",
        "_*.py",
        "*.h",
        "*.c",
    ]},
    include_package_data=True,
)
