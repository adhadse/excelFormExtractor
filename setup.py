import setuptools
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
    packages=setuptools.find_packages(include=["py_excel_form_extractor"]),
    license="MIT",
    platforms="Linux, Mac OS X",
    keywords=["go", "golang", "python", "excel", "xlsx", "form", "extractor"],
    py_modules = ["py_excel_form_extractor.extractor"],
    package_data={"py_excel_form_extractor": ["*.so"]},
    include_package_data=True,
)
