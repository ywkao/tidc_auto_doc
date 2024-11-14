from setuptools import setup, find_packages

setup(
    name="tidc_auto_doc",
    version="0.1",
    packages=find_packages(),
    py_modules=['gdrive_utils'],
    install_requires=[
        'google-colab;platform_system=="Linux"',
    ],
)
