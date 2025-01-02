# setup.py

from setuptools import setup, find_packages

setup(
    name="o365Service",
    version="1.0.2",
    packages=find_packages(include=['o365Service', 'o365Service.*']),
    description="Custom o365 service for downloading files from SharePoint drive",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    author="Jakub Urban",
    author_email="urbanjakubdev@gmail.com",
    url="https://github.com/UrbanJakubDev/o365service",
    install_requires=[
        "msal >= 1.30.0",
        "requests >= 2.32.3",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.10',
)
