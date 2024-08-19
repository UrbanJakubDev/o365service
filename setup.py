# setup.py

from setuptools import setup, find_packages

setup(
    name="o365Service",
    version="1.0.0",
    packages=find_packages(include=['o365', 'o365.*']),
    description="custom o365 service for download files from sharepoint drive",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    author="Jakub Urban",
    author_email="urbanjakubdev@gmail.com",
    url="https://github.com/UrbanJakubDev/o365service",
    install_requires=[
        # Add your dependencies here
        # Example: "numpy >= 1.18.0"
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
