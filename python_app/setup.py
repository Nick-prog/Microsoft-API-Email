#!/usr/bin/env python3
"""
Setup script for Microsoft Graph API Explorer - Dynamic Filters
"""

from setuptools import setup, find_packages
import os

# Read the README file for long description
def read_readme():
    readme_path = os.path.join(os.path.dirname(__file__), 'README.md')
    if os.path.exists(readme_path):
        with open(readme_path, 'r', encoding='utf-8') as f:
            return f.read()
    return "Microsoft Graph API Explorer with dynamic filter builders"

# Read requirements from requirements.txt
def read_requirements():
    req_path = os.path.join(os.path.dirname(__file__), 'requirements.txt')
    if os.path.exists(req_path):
        with open(req_path, 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip() and not line.startswith('#')]
    return []

setup(
    name="msgraph-api-explorer",
    version="1.0.0",
    author="Microsoft Graph API Explorer Team",
    author_email="developer@company.com",
    description="A Python GUI application for exploring Microsoft Graph APIs with dynamic filter builders",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/your-org/msgraph-api-explorer",
    packages=find_packages(),
    py_modules=['main_dynamic_filters'],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Internet :: WWW/HTTP :: Dynamic Content",
        "Topic :: Office/Business",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: OS Independent",
        "Environment :: X11 Applications :: Qt",
        "Environment :: Win32 (MS Windows)",
        "Environment :: MacOS X",
    ],
    python_requires=">=3.8",
    install_requires=read_requirements(),
    extras_require={
        "dev": [
            "pytest>=7.0",
            "pytest-cov>=4.0",
            "black>=22.0",
            "flake8>=5.0",
            "mypy>=1.0"
        ],
        "msal": [
            "msal>=1.20.0",
            "requests>=2.28.0",
            "beautifulsoup4>=4.11.0"
        ]
    },
    entry_points={
        "console_scripts": [
            "msgraph-explorer=main_dynamic_filters:main",
        ],
        "gui_scripts": [
            "msgraph-explorer-gui=main_dynamic_filters:main",
        ]
    },
    include_package_data=True,
    package_data={
        "": ["*.md", "*.txt", "*.cfg", "*.ini"],
    },
    keywords=[
        "microsoft", "graph", "api", "msal", "oauth", "azure", "office365", 
        "mail", "calendar", "gui", "tkinter", "explorer", "filter", "dynamic"
    ],
    project_urls={
        "Bug Reports": "https://github.com/your-org/msgraph-api-explorer/issues",
        "Source": "https://github.com/your-org/msgraph-api-explorer",
        "Documentation": "https://github.com/your-org/msgraph-api-explorer/wiki",
    },
    zip_safe=False,
)
