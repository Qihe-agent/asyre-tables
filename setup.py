from setuptools import setup

setup(
    name='asyre-tables',
    version='0.1.0',
    py_modules=['office'],
    install_requires=[
        'openpyxl>=3.1.0',
    ],
    extras_require={
        'xls': ['xlrd>=2.0.0'],
    },
    entry_points={
        'console_scripts': [
            'tables=office:main',
        ],
    },
    python_requires='>=3.8',
    author='Yixuan Zhang',
    description='Spreadsheet operations CLI for AI agents',
    license='MIT',
)
