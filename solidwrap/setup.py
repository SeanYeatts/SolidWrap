from setuptools import setup, find_packages

setup(
    name='solidwrap',
    version='0.0.1',
    packages=find_packages(),
    install_requires=[
        'quickpathstr',
    ],
)

"""
How to package a module:

1. Make sure 'setup.py' is located at the top level project folder

2. Use the terminal to run (from the top level folder):

    python setup.py sdist

3. To install the module in another project:

    pip install "[FILEPATH]\solidwrap-0.0.1.t"
    
    where [FILEPATH] is the top level directory of the target project

"""
