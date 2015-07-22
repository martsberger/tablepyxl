from setuptools import setup, find_packages

setup(
    name='tablepyxl',
    version='0.1',
    description='Generate Excel documents from html tables',
    url='https://github.com/martsberger/tablepyxl',
    author='Brad Martsberger',
    author_email='bmarts@procuredhealth.com',
    license='MIT',
    packages=find_packages(),
    install_requires=['openpyxl', 'beautifulsoup4', 'premailer']
)