from setuptools import setup, find_packages

setup(
    name='tablepyxl',
    version='0.3.5',
    description='Generate Excel documents from html tables',
    url='https://github.com/martsberger/tablepyxl',
    download_url='https://github.com/martsberger/tablepyxl/archive/0.3.5.tar.gz',
    author='Brad Martsberger',
    author_email='bmarts@procuredhealth.com',
    license='MIT',
    packages=find_packages(),
    install_requires=['openpyxl', 'beautifulsoup4', 'premailer', 'requests', 'lxml']
)
