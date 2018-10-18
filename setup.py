from setuptools import setup, find_packages

setup(
    name='tablepyxl',
    version='0.5',
    description='Generate Excel documents from html tables',
    url='https://github.com/martsberger/tablepyxl',
    download_url='https://github.com/martsberger/tablepyxl/archive/0.5.tar.gz',
    author='Brad Martsberger, Asma Mehjabeen, Brian Davis',
    author_email='bmarts@procuredhealth.com',
    license='MIT',
    classifiers=[
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6'
    ],
    packages=find_packages(),
    install_requires=['openpyxl', 'premailer', 'requests', 'lxml']
)
