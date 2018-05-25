import os
import platform
import warnings

from setuptools import find_packages, setup

requires = ["BeautifulSoup4", "cssutils", 'requests', 'webcolors',
            'pygments', 'lxml', 'contexttimer', 'docopt']

if platform.system() != "Windows":
    warnings.warn("wordinserter currently only supports Windows for generating documents,"
                  " functionality will be impaired.")
else:
    requires.append('comtypes')


readme = ""
if os.path.exists('README.rst'):
    with open('README.rst') as f:
        readme = f.read()

setup(
    name='wordinserter',
    version='1.1.0',
    packages=find_packages(),
    url='https://github.com/orf/wordinserter',
    license='MIT',
    author='Tom',
    author_email='tom@tomforb.es',
    description='Render HTML and Markdown to a specific portion of a word document',
    install_requires=requires,
    long_description=readme,
    package_data={'wordinserter': ['images/*']},
    entry_points={
        'console_scripts': [
            'wordinserter=wordinserter.cli:run'
        ]
    },
    classifiers=[
        'Programming Language :: Python :: 3 :: Only',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6'
    ]
)
