from setuptools import setup, find_packages
import platform
import warnings

requires = ["BeautifulSoup4", "cssutils", 'requests', 'webcolors',
            'pygments', 'lxml', 'contexttimer', 'docopt']

if platform.system() != "Windows":
    warnings.warn("wordinserter currently only supports Windows for generating documents,"
                  " functionality will be impaired.")
else:
    requires.append('comtypes')

setup(
    name='wordinserter',
    version='1.0.0',
    packages=find_packages(),
    url='https://github.com/orf/wordinserter',
    license='MIT',
    author='Tom',
    author_email='tom@tomforb.es',
    description='Render HTML and Markdown to a specific portion of a word document',
    install_requires=requires,
    long_description="""\
Render HTML and Markdown into a word document using win32com.
Check out the github repo for more information and code samples.
""",
    package_data={'wordinserter': ['images/*']},
    entry_points={
        'console_scripts': [
            'wordinserter=wordinserter.cli:run'
        ]
    },
    classifiers=[
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6'
    ]
)
