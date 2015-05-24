from setuptools import setup, find_packages

setup(
    name='wordinserter',
    version='0.5',
    packages=find_packages(),
    url='https://github.com/orf/HtmlToWord',
    license='',

    author='Tom',
    author_email='tom@tomforb.es',
    description='Render HTML to a specific portion of a word document',
    install_requires=["BeautifulSoup4", "cssutils", 'CommonMark'],
    include_package_data=True,
    long_description="""\
Render HTML to a word document using win32com.
Check out the github repo for more information and code samples.
"""
)
