from setuptools import setup, find_packages
import platform
import warnings

requires = ["BeautifulSoup4", "cssutils", 'requests', 'webcolors', 'pygments']

if platform.system() == "Windows":
    requires.append("pypiwin32")
else:
    warnings.warn("wordinserter currently only supports Windows for generating documents,"
                  " functionality will be impaired.")


setup(
    name='wordinserter',
    version='0.6.12',
    packages=find_packages(),
    url='https://github.com/orf/wordinserter',
    license='MIT',
    author='Tom',
    author_email='tom@tomforb.es',
    description='Render HTML and Markdown to a specific portion of a word document',
    install_requires=requires,
    include_package_data=True,
    long_description="""\
Render HTML and Markdown into a word document using win32com.
Check out the github repo for more information and code samples.
"""
)
