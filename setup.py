from setuptools import setup, find_packages

setup(
    name='wordinserter',
    version='0.6',
    packages=find_packages(),
    url='https://github.com/orf/wordinserter',
    license='MIT',

    author='Tom',
    author_email='tom@tomforb.es',
    description='Render HTML and Markdown to a specific portion of a word document',
    install_requires=["BeautifulSoup4", "cssutils", 'CommonMark', 'requests', 'webcolors', 'pypiwin32'],
    include_package_data=True,
    long_description="""\
Render HTML and Markdown into a word document using win32com.
Check out the github repo for more information and code samples.
""",
    entry_points={
        'console_scripts': {
            'wordrender=wordinserter.command:main'
        }
    }
)
