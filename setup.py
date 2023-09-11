from setuptools import setup, find_packages



# import time
# print('今天是本年第几周？ ', time.strftime("%W"))




GFICLEE_VERSION = 'v23.37.0'

setup(
    name='kiwigo',
    version=GFICLEE_VERSION,
    packages=find_packages(),
    include_package_data=True,
    entry_points={
        # "console_scripts": ['cfastproject = fastproject.main:main']
    },
    install_requires=[
        "configparser", "inspect",
        "eml_parser", "base64", "re", "pandas", "docx", "comtypes", "fitz"
    ],
    # url='https://github.com/',
    # license='GNU General Public License v3.0',
    author='tyhua',
    author_email='hua.x64@outlook.com',
    description='For daily work of cob.'
)