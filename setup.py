from setuptools import setup, find_packages



# import time
# print('今天是本年第几周？ ', time.strftime("%W"))




GFICLEE_VERSION = '23.45.0'

setup(
    name='kiwigo',
    version=GFICLEE_VERSION,
    packages=find_packages(),
    include_package_data=True,
    entry_points={
        # "console_scripts": ['cfastproject = fastproject.main:main']
    },
    install_requires=[
        # "configparser", "eml_parser", "pandas", "docx", "PyMuPDF", "fitz"
    ],
    url='https://github.com/gh-hty/kiwigo_box.git',
    # license='GNU General Public License v3.0',
    author='tyhua',
    author_email='hua.x64@outlook.com',
    description='For daily work.'
)