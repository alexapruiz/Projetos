from setuptools import setup
import codecs
import os
import re

here = os.path.abspath(os.path.dirname(__file__))

def find_version(*file_paths):
    with codecs.open(os.path.join(here, *file_paths), 'r', 'latin1') as f:
        version_file = f.read()

    version_match = re.search(r"^__version__ = ['\"]([^'\"]*)['\"]",
                              version_file, re.M)
    if version_match:
        return version_match.group(1)
    raise RuntimeError("Unable to find version string.")

with codecs.open('README.rst', encoding='utf-8') as f:
    long_description = f.read()

with codecs.open('LICENSE', encoding='utf-8') as f:
    license = f.read()

setup(
    name='dataview',
    version=find_version('dataview', '__init__.py'),
    description='A module that allows you to create views of your sequences or its slices',
    long_description=long_description,
    url='https://github.com/damamaty/dataview',
    author='Dmitriy Kirichenko',
    author_email='damamaty@gmail.com',
    license='MIT',
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.1',
        'Programming Language :: Python :: 3.2',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
    ],
    keywords="list slice",
    packages = ['dataview'],
)