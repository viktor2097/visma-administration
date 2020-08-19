from distutils.core import setup

# read the contents of your README file
from os import path
this_directory = path.abspath(path.dirname(__file__))
with open(path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name="visma_administration",
    packages=["visma_administration"],
    version="0.1.3",
    license="MIT",
    description="API for Visma Administration 200/500/1000/2000",
    long_description=long_description,
    long_description_content_type='text/markdown'
    author="Viktor Johansson",
    author_email="dpedesigns@hotmail.com",
    url="https://github.com/viktor2097/visma-administration",
    download_url="https://github.com/viktor2097/visma-administration/archive/0.1.3.tar.gz",
    keywords=["visma"],
    install_requires=["pythonnet>=2.5.1"],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
    ],
)

