import os
import re
import setuptools
from distutils.core import setup

with open("README.md") as f:
    long_description = f.read()

def read_version():
    with open(os.path.join('visma_administration', '__init__.py'), encoding='utf8') as f:
        m = re.search(r'''__version__\s*=\s*['"]([^'"]*)['"]''', f.read())
        if m:
            return m.group(1)
        raise ValueError("couldn't find version")

version = read_version()
setup(
    name="visma_administration",
    packages=["visma_administration"],
    version=version,
    license="MIT",
    description="API for Visma Administration 200/500/1000/2000",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Viktor Johansson",
    author_email="dpedesigns@hotmail.com",
    url="https://github.com/viktor2097/visma-administration",
    download_url=f"https://github.com/viktor2097/visma-administration/archive/{version}.tar.gz",
    keywords=["visma", "visma administration"],
    install_requires=["pythonnet>=3.0.3"],
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
    ],
)
