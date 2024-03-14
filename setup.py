from setuptools import setup, find_packages
from setuptools.command.build_py import build_py
import os


class CustomBuildCommand(build_py):
    def run(self):
        super().run()
        os.system("dotnet build py_fluid_vb")


with open("requirements.txt") as f:
    requirements = f.read().splitlines()

setup(
    name="py-fluid-vb",
    version="0.1.0",
    packages=find_packages(),
    include_package_data=True,
    package_data={"py_fluid_vb": ["vb_funcs_lookup.json"]},
    author="Josepgh McGovern",
    author_email="joseph.mcgovern16@gmail.com",
    description="py_fluid_vb, provides an interface to access fluid property calculations from a VB.NET library.",
    install_requires=requirements,
)
