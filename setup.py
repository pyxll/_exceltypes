from setuptools import setup, find_packages, Extension
import win32api
import win32com
import os

include_dirs = [
    "src",
    os.path.join(os.path.dirname(win32api.__file__), "include"),
    os.path.join(os.path.dirname(win32com.__file__), "include")
]

library_dirs = [
    os.path.join(os.path.dirname(win32api.__file__), "libs"),
    os.path.join(os.path.dirname(win32com.__file__), "libs")
]

_exceltypes = Extension("exceltypes._exceltypes",
                        sources=[
                            "src/exceltypes.cpp",
                            "src/PyIRTDUpdateEvent.cpp"
                        ],
                        include_dirs=include_dirs,
                        library_dirs=library_dirs)

if __name__ == "__main__":
    setup(name="exceltypes",
            version=0.1,
            packages=find_packages(),
            ext_modules=[_exceltypes],
            use_2to3=True)
