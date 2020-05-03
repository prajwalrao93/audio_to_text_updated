from cx_Freeze import setup, Executable
import os, sys

# Dependencies are automatically detected, but it might need
# fine tuning.

where = os.path.dirname(sys.executable)


#os.environ['TCL_LIBRARY'] = where+"\\tcl\\tcl8.6"
#os.environ['TK_LIBRARY'] = where+"\\tcl\\tk8.6"

#includefiles=["FFmpeg","venv/tcl/tcl8.6", "venv/tcl/tk8.6"]
includefiles=["Media"]

buildOptions = dict(packages = ['scipy'], excludes = [], include_files=includefiles)

base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('speech_to_text.py', base=base)
]

setup(name='Audio_to_mp3',
      version = '1.0',
      description = 'Converting Audio to txt',
      options = dict(build_exe = buildOptions),
      executables = executables)
