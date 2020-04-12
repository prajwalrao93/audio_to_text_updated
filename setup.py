from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
buildOptions = dict(packages = [], excludes = [])

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('speech_to_text.py', base=base)
]

setup(name='Audio_to_mp3',
      version = '1.0',
      description = 'Converting Audio to txt',
      options = dict(build_exe = buildOptions),
      executables = executables)
