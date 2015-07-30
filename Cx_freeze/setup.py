from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["os", "datetime", "fuzzywuzzy", "xlrd", "time", "csv", "shutil", "Levenshtein"]}

setup(	name='Pipe Converter',
    	version = '4.2',
    	description='converter of files, with pipes.',
    	options = {"build_exe": build_exe_options},
    	executables = [Executable("Pipe Converter.py", base = None)])