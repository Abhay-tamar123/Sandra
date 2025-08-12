import shutil
import os
import win32com

gen_py = win32com.__gen_path__
shutil.rmtree(gen_py, ignore_errors=True)
print(f"Deleted: {gen_py}")