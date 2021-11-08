from os import getenv
from os.path import join
from shutil import rmtree


def clear_temp_directory(path, win_dir: str = None):
    if win_dir:
        path = join(getenv(win_dir), path)
    rmtree(path)
