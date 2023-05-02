from tkinter import *
from tkinter import filedialog
import os
import platform
from docparser import sendFilesToDocParser

# find the path of the current file, working for both windows and mac

if platform.system() == 'Darwin':
    path = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
# for windows :
elif platform.system() == '':
    path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')


def main():
    filepath = filedialog.askopenfilenames(
        initialdir=path, title="Convert to excel", filetypes=(("Pdf Files", "*.pdf"),))
    if filepath:
        print(filepath)
        sendFilesToDocParser(filepath)


if __name__ == "__main__":
    main()
