import os
from src import Excel

# PATH_WATCH=os.path.join(os.getcwd(), 'watch_examples')
def main():
    pathWatch = input("Enter path: ")
    excel = Excel(pathWatch)
    excel.readFiles()

main()