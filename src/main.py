from PyQt5.QtWidgets import QApplication
from interface import Application
import sys



def main():
    fenetre = Application()
    sys.exit(fenetre.exec_())

if __name__ == "__main__":
    main()
