from PyQt5.QtWidgets import QApplication
from interface import Application
import sys
import tkinter
from tkinter import ttk, filedialog, scrolledtext
from ttkthemes import ThemedStyle
from datetime import datetime
from tkinter import *
from tkinter import filedialog, messagebox
from scraper.ZLME_scrap import ZLME_scrap
from scraper.EURX_scrap import EURX_scrap
import configparser
import os
from excel.excel_manager import write_to_excel
from tkcalendar import DateEntry
import subprocess
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import load_workbook



def main():
    fenetre = Application()
    fenetre.mainloop()

if __name__ == "__main__":
    main()
