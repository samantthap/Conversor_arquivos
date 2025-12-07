# dependencias.py
import os
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tqdm import tqdm
import pandas as pd

# Tentativas de importação 
try:
    import win32com.client
except:
    win32com = None

try:
    from pdf2docx import Converter as PDF2DOCX_Converter
except:
    PDF2DOCX_Converter = None

try:
    from docx import Document
except:
    Document = None

try:
    import tabula
except:
    tabula = None

try:
    import pdfplumber
except:
    pdfplumber = None
