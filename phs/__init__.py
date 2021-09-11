import numpy as np
import pandas as pd
import bottleneck as bn
import numexpr as ne
import openpyxl
import os
import sys
from os import listdir
from os.path import isfile, isdir, join, realpath, dirname
import win32com as win32
from win32com.client import Dispatch
import time
from datetime import datetime, timedelta, date
import requests
import json
import holidays
from typing import Union, Callable
import itertools
import pyodbc
import xlwings as xw
import openpyxl
import unidecode
import shutil
from PIL import Image
import csv

###############################################################################

import matplotlib as mpl ; mpl.use('Agg')
import matplotlib.pyplot as plt ; plt.ioff()
from matplotlib.ticker import MaxNLocator, FuncFormatter, ScalarFormatter, FormatStrFormatter
import matplotlib.ticker
from matplotlib.dates import DateFormatter
import matplotlib.transforms as transforms
from pylab import yticks

import seaborn as sns

###############################################################################

import scipy as sc
from scipy import stats
from scipy.stats import rankdata
from scipy.interpolate import interp1d

###############################################################################

from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA

###############################################################################

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions

from bs4 import BeautifulSoup
from urllib.request import urlopen
import lxml

###############################################################################

import multiprocessing