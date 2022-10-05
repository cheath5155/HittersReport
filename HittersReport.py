import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog

##Create a report that takes in hitter data from a CSV file with muliple
##Trackman games and generated a PPTX and PDF containing charts and other 
##Data from the hitter


csv_file = pd.read_csv(filedialog.askopenfilename())

def damage_heat_map():
    


