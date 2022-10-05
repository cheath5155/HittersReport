import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog

##Create a report that takes in hitter data from a CSV file with muliple
##Trackman games and generated a PPTX and PDF containing charts and other 
##Data from the hitter


csv_file = filedialog.askopenfilename()
csv_df = pd.read_csv(csv_file)
names = ['Bazzana, Travis']
def csv_to_swing_df():
    global csv_df
    player_df = csv_df
    player_df = player_df.drop(player_df[player_df.Batter != names[0]].index)
    #player_df.drop(player_df.columns[player_df.apply(lambda col: col)], axis=1)
    player_df.columns.values.tolist(remove_list[])
    print(remove_list)

    return



csv_to_swing_df()