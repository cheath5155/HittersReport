import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
import seaborn as sns

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
    player_df = player_df.drop(player_df[player_df.PitchCall == ('BallCalled' or 'StrikeCalled' or 'BallInDirt')].index)
    #player_df.drop(player_df.columns[player_df.apply(lambda col: col)], axis=1)
    remove_list = player_df.columns.values.tolist()
    remove_list.remove("Batter")
    remove_list.remove("PlateLocHeight")
    remove_list.remove("PlateLocSide")
    remove_list.remove("PitchCall")
    player_df = player_df.drop(remove_list, axis=1)
    player_df['PlateLocSide'] = (player_df['PlateLocSide'] * -1)
    print(player_df)


    return player_df

def swing2d_density_plot(player_df):
    sns.set_style("white")
    sns.kdeplot(x=player_df.PlateLocSide, y=player_df.PlateLocHeight,cmap="vlag", shade=True, bw_adjust=.5)
    plt.show()
    return

swing2d_density_plot(csv_to_swing_df())