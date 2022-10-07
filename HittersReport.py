from unicodedata import name
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
import seaborn as sns
import os

##Create a report that takes in hitter data from a CSV file with muliple
##Trackman games and generated a PPTX and PDF containing charts and other 
##Data from the hitter

#Asks for CSV File
#csv_file = filedialog.askopenfilename()
csv_file = "C:\\Users\\cmhea\\OneDrive\\Documents\\baseball\\2022-23 OSU CSVs\\Combine CSVs\\oct6combined.csv"
csv_df = pd.read_csv(csv_file)
names = ['Bazzana, Travis']

#Methods Job is to Clean Up the csv data frame to one only containg rows 
#of player we want and collumns we need for the swing density chart
def csv_to_swing_df():
    global csv_df
    player_df = csv_df

    #Drops all rows where player isn't hitting
    player_df = player_df.drop(player_df[player_df.Batter != names[0]].index)
    #Drops all rows where player didn't swing
    player_df = player_df.drop(player_df[player_df.PitchCall == 'BallCalled'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'StrikeCalled'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'BallinDirt'].index)
    #Creating a list of collums to remove from DF and then removing collums we need from list
    remove_list = player_df.columns.values.tolist()
    remove_list.remove("Batter")
    remove_list.remove("PlateLocHeight")
    remove_list.remove("PlateLocSide")
    remove_list.remove("PitchCall")
    remove_list.remove("BatterSide")
    #Removes all collums other than those with .remove above from data frame
    player_df = player_df.drop(remove_list, axis=1)
    #Switches plate loc side values to be in catcher view by multiplying by -1
    player_df['PlateLocSide'] = (player_df['PlateLocSide'] * -1)
    print(player_df)


    return player_df

def swing2d_density_plot(player_df):
    #Pulls image for background will have to imput if statemnt to deptermine right vs left
    img = plt.imread("LHH.png")
    fig, ax = plt.subplots(figsize=(6, 6))
    sns.set_style("white")
    #Creates density plot, camp is color scheme and alpha is transperacy
    sns.kdeplot(x=player_df.PlateLocSide, y=player_df.PlateLocHeight,cmap="YlOrBr", shade=True, bw_adjust=.5, ax=ax, alpha = 0.7)
    #creates demenstions for graph plus displays image
    ax.imshow(img, extent=[-2.75,2.75,-0.6,5.3], aspect=1)
    #creates path for plot to be saved in there isn't one
    newpath = os.path.join("Sheets", names[0])
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    #saves plot in folder
    plt.savefig(os.path.join("Sheets", names[0], 'swingchart.png'))

    return

def find_table_metrics():
    global csv_df
    player_df = csv_df
    #Drops all rows where player isn't hitting
    player_df = player_df.drop(player_df[player_df.Batter != names[0]].index)
    remove_list = player_df.columns.values.tolist()
    remove_list.remove("Batter")
    remove_list.remove("PlateLocHeight")
    remove_list.remove("PlateLocSide")
    remove_list.remove("PitchCall")
    remove_list.remove("BatterSide")
    remove_list.remove("KorBB")
    remove_list.remove("PlayResult")
    remove_list.remove("ExitSpeed")
    #Removes all collums other than those with .remove above from data frame
    player_df = player_df.drop(remove_list, axis=1)
    avg_ev_df = player_df.drop(player_df[player_df.ExitSpeed < 60.0].index)
    avg_ev = round(avg_ev_df["ExitSpeed"].mean(),1)
    max_ev = round(avg_ev_df["ExitSpeed"].max(),1)

    swings = (player_df["PitchCall"].value_counts("InPlay") + player_df["PitchCall"].value_counts("FoulBall") + player_df["PitchCall"].value_counts("StrikeSwinging"))
    takes = (player_df["PitchCall"].value_counts("BallCalled") + player_df["PitchCall"].value_counts("StrikeCalled"))
    swing_rate = round(100*(swings/(swings+takes)),1)



    return

def data_frame_for_damage_chart():

    return
def damage_chart():
    return

swing2d_density_plot(csv_to_swing_df())
#find_table_metrics()