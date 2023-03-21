from ast import walk
from cgi import print_directory
from email.mime import image
import math
from cmath import nan
from unicodedata import name
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
import os

from pptx import Presentation
from pptx.util import Pt
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from scipy.interpolate import interp2d
from scipy.interpolate import RectBivariateSpline
from scipy.ndimage.filters import gaussian_filter
from scipy import interpolate
import matplotlib.patches as patches
from PIL import Image
from pptx.dml.color import RGBColor
from tkinter import filedialog

##Create a report that takes in hitter data from a CSV file with muliple
##Trackman games and generated a PPTX and PDF containing charts and other 
##Data from the hitter

#Asks for CSV File
#csv_file = filedialog.askopenfilename()
csv_file = filedialog.askopenfilename()
csv_df = pd.read_csv(csv_file)
names = ['Reeder, Canon']

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
    player_df = player_df.drop(player_df[player_df.PitchCall == 'HitByPitch'].index)
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

# Currently Unused but can replace the swing df to make a whiff location chart
def csv_to_whiff_df():
    global csv_df
    player_df = csv_df

    #Drops all rows where player isn't hitting
    player_df = player_df.drop(player_df[player_df.Batter != names[0]].index)
    #Drops all rows where player didn't whiff
    player_df = player_df.drop(player_df[player_df.PitchCall == 'BallCalled'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'StrikeCalled'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'BallinDirt'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'FoulBall'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'InPlay'].index)
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

#Takes in a dataframe and creates a seaborn 2d density plot based on pitch location of dataframe
def swing2d_density_plot(player_df):
    #Pulls image for background will have to imput if statemnt to deptermine right vs left
    rhhs = ['Burke, Isaiah','Cedillo, Ruben']
    img = plt.imread("RHH.png")
    fig, ax = plt.subplots(figsize=(6, 6))
    sns.set_style("white")
    #Creates density plot, camp is color scheme and alpha is transperacy
    chart = sns.kdeplot(x=player_df.PlateLocSide, y=player_df.PlateLocHeight,cmap='rocket_r', shade=True, bw_adjust=.55, ax=ax, alpha = 0.65)
    #creates demenstions for graph plus displays image
    ax.imshow(img, extent=[-2.63,2.665,-0.35,5.30], aspect=1)

    chart.set(xticklabels=[])  
    chart.set(xlabel=None)
    chart.tick_params(bottom=False)
    chart.set(yticklabels=[])  
    chart.set(ylabel=None)
    chart.tick_params(left=False)   # remove the ticks
    #rect = patches.Rectangle((-0.708333, 1.6466667), 1.4166667, 1.90416667, linewidth=1, edgecolor='black', facecolor='none')
    #ax.add_patch(rect)
    #creates path for plot to be saved in there isn't one
    newpath = os.path.join("Sheets", names[0])
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    #saves plot in folder
    plt.axis('off')
    plt.savefig(os.path.join("Sheets", names[0], 'swingchart.png'),bbox_inches='tight', pad_inches = 0)


    return

#Gets Metrics for both tables
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
    remove_list.remove("TaggedHitType")
    remove_list.remove("RunsScored")

    #Removes all collums other than those with .remove above from data frame
    player_df = player_df.drop(remove_list, axis=1)
    avg_ev_df = player_df.drop(player_df[player_df.ExitSpeed < 60.0].index)
    avg_ev = "%.1f" % round(avg_ev_df["ExitSpeed"].mean(),1)
    max_ev = "%.1f" % round(avg_ev_df["ExitSpeed"].max(),1)
    
    #Swing %
    swings = (player_df["PitchCall"] == "InPlay").sum() + (player_df["PitchCall"] == "FoulBall").sum() + (player_df["PitchCall"] == "StrikeSwinging").sum()
    takes = (player_df["PitchCall"] == "BallCalled").sum() + (player_df["PitchCall"] == "StrikeCalled").sum()
    swing_rate = "%.1f" % round(100*(swings/(swings+takes)),1)

    #Chase Rate
    out_of_zone_df = player_df
    indexzone  = out_of_zone_df[ (out_of_zone_df['PlateLocSide'] < 0.83083) & (out_of_zone_df['PlateLocSide'] > -0.83083) & (out_of_zone_df['PlateLocHeight'] < 3.67333) & (out_of_zone_df['PlateLocHeight'] > 1.52417)].index
    out_of_zone_df = out_of_zone_df.drop(indexzone)
    chases = (out_of_zone_df["PitchCall"] == "InPlay").sum() + (out_of_zone_df["PitchCall"] == "FoulBall").sum() + (out_of_zone_df["PitchCall"] == "StrikeSwinging").sum()
    takes_out_of_zone =  (out_of_zone_df["PitchCall"] == "BallCalled").sum() + (out_of_zone_df["PitchCall"] == "StrikeCalled").sum()
    chase_rate = "%.1f" % round(100*(chases/(chases+takes_out_of_zone)),1)

    #K Rate
    strikeouts = (player_df["KorBB"] == "Strikeout").sum()
    plate_apearences = (player_df["KorBB"] == "Walk").sum() + (player_df["KorBB"] == "Strikeout").sum() + (player_df["PitchCall"] == "InPlay").sum() + (player_df["PitchCall"] == "HitByPitch").sum()
    k_rate = "%.1f" % round(100*(strikeouts/plate_apearences),1)

    #(BB+HBP)/K
    walks = (player_df["KorBB"] == "Walk").sum()
    hbps = (player_df["PitchCall"] == "HitByPitch").sum()
    bb_hbp_over_ks = "%.2f" % round(((walks + hbps)/strikeouts),2)

    #BABIP
    bip = (player_df["PitchCall"] == "InPlay").sum() - (player_df["PlayResult"] == "HomeRun").sum()
    hits_no_hrs = (player_df["PlayResult"] == "Single").sum() + (player_df["PlayResult"] == "Double").sum() + (player_df["PlayResult"] == "Triple").sum()
    babip = "%.3f" % round(hits_no_hrs/bip,3)

    #AVG
    ABs = (player_df["PitchCall"] == "InPlay").sum() - (player_df["PlayResult"] == "Sacrifice").sum() + (player_df["KorBB"] == "Strikeout").sum()
    hits = (player_df["PlayResult"] == "Single").sum() + (player_df["PlayResult"] == "Double").sum() + (player_df["PlayResult"] == "Triple").sum() + (player_df["PlayResult"] == "HomeRun").sum()
    avg = "%.3f" % round(hits/ABs,3)

    #RBIs
    inplay_df = player_df.drop(player_df[player_df.PitchCall != 'InPlay'].index)
    walk_df = player_df.drop(player_df[player_df.KorBB != 'Walk'].index)
    hbp_df = player_df.drop(player_df[player_df.PitchCall != 'HitByPitch'].index)
    rbidf = pd.concat([inplay_df,walk_df,hbp_df])
    rbi = "%.0f" % round(rbidf["RunsScored"].sum(),0)

    #Doubles, Triples, Home Runs
    doubles = (player_df["PlayResult"] == "Double").sum()
    triples = (player_df["PlayResult"] == "Triple").sum()
    homers = (player_df["PlayResult"] == "HomeRun").sum()
    
    #OBP
    obp = "%.3f" % round((walks + hbps + hits)/plate_apearences,3)

    #SLG
    slg = "%.3f" % round((hits + doubles + triples*2 + homers*3)/ABs,3)

    #OPS
    ops = "%.3f" % (round((walks + hbps + hits)/plate_apearences,3) + round((hits + doubles + triples*2 + homers*3)/ABs,3))

    #WOBA
    woba = "%.3f" % (round((0.689*walks + 0.720*hbps + 0.844*(hits-doubles-triples-homers) + 1.261*doubles + 1.601*triples + 2.072*homers)/plate_apearences,3))
    league_woba = 0.351


    #Compiles Data into one list to pass to the presentation fucntion
    data_to_pass_to_presentation = []
    data_to_pass_to_presentation.append(str(avg_ev))
    data_to_pass_to_presentation.append(str(max_ev))
    data_to_pass_to_presentation.append(str(swing_rate) + "%")
    data_to_pass_to_presentation.append(str(chase_rate) + "%")
    data_to_pass_to_presentation.append(str(k_rate) + "%")
    data_to_pass_to_presentation.append(str(bb_hbp_over_ks))
    data_to_pass_to_presentation.append(str(babip).lstrip('0'))
    data_to_pass_to_presentation.append(str(woba).lstrip('0'))
    data_to_pass_to_presentation.append(str(plate_apearences))
    data_to_pass_to_presentation.append(str(ABs))
    data_to_pass_to_presentation.append(str(hits))
    data_to_pass_to_presentation.append(str(doubles))
    data_to_pass_to_presentation.append(str(triples))
    data_to_pass_to_presentation.append(str(homers))
    data_to_pass_to_presentation.append(str(rbi))
    data_to_pass_to_presentation.append(str(walks))
    data_to_pass_to_presentation.append(str(hbps))
    data_to_pass_to_presentation.append(str(strikeouts))
    data_to_pass_to_presentation.append(str(avg).lstrip('0'))
    data_to_pass_to_presentation.append(str(obp).lstrip('0'))
    data_to_pass_to_presentation.append(str(slg).lstrip('0'))
    data_to_pass_to_presentation.append(str(ops).lstrip('0'))


    print(data_to_pass_to_presentation)
    return data_to_pass_to_presentation

#Creates data Frame for EV Heat map Catcher View
def data_frame_for_damage_chart():
    global csv_df
    damage_df = csv_df

    #Drops rows without listed hitter
    damage_df = damage_df.drop(damage_df[damage_df.Batter != names[0]].index)

    #Remove unnesassary columns
    remove_list = damage_df.columns.values.tolist()
    remove_list.remove("Batter")
    remove_list.remove("PlateLocHeight")
    remove_list.remove("PlateLocSide")
    remove_list.remove("BatterSide")
    remove_list.remove("ExitSpeed")
    damage_df = damage_df.drop(remove_list, axis=1)

    #Removes Rows where EV is nan
    drop_index = []
    for index, row in damage_df.iterrows():
        if math.isnan(row['ExitSpeed']) == True:
            drop_index.append(index)
    damage_df = damage_df.drop(drop_index)

    #Removes Rows where EV < 60
    damage_df = damage_df.drop(damage_df[damage_df.ExitSpeed < 60.0].index)

    #Switches from Pitcher view to catcher view
    damage_df['PlateLocSide'] = (damage_df['PlateLocSide'] * -1)

    return damage_df

#Creates data Frame for EV Heat map Overhead View
def data_frame_for_overhead_damage_chart():
    global csv_df
    damage_df = csv_df
    damage_df = damage_df.drop(damage_df[damage_df.Batter != names[0]].index)

    #Removes Unnessasry Collumns
    remove_list = damage_df.columns.values.tolist()
    remove_list.remove("Batter")
    #POS Z is essentally -platelocside(already in catchers view)
    remove_list.remove("ContactPositionZ")
    #POS X is distance from front of home plate where contact occured in line with pitcher
    remove_list.remove("ContactPositionX")
    remove_list.remove("BatterSide")
    remove_list.remove("ExitSpeed")
    damage_df = damage_df.drop(remove_list, axis=1)

    #Removes Rows where EV or Contact Position is nan
    drop_index = []
    for index, row in damage_df.iterrows():
        if math.isnan(row['ExitSpeed']) == True:
            drop_index.append(index)
        if math.isnan(row['ContactPositionX']) == True:
            drop_index.append(index)
        if math.isnan(row['ContactPositionZ']) == True:
            drop_index.append(index)
    damage_df = damage_df.drop(drop_index)

    #Removes Rows where EV is under 60
    damage_df = damage_df.drop(damage_df[damage_df.ExitSpeed < 60.0].index)

    return damage_df

#Creates an overhead damage chart with matplotlib and saves the file in sheets
def damage_chart_overhead(damage_df):
    print(damage_df)
    x = []
    y = []
    array = []
    print(damage_df)
    for i in range(17):
        row = []
        for j in range(8):
            temp_df = damage_df
            top_limit = 4.50-0.375*i
            bottom_limit = 4.50-0.375*(i+1)
            left_limit = -1.500 + 0.375*j
            right_limit = -1.500 + 0.375*(j+1)
            if top_limit == 0:
                print(i)
            #drops data from df not in cell needed
            too_left = temp_df[(temp_df['ContactPositionZ'] < left_limit)].index
            temp_df = temp_df.drop(too_left)
            too_right = temp_df[(temp_df['ContactPositionZ'] > right_limit)].index
            temp_df = temp_df.drop(too_right)
            too_high = temp_df[(temp_df['ContactPositionX'] > top_limit)].index
            temp_df = temp_df.drop(too_high)
            too_low = temp_df[(temp_df['ContactPositionX'] < bottom_limit)].index
            temp_df = temp_df.drop(too_low)

            avg_ev_for_zone = round(temp_df["ExitSpeed"].mean(),1)
            
            #if math.isnan(avg_ev_for_zone) == True:
                #avg_ev_for_zone = 60
                

            row.append(avg_ev_for_zone)

        array.append(row)
    df = pd.DataFrame(array)
    print(df)
    df.to_numpy()
    x = np.arange(0, df.shape[1])
    y = np.arange(0, df.shape[0])
    #mask invalid values
    df = np.ma.masked_invalid(df)
    xx, yy = np.meshgrid(x, y)
    #get only the valid values
    x1 = xx[~df.mask]
    y1 = yy[~df.mask]
    newarr = df[~df.mask]

    GD1 = interpolate.griddata((x1, y1), newarr.ravel(), (xx, yy), method='linear', fill_value=60.0)

    df = pd.DataFrame(GD1)

    print(df)

    fig, ax = plt.subplots()
    fig = plt.imshow(df, cmap = 'jet',vmin=60,vmax=100, interpolation='bicubic')
    plt.plot([1.611, 5.389], [11.5, 11.5], color='black', linestyle='-', linewidth=2)
    plt.plot([1.611, 1.611], [11.5, 13.3888889], color='black', linestyle='-', linewidth=2)
    plt.plot([5.389, 5.389], [11.5, 13.3888889], color='black', linestyle='-', linewidth=2)
    plt.plot([5.389, 3.5], [13.3888889, 15.3888889], color='black', linestyle='-', linewidth=2)
    plt.plot([1.61111, 3.61111], [13.3888889, 15.3888889], color='black', linestyle='-', linewidth=2)
    #cbar = plt.colorbar(fig)
    #plt.title("Exit Velo Heat Map")
    plt.axis('off')
    newpath = os.path.join("Sheets", names[0])
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    #saves plot in folder
    plt.savefig(os.path.join("Sheets", names[0], 'overheadheatmap.png'),bbox_inches='tight', pad_inches = 0)

    return

def damage_chart(damage_df):
    x = []
    y = []
    array = []
    print(damage_df)
    for i in range(10):
        row = []
        for j in range(8):
            temp_df = damage_df
            top_limit = 4.267222-0.35819444*i
            bottom_limit = 4.267222-0.35819444*(i+1)
            left_limit = -1.4166667 + 0.35416667*j
            right_limit = -1.416667 + 0.35416667*(j+1)

            #drops data from df not in cell needed
            too_left = temp_df[(temp_df['PlateLocSide'] < left_limit)].index
            temp_df = temp_df.drop(too_left)
            too_right = temp_df[(temp_df['PlateLocSide'] > right_limit)].index
            temp_df = temp_df.drop(too_right)
            too_high = temp_df[(temp_df['PlateLocHeight'] > top_limit)].index
            temp_df = temp_df.drop(too_high)
            too_low = temp_df[(temp_df['PlateLocHeight'] < bottom_limit)].index
            temp_df = temp_df.drop(too_low)

            avg_ev_for_zone = round(temp_df["ExitSpeed"].mean(),1)
            
            #if math.isnan(avg_ev_for_zone) == True:
                #avg_ev_for_zone = 60
                

            row.append(avg_ev_for_zone)

        array.append(row)
    df = pd.DataFrame(array)
    print(df)
    df.to_numpy()
    x = np.arange(0, df.shape[1])
    y = np.arange(0, df.shape[0])
    #mask invalid values
    df = np.ma.masked_invalid(df)
    xx, yy = np.meshgrid(x, y)
    #get only the valid values
    x1 = xx[~df.mask]
    y1 = yy[~df.mask]
    newarr = df[~df.mask]

    GD1 = interpolate.griddata((x1, y1), newarr.ravel(),
                            (xx, yy),
                                method='linear', fill_value=60.0)

    df = pd.DataFrame(GD1)
    



    #for a in range(6):
    #    if math.isnan(df.iloc[0,a]):
    #        if a > 0 and a < 6:
    #            df.iloc[0,a] = (df.iloc[0,a-1]+df.iloc[0,a+1])




    #df_smooth = gaussian_filter(df, sigma=1)
    #sns.heatmap(df_smooth, cmap='Spectral_r')
    print(df)

    fig, ax = plt.subplots()
    rect = patches.Rectangle((1.5, 1.5), 4, 6, linewidth=1, edgecolor='black', facecolor='none')
    ax.add_patch(rect)
    fig = plt.imshow(df, cmap = 'jet',vmin=60,vmax=100,  interpolation='bicubic')
    #cbar = plt.colorbar(fig)
    fig.axes.get_xaxis().set_visible(False)
    fig.axes.get_yaxis().set_visible(False)
    #plt.title("Exit Velo Heat Map")
    newpath = os.path.join("Sheets", names[0])
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    #saves plot in folder
    plt.axis('off')
    plt.savefig(os.path.join("Sheets", names[0], 'heatmap.png'), bbox_inches='tight', pad_inches = 0)



    

    
    #print(df_smooth)









    return

def presentation (tabledata):
    prs = Presentation("Template.pptx")
    prs.slide_width = Inches(11)
    prs.slide_height = Inches(8.5)
    #slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide = prs.slides.get(257)

    name = names[0].split()
    print(name)
    full_name = name[-1] + ' ' + name[0]
    full_name = full_name[:-1]
    title = slide.shapes.title
    title.text = full_name
    title.text_frame.paragraphs[0].font.size = Pt(32)
    title.text_frame.paragraphs[0].font.name = 'Beaver Bold'
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # ---add table to slide---
    x, y, cx, cy = Inches(0.15), Inches(0.95), Inches(10.7), Inches(1)
    shape = slide.shapes.add_table(2, 8, x, y, cx, cy)
    table = shape.table

    tbl =  shape._element.graphic.graphicData.tbl
    style_id = '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}'
    tbl[0][-1].text = style_id

    #creating labels for values in all tables
    cell = table.cell(0, 0)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.text = 'AVG EV'
    cell.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
    cell.text_frame.paragraphs[0].font.size = Pt(22)
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    cell01 = table.cell(0,1)
    cell01.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell01.text = 'MAX EV'
    cell01.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
    cell01.text_frame.paragraphs[0].font.size = Pt(22)
    cell01.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    cell02 = table.cell(0,2)
    cell02.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell02.text = "Swing %"
    cell02.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
    cell02.text_frame.paragraphs[0].font.size = Pt(22)
    cell02.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER    

    cell03 = table.cell(0,3)
    cell03.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell03.text = "Chase %"
    cell03.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
    cell03.text_frame.paragraphs[0].font.size = Pt(22)
    cell03.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell04 = table.cell(0,4)
    cell04.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell04.text = "Strikeout %"
    cell04.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
    cell04.text_frame.paragraphs[0].font.size = Pt(20)
    cell04.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell05 = table.cell(0,5)
    cell05.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell05.text = "BB+HBP/K"
    cell05.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
    cell05.text_frame.paragraphs[0].font.size = Pt(22)
    cell05.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell06 = table.cell(0,6)
    cell06.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell06.text = "BABIP"
    cell06.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
    cell06.text_frame.paragraphs[0].font.size = Pt(22)
    cell06.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    cell06 = table.cell(0,7)
    cell06.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell06.text = "wOBA"
    cell06.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
    cell06.text_frame.paragraphs[0].font.size = Pt(22)
    cell06.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


    #nested for loop which puts vales from averages into table
    for j in range(8):
        cell = table.cell((1),j)
        cell.text = tabledata[j]
        cell.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
        cell.text_frame.paragraphs[0].font.size = Pt(30)
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE 
        
    prs.save(os.path.join("Sheets", names[0], names[0] + '.pptx'))

    x, y, cx, cy = Inches(0.15), Inches(2.15), Inches(10.7), Inches(1)
    shape2 = slide.shapes.add_table(2, 14, x, y, cx, cy)
    table2 = shape2.table

    tbl =  shape2._element.graphic.graphicData.tbl
    style_id = '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}'
    tbl[0][-1].text = style_id
    text = ['PA', 'AB', 'H', '2B', '3B', 'HR', 'RBI*', 'BB', 'HBP', 'K', 'AVG', 'OBP','SLG', 'OPS']
    #creating labels for values in all tables
    for i in range(14):
        cell = table2.cell(0, i)
        cell.text = text[i]
        cell.text_frame.paragraphs[0].font.size = Pt(22)
        cell.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    for j in range(14):
        cell = table2.cell((1),j)
        cell.text = tabledata[j+8]
        cell.text_frame.paragraphs[0].font.name = 'Bahnschrift Condensed'
        if j > 9:
            cell.text_frame.paragraphs[0].font.size = Pt(22)
        else:
            cell.text_frame.paragraphs[0].font.size = Pt(24)
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE 

    img_swing_chart = os.path.join(os.path.join("Sheets", names[0], 'swingchart.png'))
    shape = slide.shapes.add_picture(img_swing_chart,Inches(0.30),Inches(3.8), width=Inches(3.6), height=Inches(3.85))
    line = shape.line
    line.color.rgb = RGBColor(117, 117, 117)
    line.width = Inches(0.05)
    
    prs.save(os.path.join("Sheets", names[0], names[0] + '.pptx'))

    img_heat_map = os.path.join(os.path.join("Sheets", names[0], 'heatmap.png'))
    shape = slide.shapes.add_picture(img_heat_map,Inches(4.35),Inches(3.8), width=Inches(3.07791), height=Inches(3.85))
    line = shape.line
    line.color.rgb = RGBColor(117, 117, 117)
    line.width = Inches(0.05)

    prs.save(os.path.join("Sheets", names[0], names[0] + '.pptx'))

    img_heat_map_overhead = os.path.join(os.path.join("Sheets", names[0], 'overheadheatmap.png'))
    shape = slide.shapes.add_picture(img_heat_map_overhead,Inches(7.87791),Inches(3.8), width=Inches(1.80501), height=Inches(3.85))
    line = shape.line
    line.color.rgb = RGBColor(117, 117, 117)
    line.width = Inches(0.05)

    prs.save(os.path.join("Sheets", names[0], names[0] + '.pptx'))

def pitch_strike_called_df():
    global csv_df
    player_df = csv_df

    #Drops all rows where player isn't hitting
    #Drops all rows where player didn't swing
    player_df = player_df.drop(player_df[player_df.PitchCall == 'BallCalled'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'StrikeSwinging'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'FoulBall'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'BallinDirt'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'HitByPitch'].index)
    player_df = player_df.drop(player_df[player_df.PitchCall == 'InPlay'].index)
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
    print(player_df.to_string())


    return player_df

def pitch_loc_chart(player_df):
    #Pulls image for background will have to imput if statemnt to deptermine right vs left
    rhhs = ['Burke, Isaiah','Cedillo, Ruben']
    img = plt.imread("RHH.png")
    fig, ax = plt.subplots(figsize=(6, 6))
    sns.set_style("white")
    #Creates density plot, camp is color scheme and alpha is transperacy
    chart = plt.scatter(x=player_df.PlateLocSide, y=player_df.PlateLocHeight, )
    #creates demenstions for graph plus displays image
    ax.imshow(img, extent=[-2.63,2.665,-0.35,5.30], aspect=1)
      # remove the ticks
    #rect = patches.Rectangle((-0.708333, 1.6466667), 1.4166667, 1.90416667, linewidth=1, edgecolor='black', facecolor='none')
    #ax.add_patch(rect)
    #creates path for plot to be saved in there isn't one
    newpath = os.path.join("Sheets", names[0])
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    #saves plot in folder
    #plt.savefig(os.path.join("Sheets", names[0], 'swingchart.png'))
    plt.show()

    return

swing2d_density_plot(csv_to_swing_df())
damage_chart(data_frame_for_damage_chart())
damage_chart_overhead(data_frame_for_overhead_damage_chart())
presentation(find_table_metrics())
#pitch_loc_chart(pitch_strike_called_df())

