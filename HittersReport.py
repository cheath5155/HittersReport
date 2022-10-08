from cgi import print_directory
from unicodedata import name
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
import seaborn as sns
import os
from pptx import Presentation
from pptx.util import Pt
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

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
    remove_list.remove("TaggedHitType")

    #Removes all collums other than those with .remove above from data frame
    player_df = player_df.drop(remove_list, axis=1)
    avg_ev_df = player_df.drop(player_df[player_df.ExitSpeed < 60.0].index)
    avg_ev = "%.1f" % round(avg_ev_df["ExitSpeed"].mean(),1)
    max_ev = "%.1f" % round(avg_ev_df["ExitSpeed"].max(),1)

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
    bb_hpb = (player_df["KorBB"] == "Walk").sum() + (player_df["PitchCall"] == "HitByPitch").sum()
    bb_hbp_over_ks = "%.2f" % round((bb_hpb/strikeouts),2)

    #BABIP
    bip = (player_df["PitchCall"] == "InPlay").sum()
    hits = (player_df["PlayResult"] == "Single").sum() + (player_df["PlayResult"] == "Double").sum() + (player_df["PlayResult"] == "Triple").sum() + (player_df["PlayResult"] == "HomeRun").sum()
    babip = "%.3f" % round(hits/bip,3)

    data_to_pass_to_presentation = []
    data_to_pass_to_presentation.append(str(avg_ev))
    data_to_pass_to_presentation.append(str(max_ev))
    data_to_pass_to_presentation.append(str(swing_rate) + "%")
    data_to_pass_to_presentation.append(str(chase_rate) + "%")
    data_to_pass_to_presentation.append(str(k_rate) + "%")
    data_to_pass_to_presentation.append(str(bb_hbp_over_ks))
    data_to_pass_to_presentation.append(str(babip))

    print(data_to_pass_to_presentation)
    return

def data_frame_for_damage_chart():

    return
def damage_chart():
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
    x, y, cx, cy = Inches(0.15), Inches(1), Inches(10.35), Inches(2.5)
    shape = slide.shapes.add_table(2, 7, x, y, cx, cy)
    table = shape.table

    tbl =  shape._element.graphic.graphicData.tbl
    style_id = '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}'
    tbl[0][-1].text = style_id

    #creating labels for values in all tables
    cell = table.cell(0, 0)
    cell.text = 'AVG EV'
    cell.text_frame.paragraphs[0].font.size = Pt(14)
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    cell01 = table.cell(0,1)
    cell01.text = 'MAX EV'
    cell01.text_frame.paragraphs[0].font.size = Pt(14)
    cell01.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    cell02 = table.cell(0,2)
    cell02.text = "Swing %"
    cell02.text_frame.paragraphs[0].font.size = Pt(14)
    cell02.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER    

    cell03 = table.cell(0,3)
    cell03.text = "Chase %"
    cell03.text_frame.paragraphs[0].font.size = Pt(14)
    cell03.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell04 = table.cell(0,4)
    cell04.text = "Strikeout %"
    cell04.text_frame.paragraphs[0].font.size = Pt(14)
    cell04.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell05 = table.cell(0,5)
    cell05.text = "BB+HBP/K"
    cell05.text_frame.paragraphs[0].font.size = Pt(14)
    cell05.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell06 = table.cell(0,6)
    cell06.text = "BABIP"
    cell06.text_frame.paragraphs[0].font.size = Pt(14)
    cell06.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    #nested for loop which puts vales from averages into table
    for j in range(14):
        cell = table.cell((1),j)
        cell.text = tabledata[j]
        cell.text_frame.paragraphs[0].font.size = Pt(14)
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER 
        prs.save(os.path.join("Sheets", names[0], '.pptx'))


#swing2d_density_plot(csv_to_swing_df())
find_table_metrics()