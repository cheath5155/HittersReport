from cgi import print_directory
import math
from cmath import nan
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
from pptx.enum.text import MSO_ANCHOR
from scipy.interpolate import interp2d
from scipy.interpolate import RectBivariateSpline
from scipy.ndimage.filters import gaussian_filter
from scipy import interpolate

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
    return data_to_pass_to_presentation

def data_frame_for_damage_chart():
    global csv_df
    damage_df = csv_df

    remove_list = damage_df.columns.values.tolist()
    remove_list.remove("Batter")
    remove_list.remove("PlateLocHeight")
    remove_list.remove("PlateLocSide")
    remove_list.remove("BatterSide")
    remove_list.remove("ExitSpeed")

    #Removes all collums other than those with .remove above from data frame
    damage_df = damage_df.drop(remove_list, axis=1)

    print(damage_df)
    drop_index = []
    for i in range(damage_df.shape[0]):
        if math.isnan(damage_df.iat[i, 4]) == True:
            drop_index.append(i)

    damage_df = damage_df.drop(drop_index)

    damage_df = damage_df.drop(damage_df[damage_df.ExitSpeed < 60.0].index)

    

    #Switches from Pitcher view to catcher view
    damage_df['PlateLocSide'] = (damage_df['PlateLocSide'] * -1)
    return damage_df

def damage_chart1(damage_df):
    x_values = damage_df['PlateLocSide'].tolist()
    y_values = damage_df['PlateLocHeight'].tolist()
    evs = damage_df['ExitSpeed'].tolist()

    x_list = np.array(x_values)
    y_list = np.array(y_values)
    evs_plot = np.array(evs)




    f = interpolate.RectBivariateSpline(x[0,:], y[:,0], z)

    subplot(1,2,1)
    im = imshow(x[0,:], y[:,0], z, 40)
    colorbar(im)

    xnew, ynew = meshgrid(linspace(-1, 1, 70), linspace(-1, 1, 70))
    znew = f(xnew, ynew)
    subplot(1,2,2)
    im = imshow(xnew[0,:], ynew[:,0], znew, 40)
    colorbar(im)
    suptitle('2-D grid data interpolation example')


    return
def damage_chart(damage_df):
    array = []
    for i in range(7):
        row = []
        for j in range(6):
            temp_df = damage_df
            top_limit = 4.02-0.42*i
            bottom_limit = 4.02-0.42*(i+1)
            left_limit = -1.0625 + 0.35416667*j
            right_limit = -1.0625 + 0.35416667*(j+1)

            #drops data from df not in cell needed
            too_left = temp_df[(temp_df['PlateLocSide'] < left_limit)].index
            temp_df = temp_df.drop(too_left)
            too_right = temp_df[(temp_df['PlateLocSide'] > right_limit)].index
            temp_df = temp_df.drop(too_right)
            too_high = temp_df[(temp_df['PlateLocHeight'] > top_limit)].index
            temp_df = temp_df.drop(too_high)
            too_low = temp_df[(temp_df['PlateLocHeight'] < bottom_limit)].index
            temp_df = temp_df.drop(too_low)

            avg_ev_for_zone = round(temp_df["ExitSpeed"].mean(),0)
            
            if math.isnan(avg_ev_for_zone) == False:
                avg_ev_for_zone = int(avg_ev_for_zone)
            else:
                avg_ev_for_zone = 60

            row.append(avg_ev_for_zone)
        array.append(row)
    df = pd.DataFrame(array)
    #df_smooth = gaussian_filter(df, sigma=1)
    #sns.heatmap(df_smooth, cmap='Spectral_r')
    plt.imshow(df, cmap = 'jet', interpolation='bilinear')
    plt.show()



    
    print(df)
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
    x, y, cx, cy = Inches(0.15), Inches(1), Inches(10.35), Inches(2.5)
    shape = slide.shapes.add_table(2, 7, x, y, cx, cy)
    table = shape.table

    tbl =  shape._element.graphic.graphicData.tbl
    style_id = '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}'
    tbl[0][-1].text = style_id

    #creating labels for values in all tables
    cell = table.cell(0, 0)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.text = 'AVG EV'
    cell.text_frame.paragraphs[0].font.size = Pt(22)
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    cell01 = table.cell(0,1)
    cell01.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell01.text = 'MAX EV'
    cell01.text_frame.paragraphs[0].font.size = Pt(22)
    cell01.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    cell02 = table.cell(0,2)
    cell02.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell02.text = "Swing %"
    cell02.text_frame.paragraphs[0].font.size = Pt(22)
    cell02.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER    

    cell03 = table.cell(0,3)
    cell03.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell03.text = "Chase %"
    cell03.text_frame.paragraphs[0].font.size = Pt(22)
    cell03.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell04 = table.cell(0,4)
    cell04.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell04.text = "Strikeout %"
    cell04.text_frame.paragraphs[0].font.size = Pt(22)
    cell04.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell05 = table.cell(0,5)
    cell05.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell05.text = "BB+HBP/K"
    cell05.text_frame.paragraphs[0].font.size = Pt(22)
    cell05.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  

    cell06 = table.cell(0,6)
    cell06.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell06.text = "BABIP"
    cell06.text_frame.paragraphs[0].font.size = Pt(22)
    cell06.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    


    #nested for loop which puts vales from averages into table
    for j in range(7):
        cell = table.cell((1),j)
        cell.text = tabledata[j]
        cell.text_frame.paragraphs[0].font.size = Pt(30)
        cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE 
        
    prs.save(os.path.join("Sheets", names[0], names[0] + '.pptx'))


#swing2d_density_plot(csv_to_swing_df())
#presentation(find_table_metrics())
damage_chart(data_frame_for_damage_chart())
