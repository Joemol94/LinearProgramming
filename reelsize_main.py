import pandas as pd
import numpy as np




def determine_ply(row):
    if any(pd.isna(row[['L1', 'F1', 'L2', 'F2', 'BL']])):
        return 3
    return 5
def board_width_calc(df):
    if df['BOX TYPE']=='RSC':
        b_width = df['WIDTH(mm)']+df['HEIGHT(mm)']+df['FLAP MEET(mm)']
    elif df['BOX TYPE']=='DIE CUT':
        b_width = df['WIDTH(mm)']+df['HEIGHT(mm)']+df['width_factor']
    else:
        b_width = df['WIDTH(mm)']+df['HEIGHT(mm)']+df['FLAP MEET(mm)']
    return b_width
def board_length_calc(df):
    if df['BOX TYPE']=='RSC':
        b_length = 2*(df['WIDTH(mm)']+df['LENGTH(mm)'])+df['Flap']+df['Length Trim']
    elif df['BOX TYPE']=='DIE CUT':
        b_length = 2*(df['WIDTH(mm)']+df['LENGTH(mm)'])+df['Flap']+df['Length Trim']
    else:
        b_length = df['LENGTH(mm)']
    return b_length
def board_height_calc(df):
    if df['BOX TYPE']=='RSC':
        b_height = df['HEIGHT(mm)']
    elif df['BOX TYPE']=='DIE CUT':
        b_height = df['HEIGHT(mm)']
    else:
        b_height = df['HEIGHT(mm)']
    return b_height
def board_flap_calc(df):
    if df['BOX TYPE']=='RSC':
        b_flap = (df['WIDTH(mm)']/2)+(df['FLAP MEET(mm)']/2)
    elif df['BOX TYPE']=='DIE CUT':
        b_flap = (df['WIDTH(mm)']/2)+(df['FLAP MEET(mm)']/2)
    else:
        b_flap = 0
    return b_flap
def first_cut(df,n_cut):
    cut_value = (df['BOARD WIDTH']*n_cut)+df['Width Trim']
    return cut_value

def find_closest_to_2200(row):
    closest_value = min(row, key=lambda x: abs(x - 2200))
    column_name = row.index[row.tolist().index(closest_value)]
    return closest_value,column_name

def determine_deckle_size(row,reel_standards):
    closest_cut = row['closest_cut']
    closest_std = min(reel_standards, key=lambda x: abs(x - closest_cut))
    
    if (closest_cut - closest_std) > 6:
        return nearest_std_value(closest_cut,reel_standards)
    else:
        return closest_std

def nearest_std_value(value,reel_standards):
    if value in reel_standards:
        return value
    sorted_standards = sorted(reel_standards, key=lambda x: abs(x - value))
    next_closest_value = sorted_standards[1] 
    return next_closest_value


def overall_calc(df_raw):
    flutestyle_assumption_df = pd.read_excel("assumptions.xlsx",sheet_name="flutestyle")
    boxtype_assumption_df = pd.read_excel("assumptions.xlsx",sheet_name="boxtype")
    reel_std_assumption = pd.read_excel("assumptions.xlsx",sheet_name="reel_standards")
    reel_standards = (list(reel_std_assumption['PAPER WIDTH']))
    df_raw['PLY'] = df_raw.apply(determine_ply, axis=1)
    df_raw_cpy = df_raw.copy()
    #Calculation - 'TOTAL GSM', 'FLAP MEET(mm)', 'Width Trim','Length Trim' and 'Length Trim'
    merged_df1 = df_raw_cpy.merge(flutestyle_assumption_df, on='Fluting style', how='inner')
    merged_df1.fillna(0, inplace=True)
    merged_df1['TOTAL GSM'] = merged_df1['L1']+(merged_df1['F1']*merged_df1['F1 factor'])+merged_df1['L2']+(merged_df1['F2']*merged_df1['F2 factor'])+merged_df1['BL']
    merged_df1['TOTAL GSM'] = merged_df1['TOTAL GSM'].round()
    merged_df1.replace(0, np.nan, inplace=True)
    merged_df1['FLAP MEET(mm)'] = merged_df1['PLY'].apply(lambda x: 0 if x == 3 else 4)
    merged_df1['Width Trim'] = merged_df1['PLY'].apply(lambda x: 20 if x == 3 else 30)
    merged_df1['Length Trim'] = merged_df1['PLY'].apply(lambda x: 10 if x == 3 else 10)
    merged_df1['Flap'] = 35
    #Calculation - BOARD WIDTH, BOARD LENGTH, BOARD HEIGHT, BOARD FLAP
    merged_df2 = merged_df1.merge(boxtype_assumption_df,on='BOX TYPE',how='left')
    merged_df2.fillna(0, inplace=True)
    merged_df2['BOARD WIDTH'] = merged_df2.apply(board_width_calc,axis=1)
    merged_df2['BOARD LENGTH'] = merged_df2.apply(board_length_calc,axis=1)
    merged_df2['BOARD HEIGHT'] = merged_df2.apply(board_height_calc,axis=1)
    merged_df2['BOARD FLAP'] = merged_df2.apply(board_flap_calc,axis=1)
    #Calculation - deckle size, UPS
    cuts_columns = ['First cut', 'Second cut', 'Third cut', 'Fourth cut', 'Fifth cut']
    for i in range(1, 6):
        merged_df2[cuts_columns[i - 1]] = merged_df2.apply(first_cut, axis=1, n_cut=i)
    merged_df3 = merged_df2[['First cut','Second cut','Third cut','Fourth cut','Fifth cut']]
    merged_df3 = merged_df3[merged_df3 < 2200]
    deckle_df = pd.DataFrame()
    deckle_values = merged_df3.apply(find_closest_to_2200, axis=1)
    deckle_df[['closest_cut','UPS']] = pd.DataFrame(deckle_values.tolist(), index=deckle_values.index)
    column_name_to_value = {'First cut': 1, 'Second cut': 2, 'Third cut': 3, 'Fourth cut': 4, 'Fifth cut': 5}
    deckle_df['UPS'] = deckle_df['UPS'].map(column_name_to_value)#apply(lambda x: column_name_to_value.get(x))
    deckle_df['DECKLE SIZE(mm)'] = deckle_df.apply(determine_deckle_size, axis=1, reel_standards=reel_standards)
    deckle_df['DECKLE SIZE(mm)'] = deckle_df['DECKLE SIZE(mm)'].astype(int)
    merged_df4 = pd.concat([merged_df2,deckle_df],axis=1)
    #Calculation - Design watse, trim%, area including, area excluding, per board wt, length meter
    merged_df4['DESIGN WASTE(mm)'] = merged_df4['DECKLE SIZE(mm)'] - merged_df4['UPS']*merged_df4['BOARD WIDTH']
    merged_df4[' Trim % '] = round((merged_df4['DESIGN WASTE(mm)']/merged_df4['DECKLE SIZE(mm)'])*100,2)
    merged_df4['AREA INCLUDING TRIM(SQ.M)'] = round((((merged_df4['BOARD WIDTH']*merged_df4['BOARD LENGTH'])+(merged_df4['BOARD WIDTH']*merged_df4['BOARD LENGTH']*(merged_df4[' Trim % ']/100)))/(1000000))*merged_df4['QTY'],3)
    merged_df4['AREA EXCLUDING TRIM(SQ.M)'] = round(((merged_df4['BOARD WIDTH']*merged_df4['BOARD LENGTH'])/(1000000))*merged_df4['QTY'],3)
    merged_df4['Per BOARD WEIGHT'] = round((merged_df4['BOARD WIDTH']*merged_df4['BOARD LENGTH']*merged_df4['TOTAL GSM'])/(1000000),0)
    merged_df4['Length METER'] = round((merged_df4['QTY']/merged_df4['UPS'])*merged_df4['BOARD LENGTH']/1000,0)
    #Calculation - Layerwise L1,F1,L2,F2,BL
    merged_df5 = merged_df4.copy()
    merged_df5['L1(in tons)'] = (merged_df5['Length METER'] * merged_df5['DECKLE SIZE(mm)']*merged_df5['L1'])/(10**9)
    merged_df5['F1(in tons)'] = (merged_df5['Length METER'] * merged_df5['DECKLE SIZE(mm)']*merged_df5['F1']*merged_df5['F1 factor'])/(10**9)
    merged_df5['L2(in tons)'] = (merged_df5['Length METER'] * merged_df5['DECKLE SIZE(mm)']*merged_df5['L2'])/(10**9)
    merged_df5['F2(in tons)'] = (merged_df5['Length METER'] * merged_df5['DECKLE SIZE(mm)']*merged_df5['F2']*merged_df5['F2 factor'])/(10**9)
    merged_df5['BL(in tons)'] = (merged_df5['Length METER'] * merged_df5['DECKLE SIZE(mm)']*merged_df5['BL'])/(10**9)
    final_df = merged_df5.drop(["F1 factor","F2 factor","width_factor","length_factor","closest_cut","SHADE"], axis=1)
    return final_df

def layerwise_requirement(final_df):
    top_layer = final_df.groupby(['TOP PAPER','L1','DECKLE SIZE(mm)'])['L1(in tons)'].sum().reset_index()
    flutting_layer = final_df.groupby(['F1','F2','DECKLE SIZE(mm)'])[['F1(in tons)','F2(in tons)']].sum().reset_index()
    flutting_layer['Flutting Total (in tons)']  = flutting_layer['F1(in tons)']+ flutting_layer['F2(in tons)']
    middle_layer = final_df.groupby(['L2','DECKLE SIZE(mm)'])['L2(in tons)'].sum().reset_index()
    bottom_layer = final_df.groupby(['BL','DECKLE SIZE(mm)'])['BL(in tons)'].sum().reset_index()
    bottom_layer = bottom_layer[bottom_layer['BL'] !=0]
    bottom_layer = bottom_layer[bottom_layer['BL(in tons)'] !=0]
    middle_layer = middle_layer[middle_layer['L2'] !=0]
    middle_layer = middle_layer[middle_layer['L2(in tons)'] !=0]
    flutting_layer = flutting_layer[flutting_layer['Flutting Total (in tons)']!=0]
    top_layer = top_layer[top_layer['L1(in tons)']!=0]
    return top_layer,flutting_layer,bottom_layer,middle_layer


