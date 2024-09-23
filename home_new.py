import dash
import dash_bootstrap_components as dbc
import dash_core_components as dcc
import dash_html_components as html
import dash_table as dt
from dash.dependencies import Input, Output,State 
from dash import no_update
from server import app, server
from flask import Flask,request

import pandas as pd
import os
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.application import MIMEApplication
import io
from pulp import *
from datetime import datetime
import numpy as np
import main
import gsm_main
import reelsize_main




base_path = r"path_to_project\carton project\GSM_mod/"
#base_path = r"/home/ubuntu/demand_optimization/"
reel_base_path = r"path_to_project\carton project\reel_optimization/"
reelsize_data = pd.read_excel(reel_base_path+ "reel_size_input.xlsx")
#gsm_data = pd.read_excel(base_path+"data/sample_data_full_gsm.xlsx")
gsm_data = pd.read_excel(base_path+"data/GSM_data_newformat.xlsx")


NAVBAR_STYLE={
    "width":"1275px",#1530
    "height": "45px",
    #"position": "fixed",
    "float":"left",
    "fluid":True,
    "top": 0,
    "left":2,
    "fontFamily": "calibri,Tahoma Regular, sans-serif",
    "fontSize": "6px",
    "fontColor":"white",
    "border":"1px gray lightgray",
    "backgroundColor": '#228B22',#'#D10A2D',#'#FA395B',#"#F64060",
    "textAlign":"center",
    "boxShadow": "3px 3px 3px lightgray",
    }

image_style={
    "height": "45px",
    "width":"146px",
    #"position": "fixed",
    "float":"left",
    "fluid":True,
    "top": 0,
    "left":2,
    "fontFamily": "calibri,Tahoma Regular, sans-serif",
    "fontSize": "6px",
    "fontColor":"dark",
    "backgroundColor": '#228B22',#"#D10A2D",
    "align":"left"   
    }

tabs_styles = {
    'height': '45px',
    "width":"1250px",
    "margin-left": "1rem",
    #"right": "0.5rem",
    #"position": "fixed",
    #"fluid":"True",
    #"top":500,
    #"left":2,
    "font-family": "calibri,Tahoma Regular, sans-serif",
    "border-top": "3px solid lightgray",
    "background-color":"#fafbfc",
    #"background": "white", 
    "border": "lightgray",
    "sticky":"top"   
}

tab_style = {
    'borderBottom': '1px solid lightgray',
    "borderTop": "1px solid lightgray",
    'padding': '6px',
    'fontWeight': 'bold',
    "background": "#ffffff", 
    #"background":"#42c2b3",
    #"background-color":"#286DA8",
    #"border": "#286DA8",
    "boxShadow": "3px 3px 3px lightgray",
}

tab_style1 = {
    'borderBottom': '1px solid lightgray',
    "borderTop": "1px solid lightgray",
    'padding': '6px',
     "width":"1300px",
    'fontWeight': 'bold',
    #"background": "#42c2b3", 
    "background-color":"#ffffff",
    #"border": "#286DA8",
    #"background": "",
    "boxShadow": "3px 3px 3px lightgray",
}

tab_selected_style = {
    #'borderTop': '1px solid #d6d6d6',
    'borderBottom': '1px solid lightgray',
    #'backgroundColor': 'white',
    'color': 'dark',
    'padding': '6px',
    "background": '#228B22',#"#D10A2D", 
    "boxShadow": "3px 3px 3px lightgray",
    "borderLeft": "1px solid lightgray",
    "borderRight": "1px solid lightgray",
    "borderTop": "1px solid lightgray"
    
}
FIVE_FILTER_STYLE={
    "width":"200px",
    #"height": "40px",
    #"position": "fixed",
    #"fluid":True,
    #"float":"left",
    #"top":50,
    #"left":40,
    #"margin-left": "0px",
    #"margin-right": "2rem",
    #"margin-top": "1rem",
    #"right":0,
    "font-family": "calibri,Tahoma Regular, sans-serif",
    "font-size": "16px",
    "font-color":"dark",
    "align":"centre",
    # "border":"1px lightgray solid",
    #"boxShadow": "3px 3px 3px lightgray",
    "background-color": "#ffffff",
    "text-color":"black"
    }
FIVE_FILTER_CARD_STYLE={
    "width":"230px",
    "height": "95px",
    #"margin-left": "2rem",
    "margin-top": "1.25rem",
    "margin-bottom": "1.25rem",
    "padding-top": "5px",
    #"position": "fixed",
    #"fluid":True,
    #"float":"left",
    #"top": 140,
    #"left": 10,
    "font-family": "calibri,Tahoma Regular, sans-serif",
    "font-size": "16px",
    "font-color":"dark",
    "font-style":"bold",
    "font-weight":"700",
    #"align":"centre",
    #"border":"3px lightgray solid",
    "boxShadow": "3px 3px 3px lightgray",
    "background-color": "#ffffff",
    "text-color":"black"
    }
NAVBAR_STYLE3={
    "width":"1515px",
    #"height": "40px",
    #"position": "fixed",
    #"fluid":True,
    #"float":"left",
    #"top": 200,
    #"left":40,
    "margin-left": "2rem",
    #"margin-right": "2rem",
    #"margin-top": "40rem",
    "font-family": "calibri,Tahoma Regular, sans-serif",
    "font-size": "16px",
    "font-color":"dark",
    #"border":"1px lightgray solid",
    #"boxShadow": "3px 3px 3px lightgray",
    "background-color": "#ffffff",
    "text-color":"black",
    #"align":"centre"
    }
CARD4_STYLE={
    "width":"1200px",#1472 #750
    #"height": "400px",
    #"position": "fixed",
    #"fluid":True,
    #"float":"left",
    #"top": 140,
    #"left": 10,
    #"margin-left": "2rem",
    "margin-top": "1.25rem",
    #"margin-right": "2rem",
    #"margin-bottom": "2rem",
    "font-family": "calibri,Tahoma Regular, sans-serif",
    "font-size": "12px",
    "font-color":"dark",
    #"align":"center",
    #"border":"3px lightgray solid",
    #"boxShadow": "3px 3px 3px lightgray",
    "background-color": "#ffffff",
    "text-color":"black"
    }

CARD_STYLE01={
    "width":"275px",#165
    "height": "40px",
    #"margin-left": "85rem",
    "margin-top": "1.25rem",
    #"position": "fixed",
    "fluid":True,
    "float":"left",
    #"top": 140,
    #"left": 1125,
    "font-family": "calibri,Tahoma Regular, sans-serif",
    "font-size": "16px",
    "font-color":"dark",
    "font-style":"bold",
    "font-weight":"700",
    "align":"center",
    #"border":"3px lightgray solid",
    "boxShadow": "3px 3px 3px lightgray",
    "background-color": "#ffffff",
}
CARD_STYLE_3 = {
    #"margin-bottom":"60",
    #"top": 50,
    #"left": 0,
    "margin-top": "0.25rem",
    #"margin-left": "886px",
    "margin-bottom": "0.25rem",
    #"background-color": "white",
    #"position": "fixed",
    "fluid":True,
    "height": "10px",
    "float":"left",
    "width": "1470px",
    #"height":"50px",
    #"color": "white",
    "align": "center",
    "border-style": "none",
    #"boxShadow": "0px 0px 0px #ffffff",
    # "padding":"5px 5px 5px 15px",
    "font-family": "calibri,Tahoma Regular, sans-serif",
    }

HEADER_STYLE = {
    #"margin-bottom":"60",
    #"top": 50,
    #"left": 0,
    #"margin-top": "0.5rem",
    #"margin-left": "2rem",
    #"background-color": "white",
    #"position": "fixed",
    "fluid":True,
    "float":"left",
    "width": "200px",
    #"height":"50px",
    "color": "white",
    "align": "center",
    #"padding":"20px",
    "font-family": "calibri,Tahoma Regular, sans-serif",
    }

email_recipient_box_style={
    'width': "450px",
    'height': '40px',
    'color': '#000000',
    "font-style":"bold",
    "font-size": "16px",
    'font-weight':"400",
    "margin-left":"16rem",
    "top":3,
    "margin-top":"0.25rem",
    #"border":"2px 2px 2px 2px gray",
    #'borderStyle': 'solid',
    "textAlign":"left",
    "borderColor": '#228B22',#"#D10A2D",
    "font-family": "calibri,Tahoma Regular, sans-serif",
    "font-color":"green",
    #"boxShadow": "2px 2px 2px 2px lightgray",
    'borderWidth': '1px',
    'borderStyle': 'solid',
    'borderRadius': '20px',
    "boxShadow": "3px 3px 3px lightgray",

}

email_report_style = {
'width': '100px',
'height': '50px',
"top": 3,
'borderWidth': '1px',
'borderStyle': 'solid',
'borderRadius': '20px',
'textAlign': 'center',
"margin-left": '1rem',#2.5
"font-family": "calibri,Tahoma Regular, sans-serif",
"font-size": "15px",
"fontColor":"white",
"boxShadow": "3px 3px 3px lightgray",
"borderColor": '#228B22',#"#D10A2D",
"backgroundColor":'#228B22',#"#D10A2D",#42c2b3",
"font-style":"bold",
}

PLOTLY_LOGO= "../assets/"
location = dcc.Location(id='url1',pathname='/home')




navbar = dbc.Navbar(
    [

        #dbc.NavbarToggler(id="navbar-toggler"),
        dbc.Collapse(id="navbar-collapse", navbar=True),
        #dbc.Col(html.H3("GREEN PACKAGINGS"),md=14,style={"color":"white","textAlign":"center"}),
        dbc.Col([html.Img(src=app.get_asset_url("GPL_logo.jpg"), height="40px",width = "120px")],style=HEADER_STYLE),
        dbc.Col(
                    #dcc.Upload(id='upload-data',children=html.Div([html.H5('Upload Files')]),multiple=True,style={"color":"white","margin-left": '24rem'})
                ),  

    ],
    color="#ffffff",#' #ffffff #557bec #"#FF1139"
    dark=False,
    style=NAVBAR_STYLE
)

tab_layout= html.Div([
        dcc.Tabs(id="tab_switch", value='reel_tab', children=[
        #dcc.Tab(label='Demand Optimization', value='demand_tab', style=tab_style, selected_style=tab_selected_style),
        dcc.Tab(label='Reel size Optimization', value='reel_tab', style=tab_style, selected_style=tab_selected_style),
        dcc.Tab(label='GSM Optimization', value='gsm_tab', style=tab_style, selected_style=tab_selected_style),
        dcc.Tab(style=tab_style1,selected_style=tab_style1)
    ], style=tabs_styles),
    #html.Br(),html.Br(),
    html.Div(id='tab_switch_content')
])



main_df = pd.read_excel(base_path+"data/sample_data.xlsx")
leng = {}
length_list =['ALL']+list(main_df.length.unique())
leng['ALL']  = length_list

gsmv = {}
gsm_list =['ALL']+list(main_df.gsm.unique())
gsmv['All'] = gsm_list

paper_typv = {}
paper_typ_list = ['ALL','White Top Kraft Line','Fluting SCF','Kraft Brown']
paper_typv['ALL'] = paper_typ_list
length_filter = dbc.Card(
    [
     dbc.Row([
    dbc.Col(
    dbc.FormGroup(
            [
                dbc.Label("Length"),
                dcc.Dropdown(id="length",
                    options=[{"label": i, "value": i} for i in length_list],
                    value=length_list[0],
                    multi=False,
                    clearable=False,
                    style={"boxShadow": "3px 3px 3px lightgray"},
                ),
            ]),
         ),
    ],style=FIVE_FILTER_STYLE),     
    ],
    body=True,
    style=FIVE_FILTER_CARD_STYLE    
)

trim_filter = dbc.Card(
    [
     dbc.Row([
    dbc.Col(
    dbc.FormGroup(
            [
                dbc.Label("Trim value"),
                # dcc.Dropdown(id="trim",
                #     options=[{"label": i, "value": i} for i in trim_list],
                #     value=trim_list[0],
                #     multi=False,
                #     clearable=False,
                #     style={"boxShadow": "3px 3px 3px lightgray"},
                # ),
                dcc.Input(id='trim_input_box', type='search', placeholder='Enter a value', 
                autoComplete = 'off', autoFocus=True, disabled=False, n_submit=0)
            ]),
    
         ),
    ],style=FIVE_FILTER_STYLE),  
    ],
    body=True,
    style=FIVE_FILTER_CARD_STYLE    
)
gsm_filter = dbc.Card(
    [
     dbc.Row([
    dbc.Col(
    dbc.FormGroup(
            [
                dbc.Label("GSM value"),
                dcc.Dropdown(id="gsm",
                    options=[{"label": i, "value": i} for i in gsm_list],
                    value=gsm_list[0],
                    multi=False,
                    clearable=False,
                    style={"boxShadow": "3px 3px 3px lightgray"},
                ),
            ]),
         ),
    ],style=FIVE_FILTER_STYLE),  
    ],
    body=True,
    style=FIVE_FILTER_CARD_STYLE    
)

row0= dbc.Nav([
                dbc.Card([
                    dbc.Row([
                    dbc.Col([html.Img(src=app.get_asset_url("GPL_logo.jpg"), height="60px")],style=HEADER_STYLE),#tolaram-logo.png
                    ])
                    ],style=CARD_STYLE_3),
                html.Div([dcc.ConfirmDialog(id='user_input_alert')]),
        ],style = NAVBAR_STYLE3)

filter_row = dbc.Nav([
                dbc.Row([
                    dbc.Col([
                        dbc.Row([dbc.Col(length_filter),dbc.Col(gsm_filter),dbc.Col(trim_filter)])
                    ]),
                ]),],style=NAVBAR_STYLE3)

table_row= dbc.Nav([
              dbc.Row([
                dbc.Col([
                            dbc.Card([
                                        #html.Div(id="tableout2",style=CARD3_STYLE)
                                        html.Div([dt.DataTable(id="demand_table",#dm_tableout2
                                                                       sort_action='native',
                                                                       fixed_rows={'headers':True},
                                                                       #page_size=10,
                                                                       editable=True,
                                                                       style_data={'whiteSpace': 'normal', 'height': 'auto','border': '1px solid rgb(220,220,220)', 'fontWeight': 'bold'},
                                                                       style_header={'whiteSpace': 'normal','textAlign': 'center','fontWeight': 'bold', "color":"white",
                                                                                     'height': 'auto','backgroundColor': '#228B22',#'#D10A2D',
                                                                                     'border': '1px solid rgb(220, 220, 220)' },
                                                                       style_cell={'textAlign': 'center', 'padding': '0px', 'width': 'auto',
                                                                                   'overflow': 'hidden','textOverflow': 'ellipsis'},
                                                                       style_data_conditional=[
                                                                           {
                                                                               'if': {
                                                                                   'column_editable': False
                                                                                   },
                                                                               'backgroundColor': 'rgb(240, 240, 240)',
                                                                               'cursor': 'not-allowed'
                                                                           },                                                                            
                                                                           {
                                                                               'if': {
                                                                                   'state': 'active',  # 'active' | 'selected'
                                                                                   
                                                                               },
                                                                               #'backgroundColor': 'rgb(66, 194, 179)',
                                                                               'backgroundColor':'#228B22',#'#D10A2D',
                                                                               'border': '1px solid rgb(66, 194, 179)'
                                                                           },
                                    
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"sku"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"papertype"
                                                                          
                                                                                     },
                                                                                     'width':'80px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"length"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"gsm"
                                                                          
                                                                                     },
                                                                                     'width':'110px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"width_type"
                                                                          
                                                                                     },
                                                                                     'width':'110px',
                                                                           },

                                                                            {
                                                                                 'if':{
                                                                                     'column_id':"demand"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                        ],
                                                                       style_table={'height':'400px'},
                                                                       #export_format='xlsx',
                                                                       #export_headers='display',
                                                                       filter_action= "native",
                                                                       merge_duplicate_headers=True,)])
                                    ],style=CARD4_STYLE),
                                    html.Div([
                                        dbc.Alert("Submit the demand table to get optimized results",dismissable=True)]
                                        ,style=CARD4_STYLE),
                                    #html.Div(id='dm_table_alert2'),
                            ]),
                       ])
                ],style=NAVBAR_STYLE3)

gsm_input_desc = dbc.Nav([
                        dbc.Row([
                                dbc.Col([
                                        html.Div([
                                            dbc.Alert("Use the template to upload the updated values for all paper types",dismissable=False)])
                                        ]),
                                    ]),
                                ],style=NAVBAR_STYLE3) 

reel_assumption_desc = dbc.Nav([
                        dbc.Row([
                                dbc.Col([
                                        html.Div([
                                            dbc.Alert("Use the template to update the assumption table",dismissable=False)])
                                        ]),
                                    ]),
                                ],style=NAVBAR_STYLE3) 
     


gsm_user_input_row = dbc.Nav([
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            html.Div([
                                dbc.Button("Template", id="gsm_template_btn",color='#228B22',#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center', 'width': '270px'}),#'#D10A2D',
                                dcc.Download(id="gsm_template_link"),
                                ]),
                        ]),
                        
                    ]),
                    dbc.Col([
                        dbc.Card([
                                html.Div([
                                dcc.Upload(id="upload_gsm_user_input",
                                           children = html.Div([html.A("Upload Bucket values")]),
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center','width': '270px', 'cursor': 'pointer','height':'35px'})#'padding-top':'5px','#D10A2D',
                               ]),
                               dcc.Store(id='gsm_stored_data_tl_white'),
                               dcc.Store(id='gsm_stored_data_flute'),
                               dcc.Store(id='gsm_stored_data_brown_l2_l3'),
                               html.Div([dcc.ConfirmDialog(id='gsm_user_upload_msg')]),
                               html.Div(dcc.ConfirmDialog(id='gsm_user_input_msg')),
                        ]),
                        
                    ]),
                ])
              ],style=NAVBAR_STYLE3)
reel_assumption_input_row = dbc.Nav([
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            html.Div([
                                dbc.Button("Assumption Table", id="reel_assump_btn",color='#228B22',#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center', 'width': '270px'}),#'#D10A2D',
                                dcc.Download(id="reel_assump_link"),
                                ]),
                        ]),
                        
                    ]),
                    dbc.Col([
                        dbc.Card([
                                html.Div([
                                dcc.Upload(id="reel_upload_assump_input",
                                           children = html.Div([html.A("Upload Assumption values")]),
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center','width': '270px', 'cursor': 'pointer','height':'35px'})#'padding-top':'5px','#D10A2D',
                               ]),
                               dcc.Store(id='assumption_data_store'),
                               html.Div([dcc.ConfirmDialog(id='reel_assump_upload_msg')]),
                        ]),
                        
                    ]),
                ])
              ],style=NAVBAR_STYLE3)


gsm_table_row= dbc.Nav([
              dbc.Row([
                dbc.Col([
                            dbc.Card([
                                        #html.Div(id="tableout2",style=CARD3_STYLE)
                                        html.Div([dt.DataTable(id="gsm_table",#dm_tableout2
                                                                       sort_action='native',
                                                                       fixed_rows={'headers':True},
                                                                       #page_size=10,
                                                                       editable=True,
                                                                       style_data={'whiteSpace': 'normal', 'height': 'auto','border': '1px solid rgb(220,220,220)', 'fontWeight': 'bold'},
                                                                       style_header={'whiteSpace': 'normal','textAlign': 'center','fontWeight': 'bold', "color":"white",
                                                                                     'height': 'auto','backgroundColor': '#228B22',#'#D10A2D',
                                                                                     'border': '1px solid rgb(220, 220, 220)' },
                                                                       style_cell={'textAlign': 'center', 'padding': '0px', 'width': 'auto',
                                                                                   'overflow': 'hidden','textOverflow': 'ellipsis'},
                                                                       style_data_conditional=[
                                                                           {
                                                                               'if': {
                                                                                   'column_editable': False
                                                                                   },
                                                                               'backgroundColor': 'rgb(240, 240, 240)',
                                                                               'cursor': 'not-allowed'
                                                                           },                                                                            
                                                                           {
                                                                               'if': {
                                                                                   'state': 'active',  # 'active' | 'selected'
                                                                                   
                                                                               },
                                                                               #'backgroundColor': 'rgb(66, 194, 179)',
                                                                               'backgroundColor':'#228B22',#'#D10A2D',
                                                                               'border': '1px solid rgb(66, 194, 179)'
                                                                           },
                                    
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"sku"
                                                                          
                                                                                     },
                                                                                     'width':'110px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':'topliner'
                                                                          
                                                                                     },
                                                                                     'width':'110px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"length"
                                                                          
                                                                                     },
                                                                                     'width':'70px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"ply"
                                                                          
                                                                                     },
                                                                                     'width':'50px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"width"
                                                                          
                                                                                     },
                                                                                     'width':'70px',
                                                                           },

                                                                            {
                                                                                 'if':{
                                                                                     'column_id':"target_gsm"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"tl_l1"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"tl_f1"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"tl_l2"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"tl_f2"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"tl_l3"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"total_optimized_gsm"
                                                                          
                                                                                     },
                                                                                     'width':'90px',
                                                                           },
                                                                        ],
                                                                       style_table={'height':'400px'},
                                                                       #export_format='xlsx',
                                                                       #export_headers='display',
                                                                       filter_action= "native",
                                                                       merge_duplicate_headers=True,)])
                                    ],style=CARD4_STYLE),
                                    html.Div([
                                        dbc.Alert("Click 'Calculate' to determine Optimised GSM values",dismissable=True)]
                                        ,style=CARD4_STYLE),
                                    #html.Div(id='dm_table_alert2'),
                            ]),
                       ])
                ],style=NAVBAR_STYLE3)

reel_table_row= dbc.Nav([
              dbc.Row([
                dbc.Col([
                            dbc.Card([
                                        #html.Div(id="tableout2",style=CARD3_STYLE)
                                        html.Div([dt.DataTable(id="reel_table",#dm_tableout2
                                                                       sort_action='native',
                                                                       fixed_rows={'headers':True},
                                                                       #page_size=10,
                                                                       editable=True,
                                                                       style_data={'whiteSpace': 'normal', 'height': 'auto','border': '1px solid rgb(220,220,220)', 'fontWeight': 'bold'},
                                                                       style_header={'whiteSpace': 'normal','textAlign': 'center','fontWeight': 'bold', "color":"white",
                                                                                     'height': 'auto','backgroundColor': '#228B22',#'#D10A2D',
                                                                                     'border': '1px solid rgb(220, 220, 220)' },
                                                                       style_cell={'textAlign': 'center', 'padding': '0px', 'width': 'auto',
                                                                                   'overflow': 'hidden','textOverflow': 'ellipsis'},
                                                                       style_data_conditional=[
                                                                           {
                                                                               'if': {
                                                                                   'column_editable': False
                                                                                   },
                                                                               'backgroundColor': 'rgb(240, 240, 240)',
                                                                               'cursor': 'not-allowed'
                                                                           },                                                                            
                                                                           {
                                                                               'if': {
                                                                                   'state': 'active',  # 'active' | 'selected'
                                                                                   
                                                                               },
                                                                               #'backgroundColor': 'rgb(66, 194, 179)',
                                                                               'backgroundColor':'#228B22',#'#D10A2D',
                                                                               'border': '1px solid rgb(66, 194, 179)'
                                                                           },
                                    
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"Fluting style"
                                                                          
                                                                                     },
                                                                                     'width':'75px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':'Desp'
                                                                          
                                                                                     },
                                                                                     'width':'120px',
                                                                           },
                                                                           {      
                                                                                 'if':{
                                                                                     'column_id':"LENGTH(mm)"
                                                                          
                                                                                     },
                                                                                     'width':'100px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"WIDTH(mm)"
                                                                          
                                                                                     },
                                                                                     'width':'100px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"HEIGHT(mm)"
                                                                          
                                                                                     },
                                                                                     'width':'100px',
                                                                           },

                                                                            {
                                                                                 'if':{
                                                                                     'column_id':"TOP PAPER"
                                                                          
                                                                                     },
                                                                                     'width':'60px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"BOX TYPE"
                                                                          
                                                                                     },
                                                                                     'width':'80px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"SHADE"
                                                                          
                                                                                     },
                                                                                     'width':'80px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"L1"
                                                                          
                                                                                     },
                                                                                     'width':'50px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"F1"
                                                                          
                                                                                     },
                                                                                     'width':'50px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"L2"
                                                                          
                                                                                     },
                                                                                     'width':'50px',
                                                                           },
                                                                            {
                                                                                 'if':{
                                                                                     'column_id':"F2"
                                                                          
                                                                                     },
                                                                                     'width':'50px',
                                                                           },
                                                                            {
                                                                                 'if':{
                                                                                     'column_id':"BL"
                                                                          
                                                                                     },
                                                                                     'width':'50px',
                                                                           },
                                                                           {
                                                                                 'if':{
                                                                                     'column_id':"QTY"
                                                                          
                                                                                     },
                                                                                     'width':'100px',
                                                                           },
                                                                        ],
                                                                       style_table={'height':'400px'},
                                                                       #export_format='xlsx',
                                                                       #export_headers='display',
                                                                       filter_action= "native",
                                                                       merge_duplicate_headers=True,)])
                                    ],style=CARD4_STYLE),
                                    html.Div([
                                        dbc.Alert("Click 'Calculate' to determine Layer-wise Paper requirement",dismissable=True)]
                                        ,style=CARD4_STYLE),
                                    #html.Div(id='dm_table_alert2'),
                            ]),
                       ])
                ],style=NAVBAR_STYLE3)

demand_table_columns_editable=[
            {'name': 'SKU', 'id': 'sku', 'type': 'text', 'editable':False},
            {'name': 'Paper Type', 'id': 'papertype', 'type': 'text', 'editable':False},
            {'name': 'Length', 'id': 'length', 'type': 'numeric', 'editable':False},
            {'name': 'Width', 'id': 'width_type', 'type': 'numeric', 'editable':False},
            {'name': 'GSM', 'id': 'gsm', 'type': 'numeric', 'editable':False},
            {'name': 'Demand', 'id': 'demand', 'type': 'numeric'}
]
gsm_table_columns_editable=[
            {'name': 'SKU', 'id': 'sku', 'type': 'text', 'editable':False},
            {'name': 'TOP LINER', 'id': 'topliner', 'type': 'text', 'editable':False},
            {'name': 'PLY', 'id': 'ply', 'type': 'numeric', 'editable':False},
            {'name': 'WIDTH', 'id': 'width', 'type': 'numeric', 'editable':False},  
            {'name': 'LENGTH', 'id': 'length', 'type': 'numeric', 'editable':False},                     
            {'name': 'Target GSM', 'id': 'target_gsm', 'type': 'numeric'},
            {'name': 'TL White', 'id': 'tl_l1', 'type': 'numeric', 'editable':False},
            {'name': 'Flutting F1', 'id': 'tl_f1', 'type': 'numeric', 'editable':False},
            {'name': 'L2', 'id': 'tl_l2', 'type': 'numeric', 'editable':False},
            {'name': 'Flutting F2', 'id': 'tl_f2', 'type': 'numeric', 'editable':False},
            {'name': 'L3', 'id': 'tl_l3', 'type': 'numeric', 'editable':False},
            {'name':'Total Optimized GSM','id':'total_optimized_gsm','type': 'numeric', 'editable':False}
]
reel_table_columns_editable=[
            {'name': 'Fluting style', 'id': 'Fluting style', 'type': 'text', 'editable':False},
            {'name': 'Desp', 'id': 'Desp', 'type': 'text', 'editable':False},
            {'name': 'LENGTH(mm)', 'id': 'LENGTH(mm)', 'type': 'numeric', 'editable':False},
            {'name': 'WIDTH(mm)', 'id': 'WIDTH(mm)', 'type': 'numeric', 'editable':False},  
            {'name': 'HEIGHT(mm)', 'id': 'HEIGHT(mm)', 'type': 'numeric', 'editable':False},                     
            {'name': 'TOP PAPER', 'id': 'TOP PAPER', 'type': 'numeric','editable':False},
            {'name': 'BOX TYPE', 'id': 'BOX TYPE', 'type': 'numeric', 'editable':False},
            {'name': 'SHADE', 'id': 'SHADE', 'type': 'numeric', 'editable':False},
            {'name': 'L1', 'id': 'L1', 'type': 'numeric', 'editable':False},
            {'name': 'F1', 'id': 'F1', 'type': 'numeric', 'editable':False},
            {'name': 'L2', 'id': 'L2', 'type': 'numeric', 'editable':False},
            {'name': 'F2', 'id': 'F2', 'type': 'numeric', 'editable':False},
            {'name': 'BL', 'id': 'BL', 'type': 'numeric', 'editable':False},
            {'name':'QTY','id':'QTY','type': 'numeric'}
]

tail_row = dbc.Nav([
              dbc.Row([
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Download Data", id="download_btn",color='#228B22',#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'centre', 'width': '270px'}),#'#D10A2D',
                                dcc.Download(id="download_link"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dcc.Upload(id="upload_btn",
                                           children = html.Div([html.A("Upload File")]),
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center','width': '270px', 'padding-top':'5px'}),#'#D10A2D',
                                ]),
                        html.Div(id='upload_datatable'),
                        dcc.Store(id='stored_data'),
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Edit/save", id="edit_btn",color='#228B22',disabled=False,#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center', 'width': '270px'}),#"#D10A2D",
                                #dcc.Download(id="cdownload-link3"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Submit", id="submit_btn",color='#228B22',disabled=False,#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center','width': '270px' }),#'#"#D10A2D",
                                #dcc.Download(id="cdownload-link3"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),

                  #html.Div(id='dm_container_button'),
                  #html.Br(),
                  html.Div([dcc.ConfirmDialog(id='email_sent_alert',message="The report has been mailed.")])
                  
                
                    ]),
              ],style=NAVBAR_STYLE3)

gsm_tail_row = dbc.Nav([
              dbc.Row([
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Download Data", id="gsm_download_btn",color='#228B22',#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'centre', 'width': '270px'}),#"#D10A2D",
                                dcc.Download(id="gsm_download_link"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dcc.Upload(id="gsm_upload_btn",
                                           children = html.Div([html.A("Upload File")]),
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center','width': '270px', 'padding-top':'5px'}),#"#D10A2D",
                                ]),
                        html.Div(id='gsm_upload_datatable'),
                        
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Edit/save", id="gsm_edit_btn",color='#228B22',disabled=False,#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center', 'width': '270px'}),#"#D10A2D",
                                #dcc.Download(id="cdownload-link3"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Calculate", id="gsm_submit_btn",color='#228B22',disabled=False,#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center','width': '270px' }),#'#"#D10A2D",
                                #dcc.Download(id="cdownload-link3"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),

                  #html.Div(id='dm_container_button'),
                  #html.Br(),
                  html.Div([dcc.ConfirmDialog(id='gsm_email_sent_alert',message="The report has been mailed.")])
                  
                
                    ]),
              ],style=NAVBAR_STYLE3)

reel_tail_row = dbc.Nav([
              dbc.Row([
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Download Data", id="reel_download_btn",color='#228B22',#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'centre', 'width': '270px'}),#"#D10A2D",
                                dcc.Download(id="reel_download_link"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dcc.Upload(id="reel_upload_btn",
                                           children = html.Div([html.A("Upload File")]),
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center','width': '270px', 'padding-top':'5px'}),#"#D10A2D",
                                ]),
                        html.Div(id='reel_upload_datatable'),
                        
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Edit/save", id="reel_edit_btn",color='#228B22',disabled=False,#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center', 'width': '270px'}),#"#D10A2D",
                                #dcc.Download(id="cdownload-link3"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),
                dbc.Col([
                    dbc.Card([
                        html.Div([
                                dbc.Button("Calculate", id="reel_submit_btn",color='#228B22',disabled=False,#"#D10A2D",
                                           style={'color': '#228B22',"font-style":"bold","font-size": "16px", 'font-weight':"700",'text-align': 'center','width': '270px' }),#'#"#D10A2D",
                                #dcc.Download(id="cdownload-link3"),
                                ])
                            ],style=CARD_STYLE01)
                        ]),

                  #html.Div(id='dm_container_button'),
                  #html.Br(),
                  #html.Div([dcc.ConfirmDialog(id='gsm_email_sent_alert',message="The report has been mailed.")])
                  html.Div([dcc.ConfirmDialog(id='reel_email_sent_alert',message="The report has been mailed.")])
                
                    ]),
              ],style=NAVBAR_STYLE3)
output_table_row0 = dbc.Nav([
                            dbc.Row([
                                dbc.Col([
                                    dbc.Card([
                                        html.Div(id='output_table_row')
                                ],style=CARD4_STYLE),
                            ]),
                       ])
                ],style=NAVBAR_STYLE3)

reel_output_table_row = dbc.Nav([
                            dbc.Row([
                                dbc.Col([
                                    dbc.Card([
                                        html.Div(id='reel_output_table_row')
                                ],style=CARD4_STYLE),
                            ]),
                       ])
                ],style=NAVBAR_STYLE3)

loading_state = html.Div(
    id='loading-output',
    children=[
        dcc.Loading(
            id='loading',
            children=[html.Div(id='loading-output-content')],
            type='default'
        ),
     ]
)
gsm_loading_state = html.Div(
    id='gsm_loading-output',
    children=[
        dcc.Loading(
            id='gsm_loading',
            children=[html.Div(id='gsm_loading-output-content')],
            type='default'
        ),
     ]
)
reel_loading_state = html.Div(
    id='reel_loading-output',
    children=[
        dcc.Loading(
            id='reel_loading',
            children=[html.Div(id='reel_loading-output-content')],
            type='default'
        ),
     ]
)
demand_tab_view = html.Div([filter_row,table_row,loading_state,tail_row,html.Br(),output_table_row0,dcc.Store(id='paper_req_json',storage_type = 'memory')])

gsm_tab_view = html.Div([gsm_input_desc,gsm_user_input_row,gsm_table_row,gsm_loading_state,gsm_tail_row,html.Br()])


reel_tab_view = html.Div([reel_assumption_desc,reel_assumption_input_row,reel_table_row,reel_loading_state,reel_tail_row,html.Br(),reel_output_table_row,dcc.Store(id='layerwise_req_json',storage_type = 'memory'),dcc.Store(id='layerwise_top_json',storage_type = 'memory'),dcc.Store(id='layerwise_flute_json',storage_type = 'memory'),dcc.Store(id='layerwise_bottom_json',storage_type = 'memory'),dcc.Store(id='layerwise_middle_json',storage_type='memory')])
#layout= html.Div([location,navbar,row0,html.Br(),html.Br(),html.Br(),html.Div([tab_layout])])
layout= html.Div([location,navbar,html.Br(),html.Br(),html.Br(),tab_layout])
app.layout = layout


def update_table(df,len_val,gsm_val):
        summary_df = df.copy()
        if len_val == 'ALL' and gsm_val == 'ALL':
            return summary_df
        elif len_val == 'ALL' and gsm_val != 'ALL':
            summary_df = summary_df[summary_df['gsm'].isin([gsm_val])]
            summary_df = summary_df.reset_index(drop=True)
            return summary_df
        elif len_val != 'ALL' and gsm_val == 'ALL':
            summary_df = summary_df[summary_df['length'].isin([len_val])]
            summary_df = summary_df.reset_index(drop=True)
            return summary_df
        else:
            summary_df = summary_df[summary_df['length'].isin([len_val])]
            summary_df = summary_df.reset_index(drop=True)
            summary_df = summary_df[summary_df['gsm'].isin([gsm_val])]
            summary_df = summary_df.reset_index(drop=True)
            return summary_df


def parse_contents(contents, filename):
    content_type, content_string = contents.split(',')

    decoded = base64.b64decode(content_string)
    try:
        if 'csv' in filename:
            df = pd.read_csv(
                io.StringIO(decoded.decode('utf-8')))
        elif 'xls' in filename:
            df = pd.read_excel(io.BytesIO(decoded))
    except Exception as e:
        #print(e)
        return no_update, no_update
    return df.to_dict('records'), demand_table_columns_editable

def gsm_parse_contents(contents, filename):
    content_type, content_string = contents.split(',')

    decoded = base64.b64decode(content_string)
    try:
        if 'csv' in filename:
            df = pd.read_csv(
                io.StringIO(decoded.decode('utf-8')))
        elif 'xls' or 'xlsx' in filename:
            df = pd.read_excel(io.BytesIO(decoded))
    except Exception as e:
        #print(e)
        return no_update, no_update
    return df.to_dict('records'), gsm_table_columns_editable

def reel_parse_contents(contents, filename):
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    data_frames = []
    try:
        if 'csv' in filename:
            df = pd.read_csv(io.StringIO(decoded.decode('utf-8')))
            data_frames.append(df)
        elif filename.endswith(('.xls', '.xlsx')):
            xls = pd.ExcelFile(io.BytesIO(decoded))
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                data_frames.append(df)
    except Exception as e:
        # Handle any exceptions here
        return no_update, no_update
    
    return [df.to_dict('records') for df in data_frames], xls.sheet_names

def file_check(master_file,uploaded_file):
    if len(master_file.columns) == len(uploaded_file.columns):# and len(master_file) == len(uploaded_file):
        master_file_columns = master_file.columns
        uploaded_file_columns = uploaded_file.columns
        upload_int_cols = uploaded_file.select_dtypes(include=['int']).columns
        print(master_file.dtypes)
        uploaded_file[upload_int_cols] = uploaded_file[upload_int_cols].astype(float)
        print(uploaded_file.dtypes)
        column_diff = master_file_columns.difference(uploaded_file_columns)
        if len(column_diff) == 0:
            data_type = (master_file.dtypes == uploaded_file.dtypes)
            if (data_type.all() == True):
                return("Targets uploaded successfully")
            else:
                return("Datatypes of the columns in the uploaded file are different")
        else:
            return("The column names in the uploaded file are different")
    else:
        return("The uploaded file does not match with the Targets' file") 

def check_user_gsm_inputs(len_val,gsm_val,trim_val):
    if len_val =='ALL' and gsm_val =='ALL' and trim_val is not None:
        return 'Scroll down to view the results',1
    elif len_val =='ALL' and gsm_val =='ALL' and trim_val is None:
        return 'Please specify trim value before submitting.',0
    elif len_val !='ALL' and gsm_val !='ALL' and trim_val is not None:
        return "Please select 'ALL' for GSM and Length before submitting",0
    elif len_val =='ALL' and gsm_val !='ALL' and trim_val is not None:
        return "Please select 'ALL' for GSM before submitting",0
    elif len_val !='ALL' and gsm_val =='ALL' and trim_val is not None:
        return "Please select 'ALL' for Length before submitting",0



counter = 0
reel_counter = 0

@app.callback(
    Output('demand_table', 'data'),
    Output('demand_table','columns'),
    Output('demand_table','editable'),
    Output('download_link', 'data'),
    Output('output_table_row','children'),
    Output('trim_input_box','value'),
    Output('user_input_alert','displayed'),
    Output('user_input_alert','message'),
    Output('loading-output-content','children'),
    Output('paper_req_json','data'),
    Input('length', 'value'),
    Input('gsm', 'value'), 
    #Input('trim','value'),
    Input('trim_input_box','n_submit'),
    Input('edit_btn','n_clicks'),
    Input('submit_btn','n_clicks'),
    Input("download_btn", "n_clicks"),
    Input('upload_btn','contents'),
    State('upload_btn', 'filename'),
    State('upload_btn','last_modified'),
    State('demand_table', 'data'),
    State('demand_table','columns'),
    State('trim_input_box','value')

)
def input_table(len_val,gsm_val,trim_submit,edit_n_clicks,submit_n_clicks, download_clicks, upload_contents, upload_filename, list_of_dates, table_rows, table_columns,trim_val):
    ctx = dash.callback_context
    summary_df = pd.read_excel(base_path+"data/sample_data.xlsx")
    final_df = update_table(summary_df,len_val,gsm_val)
    data=final_df.to_dict('records')
    columns=demand_table_columns_editable
    if not ctx.triggered:
        return data, columns, True, no_update, no_update,no_update,no_update,no_update,no_update,no_update
    else:
        button_id= ctx.triggered[0]['prop_id'].split('.')[0]
        if trim_submit >= 0:
            #trim_val = int(trim_val)
            pass
        if button_id == 'download_btn':
            final_df = pd.read_excel(base_path+"data/sample_data.xlsx")
            
            if os.path.exists(base_path+"export.xlsx"): 
                    os.remove(base_path+"export.xlsx")

            writer = pd.ExcelWriter(base_path+"export.xlsx", engine='xlsxwriter')
            final_df.to_excel(writer, sheet_name=" demand table", index=False)
            workbook = writer.book
            worksheet = writer.sheets[" demand table"]
            format = workbook.add_format({'bg_color': 'yellow'})
            cols = ['F','G']
            for col in cols:
                #excel_header = str(d[final_df.columns.get_loc(col)])
                excel_header = col
                len_df = str(len(final_df.index) + 1)
                rng = excel_header + '2:' + excel_header + len_df
                worksheet.conditional_format(rng, {'type': 'no_blanks',
                                                'format': format})

            writer.save()
            return no_update, no_update, no_update, dcc.send_file(base_path+"export.xlsx", "demand_table.xlsx"),no_update,trim_val,no_update,no_update,no_update,no_update
        elif button_id == 'upload_btn':
            if upload_contents is not None:
                data, columns = parse_contents(upload_contents, upload_filename)
                upload_df=pd.DataFrame(data)
                new_final_data_master = pd.read_excel(base_path+"data/sample_data.xlsx")
                file_check_status = file_check(new_final_data_master,upload_df)
                if file_check_status == "Targets uploaded successfully":
                    new_final_data_master = new_final_data_master.drop(['demand'], axis=1)
                    upload_df = upload_df.drop(["sku",'papertype','length','width_type','gsm'], axis=1)
                    upload_df = pd.concat([new_final_data_master,upload_df],axis = 1)
                    upload_df.to_excel(base_path+"data/sample_data.xlsx", index=False)
                    final_df = pd.read_excel(base_path+"data/sample_data.xlsx")
                    data = final_df.to_dict('records')
                    columns=demand_table_columns_editable
                    return data, columns, True, no_update,no_update,trim_val,no_update,no_update,no_update,no_update
                else:
                    return no_update, no_update, no_update, no_update,no_update,trim_val,no_update,no_update,no_update,no_update

        elif button_id == 'edit_btn':
            global counter
            if 'edit_btn'!=0:
                counter += 1
                if counter % 2 != 0:
                    data = pd.read_excel(base_path+"data/sample_data.xlsx")
                    data = update_table(data,len_val,gsm_val)
                    data=data.to_dict('records')
                    columns=demand_table_columns_editable
                    return data, columns, True, no_update,no_update,trim_val,no_update,no_update,no_update,no_update
                else:
                    df1 = pd.read_excel(base_path+"data/sample_data.xlsx")
                    df2=pd.DataFrame(table_rows)
                    df1.set_index('sku', inplace=True)
                    df1.update(df2.set_index('sku'))
                    df1 = df1.reset_index()
                    df1.to_excel(base_path+"data/sample_data.xlsx", index=False)
                    df1 = update_table(df1,len_val,gsm_val)
                    data=df1.to_dict('records')
                    columns=demand_table_columns_editable
                    return data, columns, False, no_update,no_update,trim_val,no_update,no_update,no_update,no_update

        elif button_id == 'submit_btn':
            #check if trim value, length and gsm values are entered by the user 
            msg,msg_val = check_user_gsm_inputs(len_val,gsm_val,trim_val)
            print('msg value: ',msg_val)
            if msg_val == 1:
                then = datetime.now()
                then_time = then.strftime("%H:%M:%S")
                print('time 1: ',then_time)
                trim_val = int(trim_val)
                final_df = pd.DataFrame(table_rows)          
                output_df_final,width_value_list,length_value_list,gsm_value_list = main.main_process(final_df,trim_val)
                output_df_final_mod = main.preprocess_output_data(output_df_final,width_value_list,length_value_list,gsm_value_list)
                output_df_final_mod = main.paper_req_calc(output_df_final_mod)
                now = datetime.now()
                now_time = now.strftime("%H:%M:%S")
                print('time 2: ',now_time)
                start  = datetime.strptime(then_time, "%H:%M:%S")
                end  = datetime.strptime(now_time, "%H:%M:%S")   
                runtime = end-start
                print(runtime)
                output_for_store = output_df_final_mod.to_dict('records')
                output_df_final_mod_subset = output_df_final_mod.iloc[:, list(range(9)) + [-3,-2,-1]]
                table_results = html.Div([ 
                                    #dbc.Row([
                                        dt.DataTable(
                                        data=output_df_final_mod_subset.to_dict('records'),
                                        fixed_rows={'headers':True},
                                        sort_action='native',
                                        columns=[{'name': i, 'id': i} for i in output_df_final_mod_subset.columns],
                                        #style_as_list_view=True,
                                        #editable=True,
                                        page_size=8,
                                        style_data={'whiteSpace': 'normal', 'height': 'auto','border': '1px solid rgb(220,220,220)', 'fontWeight': 'bold'},
                                        style_header={'whiteSpace': 'normal','textAlign': 'center','fontWeight': 'bold', "color":"white",
                                                    'height': 'auto','backgroundColor': '#228B22',#'#D10A2D',
                                                    'border': '1px solid rgb(220, 220, 220)' },
                                        style_cell={'textAlign': 'center', 'padding': '0px', 'width': 'auto',
                                                    'overflow': 'hidden','textOverflow': 'ellipsis'},
                                        style_data_conditional=[
                                            {
                                                'if': {'row_index': 'odd'},
                                                'backgroundColor': 'rgb(240, 240, 240)',
                                            },
                                            {
                                                'if': {
                                                    'state': 'active'  # 'active' | 'selected'
                                                },
                                                'backgroundColor': 'rgb(66, 194, 179)',
                                                'border': '1px solid rgb(66, 194, 179)'
                                            },
                                            {
                                                'if':{
                                                    'column_id':"Index"
                                        
                                                    },
                                                    'width':'180px',
                                            },
                                            {
                                                'if':{
                                                    'column_id':"Optimum Reel Size (in mm)"
                                        
                                                    },
                                                    'width':'100px',
                                            },
                                            {
                                                'if':{
                                                    'column_id':"Iteration"
                                        
                                                    },
                                                    'width':'90px',
                                            },
                                            {
                                                'if':{
                                                    'column_id':"Total Demand (in cartons)"
                                        
                                                    },
                                                    'width':'100px',
                                            },
                                            {
                                                'if':{
                                                    'column_id':"Paper Requirement (in tons)"
                                        
                                                    },
                                                    'width':'100px',
                                            },

                                            ],style_table={'height':'448px'}),
                                    #]),
                                dbc.Row([
                                    dbc.Col([
                                        dbc.Button("Download Result", id="output_download_btn",color='#228B22',n_clicks = 0,
                                           style={'color': 'white',"font-size": "15px",'text-align': 'centre', 'width': '150px','height': '50px',"borderColor": '#228B22',"backgroundColor":'#228B22','borderWidth': '1px','borderStyle': 'solid','borderRadius': '20px',"margin-left": '2.5rem'}),
                                           dcc.Download(id="output_download_link"),
                                    ]),
                                    dbc.Col([
        
                                        dcc.Input(id='recipient_box', type='text', placeholder='Enter recipient email id. If there are multiple mail ids, add commas to seperate them.', autoComplete = 'off', style=email_recipient_box_style),
                                    ]),
                                    dbc.Col([
                                        dbc.Button("Email", id="email_data",  style=email_report_style),
                                        
                                    ])
                                        
                            ]),])

                return data, columns, False, no_update,table_results,trim_val,True,msg,html.H4(''),output_for_store
            else:
                return data, columns, False, no_update,no_update,no_update,True,msg,no_update,no_update



    return data, columns, no_update, no_update,no_update,trim_val,no_update,no_update,no_update,no_update

@app.callback(
        Output('gsm_template_link','data'),
        Input('gsm_template_btn','n_clicks')
)
def template_download(n_click):
    ctx = dash.callback_context
    if ctx.triggered:
        #return  dcc.send_file(base_path+"data/user_inputs.csv", "UserInputs.csv")
        return  dcc.send_file(base_path+"data/UserInputs_newformat.csv", "UserInputs.csv")
    else:
        no_update

@app.callback(
        Output('reel_assump_link','data'),
        Input('reel_assump_btn','n_clicks')
)
def reel_assump_template_download(n_click):
    ctx = dash.callback_context
    if ctx.triggered:
        return  dcc.send_file(reel_base_path+"assumptions.xlsx", "assumptions.xlsx")
    else:
        no_update

@app.callback(
    Output('gsm_stored_data_tl_white','data'),
    Output('gsm_stored_data_flute','data'),
    Output('gsm_stored_data_brown_l2_l3','data'),
    Output('gsm_user_upload_msg','displayed'),
    Output('gsm_user_upload_msg','message'),
    Input('upload_gsm_user_input','contents'),
    State('upload_gsm_user_input', 'filename'),
    State('upload_gsm_user_input','last_modified'),
    prevent_initial_callbacks=True
)

def get_gsm_input_val(upload_contents, upload_filename, list_of_dates):
    ctx = dash.callback_context
    if ctx.triggered:
        data, columns = gsm_parse_contents(upload_contents, upload_filename)
        user_input_df=pd.DataFrame(data)
        tl_white = list(user_input_df[user_input_df['TL_White '].notna()]['TL_White '])
        flutting_f1_f2 = list(user_input_df[user_input_df['Fluting_F1_F2'].notna()]['Fluting_F1_F2'])
        #flutting_f2 = list(user_input_df[user_input_df['Fluting_F1_F2'].notna()]['Fluting_F1_F2'])
        tl_brown_l2_l3 = list(user_input_df[user_input_df['TLBrown_L2 _ L3'].notna()]['TLBrown_L2 _ L3'])
        #tl_brown_l2 = list(user_input_df[user_input_df['TLBrown_L2 _ L3'].notna()]['TLBrown_L2 _ L3'])
        #tl_brown_l3 = list(user_input_df[user_input_df['TLBrown_L2 _ L3'].notna()]['TLBrown_L2 _ L3'])
        return tl_white,flutting_f1_f2,tl_brown_l2_l3,True,'Values are updated!!'
    else:
        return no_update,no_update,no_update,no_update,no_update

@app.callback(
    Output('assumption_data_store','data'),
    Output('reel_assump_upload_msg','displayed'),
    Output('reel_assump_upload_msg','message'),
    Input('reel_upload_assump_input','contents'),
    State('reel_upload_assump_input', 'filename'),
    State('reel_upload_assump_input','last_modified'),
    prevent_initial_callbacks=True
)

def get_reel_assumption_val(upload_contents, upload_filename, list_of_dates):
    ctx = dash.callback_context
    if ctx.triggered:
        data, sheet_names = reel_parse_contents(upload_contents, upload_filename)
        with pd.ExcelWriter(reel_base_path + "assumptions.xlsx", engine='xlsxwriter') as writer:
            for sheet_name, sheet_data in zip(sheet_names, data):
                df = pd.DataFrame(sheet_data)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        return data,True,'Values are updated!!'
    else:
        return no_update,no_update,no_update
@app.callback(
    Output('gsm_table', 'data'),
    Output('gsm_table','columns'),
    Output('gsm_table','editable'),
    Output('gsm_download_link', 'data'),
    Output('gsm_loading-output-content','children'),
    Output('gsm_user_input_msg','displayed'),
    Output('gsm_user_input_msg','message'),
    Input('gsm_edit_btn','n_clicks'),
    Input('gsm_submit_btn','n_clicks'),
    Input("gsm_download_btn", "n_clicks"),
    Input('gsm_upload_btn','contents'),
    Input('gsm_stored_data_tl_white','data'),
    Input('gsm_stored_data_flute','data'),
    Input('gsm_stored_data_brown_l2_l3','data'),
    State('gsm_upload_btn', 'filename'),
    State('gsm_upload_btn','last_modified'),
    State('gsm_table', 'data'),
    State('gsm_table','columns'),
)
def input_table(edit_n_clicks,submit_n_clicks, download_clicks, upload_contents, tl_white,flutting, brown_l2_l3,upload_filename, list_of_dates, table_rows, table_columns):
    ctx = dash.callback_context
    #final_df = pd.read_excel(base_path+"data/gsm_data.xlsx")
    final_df = pd.read_excel(base_path+"data/GSM_data_newformat.xlsx")
    data=final_df.to_dict('records')
    columns=gsm_table_columns_editable
    if not ctx.triggered:
        return data, columns, True, no_update, no_update,no_update,no_update
    else:
        button_id= ctx.triggered[0]['prop_id'].split('.')[0]
        if button_id == 'gsm_download_btn':
            #final_df = pd.read_excel(base_path+"data/GSM_data_newformat.xlsx")
            final_df = pd.DataFrame(table_rows)
            if os.path.exists(base_path+"gsm_export.xlsx"): 
                    os.remove(base_path+"gsm_export.xlsx")

            writer = pd.ExcelWriter(base_path+"gsm_export.xlsx", engine='xlsxwriter')
            final_df.to_excel(writer, sheet_name=" gsm table", index=False)
            workbook = writer.book
            worksheet = writer.sheets[" gsm table"]
            format = workbook.add_format({'bg_color': 'yellow'})
            cols = ['F']
            for col in cols:
                #excel_header = str(d[final_df.columns.get_loc(col)])
                excel_header = col
                len_df = str(len(final_df.index) + 1)
                rng = excel_header + '2:' + excel_header + len_df
                worksheet.conditional_format(rng, {'type': 'no_blanks',
                                                'format': format})

            writer.save()
            return no_update, no_update, no_update, dcc.send_file(base_path+"gsm_export.xlsx", "gsm_table.xlsx"),no_update,no_update,no_update
        elif button_id == 'gsm_upload_btn':
            if upload_contents is not None:
                data, columns = gsm_parse_contents(upload_contents, upload_filename)
                upload_df=pd.DataFrame(data)
                #new_final_data_master = pd.read_excel(base_path+"data/GSM_data_newformat.xlsx")
                # file_check_status = file_check(new_final_data_master,upload_df)
                # if file_check_status == "Targets uploaded successfully":
                    #new_final_data_master = new_final_data_master.drop(['target_gsm','optimized_gsm','total_optimized_gsm'], axis=1)
                #new_final_data_master = new_final_data_master.drop(['target_gsm'], axis=1)
                #upload_df = upload_df.drop(["sku","topliner","ply","width","length","tl_l1","tl_f1","tl_l2","tl_f2","tl_l3","total_optimized_gsm"],axis=1)
                #upload_df = pd.concat([new_final_data_master,upload_df],axis = 1)
                upload_df = upload_df[["sku","topliner","ply","width","length","target_gsm","tl_l1","tl_f1","tl_l2","tl_f2","tl_l3","total_optimized_gsm"]]
                upload_df.to_excel(base_path+"data/gsm_data.xlsx", index=False)
                final_df = pd.read_excel(base_path+"data/gsm_data.xlsx")
                data = final_df.to_dict('records')
                columns=gsm_table_columns_editable
                return data, columns, True, no_update,no_update,no_update,no_update
                # else:
                #     return no_update, no_update, no_update, no_update,no_update,no_update,no_update

        elif button_id == 'gsm_edit_btn':
            global counter
            if 'gsm_edit_btn'!=0:
                counter += 1
                if counter % 2 != 0:
                    data = pd.read_excel(base_path+"data/gsm_data.xlsx")
                    data=data.to_dict('records')
                    columns=gsm_table_columns_editable
                    return data, columns, True, no_update,no_update,no_update,no_update
                else:
                    df2=pd.DataFrame(table_rows)
                    df2.to_excel(base_path+"data/gsm_data.xlsx", index=False)
                    data=df2.to_dict('records')
                    columns=gsm_table_columns_editable
                    return data, columns, False, no_update,no_update,no_update,no_update

        elif button_id == 'gsm_submit_btn':
            then = datetime.now()
            then_time = then.strftime("%H:%M:%S")
            print('time 1: ',then_time)
            final_df = pd.DataFrame(table_rows)
            print('Inside submit - DataFrame')
            print(data)
            if tl_white is not None:      
                gsm_output = gsm_main.find_best_match(final_df,tl_white,flutting, brown_l2_l3)
                now = datetime.now()
                now_time = now.strftime("%H:%M:%S")
                print('time 2: ',now_time)
                start  = datetime.strptime(then_time, "%H:%M:%S")
                end  = datetime.strptime(now_time, "%H:%M:%S")   
                runtime = end-start
                print(runtime)
                data=gsm_output.to_dict('records')
                columns=gsm_table_columns_editable
                gsm_output.to_excel(base_path+"data/gsm_data.xlsx", index=False)         
                return data, columns, False, no_update,no_update,False,''
            else:
                return data,columns,False,no_update,no_update,True,'Please upload bucket values for all paper types.'
 
    return data, columns, no_update, no_update,no_update,no_update,no_update


@app.callback(
    Output('reel_table', 'data'),
    Output('reel_table','columns'),
    Output('reel_table','editable'),
    Output('reel_download_link', 'data'),
    Output('reel_output_table_row','children'),
    Output('reel_loading-output-content','children'),
    Output('layerwise_req_json','data'),
    Output('layerwise_top_json','data'),
    Output('layerwise_flute_json','data'),
    Output('layerwise_bottom_json','data'),
    Output('layerwise_middle_json','data'),
    Input('reel_edit_btn','n_clicks'),
    Input('reel_submit_btn','n_clicks'),
    Input("reel_download_btn", "n_clicks"),
    Input('reel_upload_btn','contents'),
    State('reel_upload_btn', 'filename'),
    State('reel_upload_btn','last_modified'),
    State('reel_table', 'data'),
    State('reel_table','columns')

)
def reel_input_table(edit_n_clicks,submit_n_clicks, download_clicks, upload_contents, upload_filename, list_of_dates, table_rows, table_columns):
    ctx = dash.callback_context
    summary_df = pd.read_excel(reel_base_path+"reel_size_input.xlsx")
    data=summary_df.to_dict('records')
    columns=reel_table_columns_editable
    if not ctx.triggered:
        return data, columns, True, no_update, no_update,no_update,no_update, no_update,no_update,no_update,no_update
    else:
        button_id= ctx.triggered[0]['prop_id'].split('.')[0]
        if button_id == 'reel_download_btn':
            final_df = pd.read_excel(reel_base_path+"reel_size_input.xlsx")
            
            if os.path.exists(reel_base_path+"reelsize_table_export.xlsx"): 
                    os.remove(reel_base_path+"reelsize_table_export.xlsx")

            writer = pd.ExcelWriter(reel_base_path+"reelsize_table_export.xlsx", engine='xlsxwriter')
            final_df.to_excel(writer, sheet_name=" Reel size Opt table", index=False)
            workbook = writer.book
            worksheet = writer.sheets[" Reel size Opt table"]
            format = workbook.add_format({'bg_color': 'yellow'})
            cols = ['N']
            for col in cols:
                #excel_header = str(d[final_df.columns.get_loc(col)])
                excel_header = col
                len_df = str(len(final_df.index) + 1)
                rng = excel_header + '2:' + excel_header + len_df
                worksheet.conditional_format(rng, {'type': 'no_blanks',
                                                'format': format})

            writer.save()
            return no_update, no_update, no_update, dcc.send_file(reel_base_path+"reelsize_table_export.xlsx", "reelsize_opt_table.xlsx"),no_update,no_update,no_update, no_update,no_update,no_update,no_update
        elif button_id == 'reel_upload_btn':
            if upload_contents is not None:
                data, columns = parse_contents(upload_contents, upload_filename)
                upload_df=pd.DataFrame(data)
                # new_final_data_master = pd.read_excel(reel_base_path+"reel_size_input.xlsx")
                # file_check_status = file_check(new_final_data_master,upload_df)
                # print('whats the check')
                # print(file_check_status)
                # if file_check_status == "Targets uploaded successfully":
                    # new_final_data_master = new_final_data_master.drop(['QTY'], axis=1)
                    # upload_df = upload_df.drop(["Fluting style","Desp","LENGTH(mm)","WIDTH(mm)","HEIGHT(mm)","TOP PAPER","BOX TYPE","SHADE","L1","F1","L2","F2","BL"], axis=1)
                    # upload_df = pd.concat([new_final_data_master,upload_df],axis = 1)
                upload_df.to_excel(reel_base_path+"reel_size_input.xlsx", index=False)
                final_df = pd.read_excel(reel_base_path+"reel_size_input.xlsx")
                data = final_df.to_dict('records')
                columns=reel_table_columns_editable
                return data, columns, True, no_update,no_update,no_update,no_update, no_update,no_update,no_update,no_update
            else:
                return no_update, no_update, no_update, no_update,no_update,no_update,no_update, no_update,no_update,no_update,no_update

        elif button_id == 'reel_edit_btn':
            global reel_counter
            if 'edit_btn'!=0:
                reel_counter += 1
                if reel_counter % 2 != 0:
                    data = pd.read_excel(reel_base_path+"reel_size_input.xlsx")
                    data=data.to_dict('records')
                    columns=reel_table_columns_editable
                    return data, columns, True, no_update,no_update,no_update,no_update, no_update,no_update,no_update,no_update
                else:
                    df2=pd.DataFrame(table_rows)
                    df2.to_excel(reel_base_path+"reel_size_input.xlsx", index=False)
                    data=df2.to_dict('records')
                    columns=reel_table_columns_editable
                    return data, columns, False, no_update,no_update,no_update,no_update, no_update,no_update,no_update,no_update

        elif button_id == 'reel_submit_btn':
            then = datetime.now()
            then_time = then.strftime("%H:%M:%S")
            final_df = pd.DataFrame(table_rows)          
            output_df_final_mod = reelsize_main.overall_calc(final_df)
            layerwise_top,layerwise_flute,layerwise_bot,layerwise_middle = reelsize_main.layerwise_requirement(output_df_final_mod)
            layerwise_top_store = layerwise_top.to_dict('records')
            layerwise_flute_store = layerwise_flute.to_dict('records')
            layerwise_bottom_store = layerwise_bot.to_dict('records')
            layerwise_middle_store = layerwise_middle.to_dict('records')
            now = datetime.now()
            now_time = now.strftime("%H:%M:%S")
            start  = datetime.strptime(then_time, "%H:%M:%S")
            end  = datetime.strptime(now_time, "%H:%M:%S")   
            runtime = end-start
            print(runtime)
            output_for_store = output_df_final_mod.to_dict('records')
            output_df_final_mod_subset = output_df_final_mod.iloc[:, list(range(2)) + [-14,-13,-5,-4,-3,-2,-1]]
            table_results =  html.Div([ 
                                #dbc.Row([
                                    dt.DataTable(
                                    data=output_df_final_mod_subset.to_dict('records'),
                                    fixed_rows={'headers':True},
                                    sort_action='native',
                                    columns=[{'name': i, 'id': i} for i in output_df_final_mod_subset.columns],
                                    #style_as_list_view=True,
                                    #editable=True,
                                    page_size=8,
                                    style_data={'whiteSpace': 'normal', 'height': 'auto','border': '1px solid rgb(220,220,220)', 'fontWeight': 'bold'},
                                    style_header={'whiteSpace': 'normal','textAlign': 'center','fontWeight': 'bold', "color":"white",
                                                'height': 'auto','backgroundColor': '#228B22',#'#D10A2D',
                                                'border': '1px solid rgb(220, 220, 220)' },
                                    style_cell={'textAlign': 'center', 'padding': '0px', 'width': 'auto',
                                                'overflow': 'hidden','textOverflow': 'ellipsis'},
                                    style_data_conditional=[
                    
                                                            {
                                                                    'if':{
                                                                        'column_id':"Fluting style"
                                                            
                                                                        },
                                                                        'width':'75px',
                                                            },
                                                            {
                                                                    'if':{
                                                                        'column_id':'Desp'
                                                            
                                                                        },
                                                                        'width':'120px',
                                                            },
                                                            {      
                                                                    'if':{
                                                                        'column_id':"L1(in tons)"
                                                            
                                                                        },
                                                                        'width':'100px',
                                                            },
                                                            {
                                                                    'if':{
                                                                        'column_id':"F1(in tons)"
                                                            
                                                                        },
                                                                        'width':'100px',
                                                            },
                                                            {
                                                                    'if':{
                                                                        'column_id':"L2(in tons)"
                                                            
                                                                        },
                                                                        'width':'100px',
                                                            },

                                                            {
                                                                    'if':{
                                                                        'column_id':"F2(in tons)"
                                                            
                                                                        },
                                                                        'width':'100px',
                                                            },
                                                            {
                                                                    'if':{
                                                                        'column_id':"BL"
                                                            
                                                                        },
                                                                        'width':'100px',
                                                            },                        

                                        ],style_table={'height':'440px'}),
                                #]),
                            dbc.Row([
                                dbc.Col([
                                    dbc.Button("Download Result", id="reel_output_download_btn",color='#228B22',n_clicks = 0,
                                        style={'color': 'white',"font-size": "15px",'text-align': 'centre', 'width': '150px','height': '50px',"borderColor": '#228B22',"backgroundColor":'#228B22','borderWidth': '1px','borderStyle': 'solid','borderRadius': '20px',"margin-left": '2.5rem'}),
                                        dcc.Download(id="reel_output_download_link"),
                                ]),
                                dbc.Col([
    
                                    dcc.Input(id='reel_recipient_box', type='text', placeholder='Enter recipient email id. If there are multiple mail ids, add commas to seperate them.', autoComplete = 'off', style=email_recipient_box_style),
                                ]),
                                dbc.Col([
                                    dbc.Button("Email", id="reel_email_data",  style=email_report_style),
                                    
                                ])
                                    
                        ]),])

            return data, columns, False, no_update,table_results,no_update,output_for_store,layerwise_top_store,layerwise_flute_store,layerwise_bottom_store,layerwise_middle_store
        else:
            return data, columns, False, no_update,no_update,no_update,no_update, no_update,no_update,no_update,no_update



    return data, columns, no_update, no_update,no_update,no_update,no_update, no_update,no_update,no_update,no_update


regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
def checkmail(email):
    if(re.fullmatch(regex, email)):
        return "Valid Email"
    else:
        return "Invalid Email"

@app.callback(
    Output('output_download_link','data'),
    Input('paper_req_json','data'),
    Input('output_download_btn','n_clicks'),

    )

def download_output_table(json_data,download_click):
    ctx = dash.callback_context
    button_id= ctx.triggered[0]['prop_id'].split('.')[0]
    if not ctx.triggered:
        return no_update
    else:
        if button_id == 'output_download_btn':
            output_df = pd.DataFrame(json_data)
            output_df.to_excel(base_path+'data/paper_req_output.xlsx',index=False)
            return dcc.send_file(base_path+'data/paper_req_output.xlsx', "Paper_Requirement_data.xlsx")
@app.callback(
    Output('reel_output_download_link','data'),
    Input('layerwise_req_json','data'),
    Input('layerwise_top_json','data'),
    Input('layerwise_flute_json','data'),
    Input('layerwise_bottom_json','data'),
    Input('layerwise_middle_json','data'),
    Input('reel_output_download_btn','n_clicks'),

    )

def download_output_table(overall_json_data,top_json,flute_json,bottom_json,middle_json,download_click):
    ctx = dash.callback_context
    button_id= ctx.triggered[0]['prop_id'].split('.')[0]
    if not ctx.triggered:
        return no_update
    else:
        if button_id == 'reel_output_download_btn':
            output_overall_df = pd.DataFrame(overall_json_data)
            output_top_df = pd.DataFrame(top_json)
            output_flute_df = pd.DataFrame(flute_json)
            output_bottom_df = pd.DataFrame(bottom_json)
            output_middle_df = pd.DataFrame(middle_json)
            with pd.ExcelWriter(reel_base_path+'reel_paper_req_output.xlsx') as writer:
                output_overall_df.to_excel(writer,sheet_name="Overall Calculation",index=False)
                output_top_df.to_excel(writer,sheet_name="Top Layer",index=False)
                output_flute_df.to_excel(writer,sheet_name="Flute Layer",index=False)
                output_bottom_df.to_excel(writer,sheet_name="Bottom Layer",index=False)
                output_middle_df.to_excel(writer,sheet_name="Middle Layer",index=False)
            return dcc.send_file(reel_base_path+'reel_paper_req_output.xlsx', "Reelsize_Paper_Requirement_data.xlsx")


@app.callback(
    Output('email_sent_alert','displayed'),
    Output('email_sent_alert','message'),
    Output('recipient_box','value'),
    Input('paper_req_json','data'),
    Input("email_data", "n_clicks"),
    State('recipient_box', 'value'),
    prevent_initial_callbacks=True
    )
def email_report(json_data,n_clicks,recipient):

    if n_clicks==0 or n_clicks is None:
        return no_update, no_update, no_update
    else:
        if recipient is not None:
            recipient = recipient.rsplit(',')
            all_valid = True
            for recipie in recipient:
                if checkmail(recipie) == "Valid Email":
                    pass
                else:
                    all_valid = False
            if all_valid:
                #keyword_data = pd.DataFrame(pd.read_excel(output_path+query+'_keyword_df.xlsx'))
                #data = keyword_data[['S. No','Keyword Sentences']]
                     
                mail_content = '''Hello,
                The requested file has been shared as an attachment.
                Thank You!
                '''
                #The mail addresses and password
                sender_address = 'testmail.metayb@gmail.com'
                sender_pass = 'htpueabsnqjkhpyr'
                for i in range(0,len(recipient)):
                    recipient[i] = recipient[i].strip()
                    receiver_address = recipient[i]
                    #Setup the MIME
                    message = MIMEMultipart()
                    message['From'] = sender_address
                    message['To'] = receiver_address
                    message['Subject'] = 'Demand Allocation Data'
                    #The subject line
                    #The body and the attachments for the mail
                    message.attach(MIMEText(mail_content, 'plain'))
                    output_df = pd.DataFrame(json_data)
                    output_df.to_excel(base_path+'data/paper_req_output.xlsx',index=False)
                    f = base_path+'data/paper_req_output.xlsx'
                    attachment = MIMEApplication(open(f, "rb").read(), _subtype="txt")
                    attachment.add_header('Content-Disposition','attachment', filename=f.split("/")[-1])
                    message.attach(attachment)
                    # attach_file_name = output_path
                    # attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
                   
                    # payload = MIMEBase('application', 'octate-stream', Name=output_path[8:])
                    # payload.set_payload((attach_file).read())
                    # encoders.encode_base64(payload)
                    # payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
                    # message.attach(payload)
                    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
                    session.starttls() #enable security
                    session.login(sender_address, sender_pass) #login with mail_id and password
                    text = message.as_string()
                    try:
                        session.sendmail(sender_address, receiver_address, text)
                    except:
                        return True, "Enter valid email IDs", ""
                    session.quit()
                # return dbc.Alert("The report has been mailed.")
                return True, "The report has been mailed.", ""
            else:
                return True, "Enter valid email IDs", ""
        else:
            return True, "Enter valid email IDs", ""

@app.callback(
    Output('reel_email_sent_alert','displayed'),
    Output('reel_email_sent_alert','message'),
    Output('reel_recipient_box','value'),
    Input('layerwise_req_json','data'),
    Input('layerwise_top_json','data'),
    Input('layerwise_flute_json','data'),
    Input('layerwise_bottom_json','data'),
    Input('layerwise_middle_json','data'),
    Input("reel_email_data", "n_clicks"),
    State('reel_recipient_box', 'value'),
    prevent_initial_callbacks=True
    )
def email_report(overall_json_data,top_json,flute_json,bottom_json,middle_json,n_clicks,recipient):

    if n_clicks==0 or n_clicks is None:
        return no_update, no_update, no_update
    else:
        if recipient is not None:
            recipient = recipient.rsplit(',')
            all_valid = True
            for recipie in recipient:
                if checkmail(recipie) == "Valid Email":
                    pass
                else:
                    all_valid = False
            if all_valid:
                #keyword_data = pd.DataFrame(pd.read_excel(output_path+query+'_keyword_df.xlsx'))
                #data = keyword_data[['S. No','Keyword Sentences']]
                     
                mail_content = '''Hello,
                The requested file has been shared as an attachment.
                Thank You!
                '''
                #The mail addresses and password
                sender_address = 'testmail.metayb@gmail.com'
                sender_pass = 'htpueabsnqjkhpyr'
                for i in range(0,len(recipient)):
                    recipient[i] = recipient[i].strip()
                    receiver_address = recipient[i]
                    #Setup the MIME
                    message = MIMEMultipart()
                    message['From'] = sender_address
                    message['To'] = receiver_address
                    message['Subject'] = 'Layer-wise Paper Requirement Data'
                    #The subject line
                    #The body and the attachments for the mail
                    message.attach(MIMEText(mail_content, 'plain'))
                    #output_df = pd.DataFrame(json_data)
                    output_overall_df = pd.DataFrame(overall_json_data)
                    output_top_df = pd.DataFrame(top_json)
                    output_flute_df = pd.DataFrame(flute_json)
                    output_bottom_df = pd.DataFrame(bottom_json)
                    output_middle_df = pd.DataFrame(middle_json)
                    with pd.ExcelWriter(reel_base_path+'reel_paper_req_output.xlsx') as writer:
                        output_overall_df.to_excel(writer,sheet_name="Overall Calculation",index=False)
                        output_top_df.to_excel(writer,sheet_name="Top Layer",index=False)
                        output_flute_df.to_excel(writer,sheet_name="Flute Layer",index=False)
                        output_bottom_df.to_excel(writer,sheet_name="Bottom Layer",index=False)
                        output_middle_df.to_excel(writer,sheet_name='Middle Layer',index=False)
                    #output_overall_df.to_excel(reel_base_path+'reel_paper_req_output.xlsx',index=False)
                    f = reel_base_path+'reel_paper_req_output.xlsx'
                    attachment = MIMEApplication(open(f, "rb").read(), _subtype="txt")
                    attachment.add_header('Content-Disposition','attachment', filename=f.split("/")[-1])
                    message.attach(attachment)
                    # attach_file_name = output_path
                    # attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
                   
                    # payload = MIMEBase('application', 'octate-stream', Name=output_path[8:])
                    # payload.set_payload((attach_file).read())
                    # encoders.encode_base64(payload)
                    # payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
                    # message.attach(payload)
                    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
                    session.starttls() #enable security
                    session.login(sender_address, sender_pass) #login with mail_id and password
                    text = message.as_string()
                    try:
                        session.sendmail(sender_address, receiver_address, text)
                    except:
                        return True, "Enter valid email IDs", ""
                    session.quit()
                # return dbc.Alert("The report has been mailed.")
                return True, "The report has been mailed.", ""
            else:
                return True, "Enter valid email IDs", ""
        else:
            return True, "Enter valid email IDs", ""




@app.callback(Output('tab_switch_content', 'children'),
              Input('tab_switch', 'value'))
def render_content(tab):

    # if tab == 'demand_tab':
    #     return demand_tab_view
    if tab == 'reel_tab':
       return reel_tab_view
    elif tab == 'gsm_tab':
       return gsm_tab_view
        
if __name__ == "__main__":
    app.run_server(host='0.0.0.0', port=8050)
    #app.run_server(debug=True)

