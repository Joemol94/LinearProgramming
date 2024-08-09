import dash
import dash_bootstrap_components as dbc
from flask import Flask

FONT_AWESOME = "https://use.fontawesome.com/releases/v5.10.2/css/all.css"
STYLESHEET_LINK = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css"

meta_tags = [
        {'name': 'viewport',
         'content': 'width=device-width, initial-scale=1.0, minimum-scale = 0.5, maximum-scale=1.2'}
    ]

server = Flask(__name__)
app = dash.Dash(
    __name__,server = server,
     meta_tags= meta_tags,
    external_stylesheets=[dbc.themes.BOOTSTRAP, FONT_AWESOME, STYLESHEET_LINK]
)

app.config.suppress_callback_exceptions = True
app.title = 'Information Summary'

server = app.server
server.config['SECRET_KEY'] = 'mehuhjuysarlmlhankntam'
