"""This module contains configuration constants used across the framework"""

# The number of times the robot retries on an error before terminating.
MAX_RETRY_COUNT = 3

# Whether the robot should be marked as failed if MAX_RETRY_COUNT is reached.
FAIL_ROBOT_ON_TOO_MANY_ERRORS = True

# Error screenshot config
SMTP_SERVER = "smtp.aarhuskommune.local"
SMTP_PORT = 25
SCREENSHOT_SENDER = "robot@friend.dk"

# Constant/Credential names
ERROR_EMAIL = "Error Email"
KEYVAULT_CREDENTIALS = "Keyvault"
KEYVAULT_URI = "Keyvault URI"
KEYVAULT_PATH = "Digital_Post_Masseopslag"

# Process specific values
GRAPH_API = "Graph API"
MAIL_SOURCE_FOLDER = "Indbakke/Udtræk af Tilmelding til Digital Post"
EMAIL_STATUS_SENDER = "itk-rpa@mkb.aarhus.dk"
EMAIL_USER = "itk-rpa@mkb.aarhus.dk"
EMAIL_ATTACHMENT = "Tilmeldt Digital Post.xlsx"
EMAIL_SUBJECT = "RPA: Udtræk om Tilmelding til Digital Post"
EMAIL_BODY = """"Robotten har nu udtrukket information om tilmelding til digital post for borgere med ukendt adresse.

Vedhæftet denne mail finder du et excel-ark, som indeholder CPR-numre på navngivne borgere, for hvem robotten har slået op i Serviceplatformen og fået svar på, om de er tilmeldt digital post.

Mvh. ITK RPA","""
