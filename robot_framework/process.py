"""This module contains the main process of the robot."""

import os
import re
import json
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from itk_dev_shared_components.graph.authentication import GraphAccess
from itk_dev_shared_components.graph import mail as graph_mail
from itk_dev_shared_components.graph import authentication as graph_authentication
from itk_dev_shared_components.graph.mail import Email
from itk_dev_shared_components.smtp import smtp_util
from python_serviceplatformen import digital_post
from python_serviceplatformen.authentication import KombitAccess

from robot_framework import config


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    certificate_dir = "c:\\tmp\\serviceplatformen_test.pem"
    cvr = "55133018"
    kombit_access = KombitAccess(cvr, certificate_dir, True)

    graph_credentials = orchestrator_connection.get_credential(config.GRAPH_API)
    graph_access = graph_authentication.authorize_by_username_password(graph_credentials.username, **json.loads(graph_credentials.password))
    mails = graph_mail.get_emails_from_folder("itk-rpa@mkb.aarhus.dk", config.MAIL_SOURCE_FOLDER, graph_access)

    for mail in mails:
        requester = _get_recipient_from_email(mail.body)
        attachments = mail.list_email_attachments(mail, graph_access)
        email_attachment = mail.get_attachment_data(attachments[0], graph_access)
        return_data = handle_data(email_attachment, kombit_access)
        send_status_email(requester, return_data)


def handle_data(input_file: BytesIO, access: KombitAccess) -> BytesIO:
    """Read data from attachement, lookup each CPR number found and return a new file with added data.

    Returns:
        Filtered and formatted file with a list of people indicating whether they have Digital Post or not.
    """
    workbook = load_workbook(input_file)
    input_sheet: Worksheet = workbook.active

    new_column_index = input_sheet.max_column + 1
    iter_ = iter(input_sheet)
    next(iter_)  # Skip header row
    for row_idx, row in enumerate(iter_, start=2):
        cpr = row[0].value
        is_registered = digital_post.is_registered(cpr=cpr, service="digitalpost", kombit_access=access)
        input_sheet.cell(row=row_idx, column=new_column_index, value=is_registered)

    byte_stream = BytesIO()
    workbook.save(byte_stream)
    byte_stream.seek(0)
    return byte_stream


def _get_recipient_from_email(user_data: str) -> str:
    '''Find email in user_data using regex'''
    pattern = r"E-mail: (\S+)"
    return re.findall(pattern, user_data)[0]


def send_status_email(recipient: str, file: BytesIO):
    """Send an email to the requesting party and to the controller.

    Args:
        email: The email that has been processed.
    """
    smtp_util.send_email(
        recipient,
        config.EMAIL_STATUS_SENDER,
        "RPA: Udtræk om Tilmelding til Digital Post",
        "Robotten har nu udtrukket information om tilmelding til digital post i den forespurgte liste.\n\nVedhæftet denne mail finder du et excel-ark, som indeholder CPR-numre på navngivne borgere, for hvem robotten har slået op i Serviceplatformen og fået svar på, om de er tilmeldt digital post.\n\n Mvh. ITK RPA",
        config.SMTP_SERVER,
        config.SMTP_PORT,
        False,
        [smtp_util.EmailAttachment(file, config.EMAIL_ATTACHMENT)]
    )


if __name__ == '__main__':
    conn_string = os.getenv("OpenOrchestratorConnString")
    crypto_key = os.getenv("OpenOrchestratorKey")
    oc = OrchestratorConnection("Udtræk Tilmelding Digital Post", conn_string, crypto_key, '')
    process(oc)
