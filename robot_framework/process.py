"""This module contains the main process of the robot."""

import os
import re
import json
from io import BytesIO
from typing import Literal, List

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import concurrent.futures

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from itk_dev_shared_components.graph import mail as graph_mail
from itk_dev_shared_components.graph import authentication as graph_authentication
from itk_dev_shared_components.smtp import smtp_util
from python_serviceplatformen import digital_post
from python_serviceplatformen.authentication import KombitAccess

from robot_framework import config


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    # Prepare access to service platform
    certificate_dir = config.CERTIFICATE_DIR
    kombit_access = KombitAccess(config.ACCESS_CVR, certificate_dir, True)

    # Prepare access to email
    graph_credentials = orchestrator_connection.get_credential(config.GRAPH_API)
    graph_access = graph_authentication.authorize_by_username_password(graph_credentials.username, **json.loads(graph_credentials.password))
    mails = graph_mail.get_emails_from_folder(config.EMAIL_STATUS_SENDER, config.MAIL_SOURCE_FOLDER, graph_access)

    for mail in mails:
        orchestrator_connection.log_trace("Reading emails.")
        # Get attachment from email
        attachments = graph_mail.list_email_attachments(mail, graph_access)
        email_attachment = graph_mail.get_attachment_data(attachments[0], graph_access)
        # Get data from email text
        requester = _get_recipient_from_email(mail.body)
        request_type = _get_request_type_from_email(mail.body)
        # Send and delete email
        return_data = handle_data(email_attachment, kombit_access, request_type)
        _send_status_email(requester, return_data)
        # graph_mail.delete_email(mail, graph_access)


def handle_data(input_file: BytesIO, access: KombitAccess, service_type: Literal['Digital Post', 'Nem SMS', 'Begge']) -> BytesIO:
    """Read data from attachment, lookup each CPR number found and return a new file with added data.

    Args:
        input_file: Excel-file with rows of CPR to work on
        access: Kombit Access Token
        service_type: 'digitalpost', 'nemsms' or 'begge', to set which lookups to perform
    Returns:
        Filtered and formatted file with a list of people indicating whether they have Digital Post or not.
    """
    workbook = load_workbook(input_file)
    input_sheet: Worksheet = workbook.active

    # Check which services are requested
    service = ["Digital Post", "Nem SMS"] if service_type == "Begge" else [service_type]

    # Set column index to place sms as the last cell, even if there is no digital post cell and add column headers
    for s in service:
        input_sheet.cell(row=0, column=input_sheet.max_column + 1, value=s)
    iter_ = iter(input_sheet)

    # Call digital_post.is_registered for each input row and each required service
    data = async_service_check(input_sheet, service, access)

    # Grab workbook from memory and return it
    byte_stream = BytesIO()
    workbook.save(byte_stream)
    byte_stream.seek(0)
    return byte_stream


def async_service_check(input_sheet: Worksheet, service: List[str], kombit_access: KombitAccess) -> Dict[str, List[str]]:
    """
    Call digital_post.is_registered for each input row and each required service.

    Args:
        input_sheet (Worksheet): The input worksheet containing rows of data.
        service (List[str]): A list of services to check registration for.
        kombit_access (KombitAccess): An object providing access credentials for the API.

    Returns:
        Dict[str, List[str]]: A dictionary with CPR as keys and lists of service registration results as values.
    """
    iter_ = iter(input_sheet)
    next(iter_)  # Skip header row
    data = {}

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        all_futures = []
        for row in iter_:
            cpr = row[0].value  # Extract CPR from the row
            for s in service:
                serviceportal_type = s.replace(" ", "").lower()  # Format service name
                # Submit the API call to the thread pool
                future = executor.submit(digital_post.is_registered, cpr=cpr, service=serviceportal_type, kombit_access=kombit_access)
                all_futures.append((cpr, future))

        # Collect results as futures complete
        for cpr, future in concurrent.futures.as_completed([f[1] for f in all_futures]):
            if cpr not in data:
                data[cpr] = []
            data[cpr].append(future.result())  # Append the result to the corresponding CPR entry

    return data


def _write_registered_status(cpr: str, service: Literal["digitalpost", "nemsms"], target_sheet: Worksheet, row: int, column: int, kombit_access: KombitAccess):
    """Check if the CPR is registered for a service and adds a cell to the provided sheet.

    Args:
        cpr: The personal ID to lookup
        service: Which service to lookup (digitalpost or nemsms)
        target_sheet: The excel worksheet to modify
        row, column: Target cell to add
        kombit_access: Access token to use for connection

    """
    is_registered = digital_post.is_registered(cpr=cpr, service=service, kombit_access=kombit_access)
    status = "Tilmeldt" if is_registered else " Ikke tilmeldt"
    target_sheet.cell(row=row, column=column, value=status)


def _get_recipient_from_email(user_data: str) -> str:
    """Find email in user_data using regex."""
    pattern = r"mailto:([^\"]+)"
    return re.findall(pattern, user_data)[0]


def _get_request_type_from_email(user_data: str) -> str:
    """Find request type in user_data using regex."""
    pattern = r"Digital Post eller Nem SMS<br>([^<]+)"
    return re.findall(pattern, user_data)[0].replace(" ", "").lower()


def _send_status_email(recipient: str, file: BytesIO):
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
