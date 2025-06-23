# Udtræk af status på Tilmelding til Digital Post

This robot reads a list of CPR and/or CVR sent from OS2forms and checks status for registration of DigitalPost and/or NemSMS.
The result is sent as an Excel to the requesting email.

## Usage
Enter data via OS2forms: https://selvbetjening.aarhuskommune.dk/da/form/rpa-udtraek-af-tilmelding-til-di
You will receive an email once the robot has finished processing.

## Known errors
If a non-CPR and non-CVR is entered, the robot will enter an error instead of registration value.

## Process Variables
This robot requires the following process variables set in Open Orchestrator:
```
'{"service_cvr":"YOUR_CVR", "thread_count":Number of threads to run}'
```
