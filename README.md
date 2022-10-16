# Mist Scrape
**Author:** Erin Asilo \
**Title:** Infrastructure Network Intern \
**Organization:** RingCentral \
**Purpose:** To automate the process of grabbing certain requested metrics from Mist and put into Google sheets using the sheets API. Currently it grabs the data for client counts of each site daily, as well as successful connect percentage every Friday. 

*Note: This code will not work as is since some files are omitted to protect the privacy of API keys and login credentials.*

In order to run this code a project on cloud.google.com must be created along with a service account and service account credentials. \
Then a *.env* file containing the following is required and must be in *run_files* directory.

```python
EMAIL = 'mistlogin@example.com'

PASSWORD = 'examplepass'

DOWNLOAD_PATH = 'C:\\Users\\Path to run_files directory'

ZIP_PATH = 'C:\\Users\\Path to run_files\<Mist Analytics Zipfile>.zip'

SPREADSHEET_ID = 'ID of target Google Sheet'

SERVICE_ACCOUNT = 'C:\\Users\Path to service account .json'
```

### The code can be ran by using Mist_Scrape.bat 

*To automate the running of this program every day, I use Windows Task Scheduler.*\
*If using task scheduler to run this daily, the task must be set to only run while logged in.*\
*This is because chrome headless mode does not work for manage.mist.com*

## Task Scheduler Instructions for daily automation
```
- Create new basic task
- General
  - Run only when user is logged in
  - Run with highest priveleges
- Triggers
  - Start task on a schedule
  - Daily at 5:00pm (PST) everyday
  - Delay task for up to 1 minute
  - Stop task if it runs longer than 1 hour
- Actions
  - Start a program
  - Program/Script: Mist_Scrape.bat
  - Start in: C:\<Path to repo>\Mist-Scrape\
- Conditions
  - ONLY wake computer to start task
  - AND Start when any network connection is available
- Settings
  - Allow task to be run on demand (for testing)
  - Stop task if it runs longer than 1 hour
  - If running task does not end when requested, force to stop
  - If task is already running, stop the existing instance
```
