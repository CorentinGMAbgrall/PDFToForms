# Import all path
import os.path

# Setup PDF scrapping and dataframe tools
import pdfplumber as pp
import pandas as pd
import json
from tabulate import tabulate as tb

# Setup Google OAuth and API
from apiclient import discovery
from httplib2 import Http
from oauth2client import client, file, tools

SCOPES = "https://www.googleapis.com/auth/forms.body"
DISCOVERY_DOC = "https://forms.googleapis.com/$discovery/rest?version=v1"

store = file.Storage("token.json")
creds = None
if not creds or creds.invalid:
  flow = client.flow_from_clientsecrets("credentials.json", SCOPES)
  creds = tools.run_flow(flow, store)

form_service = discovery.build(
    "forms",
    "v1",
    http=creds.authorize(Http()),
    discoveryServiceUrl=DISCOVERY_DOC,
    static_discovery=False,
)

# Extract data from pdf
with pp.open("./questionnaire.pdf") as pdf:
    page = pdf.pages[5]
    tables = page.extract_tables()
    table = tables[0]

    df = pd.DataFrame(table)
    df.at[0, 0] = "Question"

# Create a new form
# ToDo: ask user input?
NEW_FORM = {
    "info": {
        "title": "First table test",
        "documentTitle": "First try"
    }
}

# Create a test question in the form
NEW_QUESTION = {
    "requests": [
        {
            "createItem": {
                "item": {
                    "title": "B2. Es-tu satisfait:",
                    "questionGroupItem": {
                        "grid": {
                            "columns": {
                                "type": "RADIO",
                                "options": [
                                    { "value": "Pas du tout satisfait" },
                                    { "value": "Plutôt pas satisfait" },
                                    { "value": "Plutôt satisfait" },
                                    { "value": "Tout à fait satisfait" }
                                ]
                            }
                        },
                        "questions": []
                    }
                },
                "location": { "index": 0 }
            }
        }
    ]
}
# level = item
# "location": { "index": 0 }

# Insert questions extracted from the table into NEW_QUESTION
qt = df[0].to_list()
questions = qt[:0] + qt[1:]

qt_list = []
for i in range(0, len(questions)):
    sub_dict = { "title": questions[i] }
    dict = { "rowQuestion": sub_dict }
    qt_list.append(dict)

NEW_QUESTION["requests"][0]["createItem"]["item"]["questionGroupItem"]["questions"] = qt_list

# Insert levels extracted from the table into NEW_QUESTION
qt = df.iloc[0].to_list()
qt = [sub.replace('\n', ' ') for sub in qt]
levels = qt[:0] + qt[1:]

qt_list = []
for i in range(0, len(levels)):
    dict = { "value": levels[i] }
    qt_list.append(dict)

NEW_QUESTION["requests"][0]["createItem"]["item"]["questionGroupItem"]["grid"]["columns"]["options"] = qt_list

# Creates the initial form
result = form_service.forms().create(body=NEW_FORM).execute()

# Adds the question to the form
question_setting = (
    form_service.forms()
    .batchUpdate(formId=result["formId"], body=NEW_QUESTION)
    .execute()
)

# Prints the result to show the question has been added
get_result = form_service.forms().get(formId=result["formId"]).execute()
print(get_result)
