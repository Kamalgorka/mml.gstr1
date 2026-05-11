import pandas as pd

def process_sms_report(files):

    loaded = {}

    for key, file in files.items():

        if file is not None:
            loaded[key] = pd.read_excel(file)

    return loaded
