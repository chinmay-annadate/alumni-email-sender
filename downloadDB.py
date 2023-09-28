import pandas as pd
import gspread as gs
from pickle import dump

import logging
from datetime import datetime
import traceback


def getSheet():
    # locate service account key
    gc = gs.service_account(
        filename="")

    # open required sheet
    sh = gc.open("").sheet1

    # download into dataframe
    df = pd.DataFrame(sh.get_all_records())

    # retiaining necessary columns
    df = df[['Email Address', 'First Name', 'Last Name', 'Date of Birth']]

    # typecasting dob string into datetime objects
    df['Date of Birth'] = pd.to_datetime(
        df['Date of Birth'], format='%m/%d/%Y', errors='coerce')

    # return dataframe
    return df


def updateCache():
    with open('db.pkl', 'wb') as f:
        dump(getSheet(), f)


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, filename='logs\\db_download_log.txt')
    logging.info(f'{datetime.now().strftime("%d/%m/%Y %H:%M:%S")}')

    try:
        updateCache()
        logging.info('Database updated successfully')

    except:
        # log exceptions
        logging.exception(traceback.format_exc)
