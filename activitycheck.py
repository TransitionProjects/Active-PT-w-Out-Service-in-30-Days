__author__ = "David Marienburg"
__version__ = ".1"

"""
This script will, using pandas and the datetime library, check to see if
participants with an open entry into the Support Services Employment Department
have had a service in the last 30 days.  Those who have not will be flagged.
"""

from datetime import datetime
import pandas as pd

def create_activity_report():
    """
    1) Read in the report spreadsheet as two different pandas data frames, one
    per sheet.
    2) Create the today variable.
    3) Slice the services data frame so that it only shows the most recent
    services for each unique participant.
    4) Add today's date to the services data frame.
    5) Add a column showing days between today and service date to the service's
    data frame
    6) Join the two data frames
    7) Fill nan on the days since service column with the difference between the
    entry date and today in days.
    8) Drop non-relevant columns and save the resulting data frame to an excel
    spreadsheet.

    :return: Bool to indicate script success
    """
