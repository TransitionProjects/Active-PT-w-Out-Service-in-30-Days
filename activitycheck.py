__author__ = "David Marienburg"
__version__ = ".1"

"""
This script will, using pandas and the datetime library, check to see if
participants with an open entry into the Support Services Employment Department
have had a service in the last 30 days.  Those who have not will be flagged.
"""

from datetime import datetime
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import pandas as pd

def create_activity_report():
    """
    :return: Bool to indicate script success
    """

    # create the data frames and today variables
    file = askopenfilename(
        title="Open the Employment - 30 Days Since Service Data Report"
    )
    entries_df = pd.read_excel(file, sheet_name="EntryData")
    services_df = pd.read_excel(file, sheet_name="ServiceData")
    today = datetime.today().date

    # sort and slice the services data frame so that it only shows the most
    # recent service
    services_df2 = services_df.sort_values(
        by=["Service Provide Start Date"],
        ascending=False
    ).dropna(
        how="any",
        subset="Service Provide Start Date"
    ).drop_duplicates(
        subset="Client Unique Id",
        keep="first"
    )

    # perform a left join between the two data frames
    joined_df = entries_df.merge(services_df2, how="left", on="Client Unique Id")

    # add today column to the joined data frame
    joined_df["Today"] = today

    # fill the nan in the service provide start date column with the entry date
    joined_df["Service Provide Start Date"].fillna(
        joined_df["Entry Exit Entry Date"],
        inplace=True
    )

    # add days_since_service column to the joined data frame
    joined_df["Days Since Service"] = (
        joined_df["Today"] - joined_df["Service Provide Start Date"]
    ).days

    # create a data frame containing only required output columns
    final = joined_df[[
        "Client Uid",
        "Entry Exit Provider Id",
        "Entry Exit Entry Date",
        "Entry Exit Exit Date",
        "Days Since Service"
    ]]

    # create a data frame that contains rows with Days Since Service values
    # greater than 29
    flagged = final[final["Days Since Service"] > 29]

    # write the data frames to excel
    writer = pd.ExcelWriter(
        asksaveasfilename(
            title="Save the 30 Days Since Services Report",
            defaultextension=".xlsx",
            initialfile="30 Days Since Last Service.xlsx"
        ),
        engine="xlsxwriter"
    )
    flagged.to_excel(writer, sheet_name="No Service in 30 Days", index=False)
    final.to_excel(writer, sheet_name="All Particiapnts", index=False)
    services_df.to_excel(writer, sheet_name="Raw Services", index=False)
    entries_df.to_excel(writer, sheet_name="Raw Entries", index=False)
    writer.save()

    # exit and return true
    return True
