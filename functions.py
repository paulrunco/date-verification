import datetime as dt
from re import sub
import pandas as pd

def update_report(path_to_date_verification, path_to_estimated_completion, transit_days, include_weekends):

    ## Make a dataframe of the requested lines to update
    promise_date = pd.DataFrame(pd.read_excel(
        path_to_date_verification,
        header=1,
        sheet_name='Requested Updates'
        ))

    # Create an index column using order-release
    promise_date['Order-Line'] = promise_date['PO #'].astype(str) + '-' + promise_date['Line #'].astype(str)
    promise_date = promise_date.set_index(['Order-Line'])

    ## Make a dataframe of the input data
    estimated_date = pd.DataFrame(pd.read_excel(
        path_to_estimated_completion,
        header=0,
        sheet_name=0,
        usecols='A:H'
        ))

    # Create an index column using order-release
    estimated_date['Order-Line'] = estimated_date['PO'].astype(str) + '-' + estimated_date['Line'].astype(str)
    estimated_date = estimated_date.set_index(['Order-Line'])

    ## Find and fix Excel-formatted dates 

    if estimated_date['Estimated Ship Date'].dtype == 'object':
        estimated_date = estimated_date.sort_index()
        for i, row in estimated_date[estimated_date['Estimated Ship Date'].astype(str).str.match('\d{5}')].iterrows():
            excel_date = row['Estimated Ship Date']
            try:
                fixed_date = pd.to_datetime(excel_date, unit='d', origin='1899-12-30')
                estimated_date.loc[i, 'Estimated Ship Date'] = fixed_date
            except:
                print('Dates cannot be automatically converted. Rows with malformed dates will be ignored.')

    ## Remove any malformed dates
    estimated_date['Estimated Ship Date'] = pd.to_datetime(estimated_date['Estimated Ship Date'], errors='coerce')
    estimated_date.dropna(inplace=True, subset=['Estimated Ship Date'])

    ## Look up estimated completion dates from scheduling sheet
    promise_date = promise_date.join(
        estimated_date['Estimated Ship Date'],
        on=['Order-Line'])

    # Update promise dates
    if include_weekends:
        promise_date['Verified Current Promise Date'] = promise_date['Estimated Ship Date'] + pd.tseries.offsets.DateOffset(transit_days)
    else:
        promise_date['Verified Current Promise Date'] = promise_date['Estimated Ship Date'] + pd.tseries.offsets.BusinessDay(transit_days)

    promise_date.drop_duplicates(['PO #', 'Line #'], keep="first", inplace=True)
    promise_date['Verified Current Promise Date'] = promise_date['Verified Current Promise Date'].dt.date

    ## Write promise dates into the Promise Date Verification template
    with pd.ExcelWriter(
        path_to_date_verification,
        mode="a",
        if_sheet_exists="overlay",
        date_format="YYYY-MM-DD",
        datetime_format="YYYY-MM-DD"
    ) as writer:
        promise_date['Verified Current Promise Date'].to_excel(
            writer,
            sheet_name="Requested Updates",
            startrow=2,
            startcol=7,
            header=False,
            index=False
        )