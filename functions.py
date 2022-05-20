import pandas as pd

def update_report(path_to_date_verification, path_to_estimated_completion, transit_days, include_weekends):

    ## Make a dataframe of the requested lines to update
    promise_date = pd.DataFrame(pd.read_excel(
        path_to_date_verification,
        header=1,
        sheet_name='Requested Updates'
        ))

    ## Make a dataframe of the input data
    estimated_date = pd.DataFrame(pd.read_excel(
        path_to_estimated_completion,
        header=0,
        sheet_name=0,
        usecols='A:H'
        )).set_index(['PO', 'Line'])

    ## Look up estimated completion dates from scheduling sheet
    promise_date = promise_date.join(
        estimated_date['Estimated Ship Date'],
        on=['PO #', 'Line #'])

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