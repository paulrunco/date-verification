import pandas as pd


path_to_estimated_completion_dates = 'sandbox/SpaceX estimated completion 4-27-22.xlsx'
path_to_promise_date_verification = 'sandbox/Empty_Promise_Date_Verification_042422 - Copy.xlsx'

## Make a dataframe of the requested lines to update
promise_date = pd.DataFrame(pd.read_excel(
    path_to_promise_date_verification,
    header=1,
    sheet_name='Requested Updates'
    ))

## Make a dataframe of the input data
estimated_date = pd.DataFrame(pd.read_excel(
    path_to_estimated_completion_dates,
    header=0,
    sheet_name=0,
    usecols='A:H'
    )).set_index(['PO', 'Line'])

## Look up estimated completion dates from scheduling sheet
promise_date = promise_date.join(
    estimated_date['Estimated Ship Date'],
    on=['PO #', 'Line #'])

# Update promise dates
promise_date['Verified Current Promise Date'] = promise_date['Estimated Ship Date'] + pd.tseries.offsets.BusinessDay(3)
print(promise_date)




# ## Write promise dates into the Promise Date Verification template
# with pd.ExcelWriter(
#     path_to_promise_date_verification,
#     mode='a',
#     if_sheet_exists='overlay',
#     date_format='yyyy-mm-dd',
#     datetime_format='yyyy-mm-dd'
# ) as writer:
#     test.to_excel(
#         writer,
#         sheet_name='Requested Updates',
#         startrow=2,
#         startcol=7,
#         header=False,
#         index=False
#     )




## Look up each order-line in ETC data and...
#   if exists:
#
#       if new date != old date:
#           update date with new date + 3 days
#   else:
#       flag for user review


## Format and write output data to sheet



## Manage settings file
# load settings file if exists, otherwise create new
