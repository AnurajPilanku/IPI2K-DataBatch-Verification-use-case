#Anuraj Pilanku
#Filter rows based on column values
import pandas as pd
import sys
xl=sys.argv[1]
no=sys.argv[2]
yes=sys.argv[3]
columnname='Confirmed'


data=pd.read_excel(xl,sheet_name=0,engine='openpyxl',index=False)
no_df=data.loc[data[columnname].isin(['N'])]
yes_df=data.loc[data[columnname].isin(['Y'])]
no_df.to_excel(no,index=False)
yes_df.to_excel(yes,index=False)
print('success')