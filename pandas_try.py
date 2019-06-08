import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np


code_df = pd.read_html('https://oia.yonsei.ac.kr/partner/expReport.asp?page=1&cur_pack=0&ucode=DK000003&bgbn=A', header=0)[0]
print(code_df)

writer = ExcelWriter('department-example.xlsx')
code_df.to_excel(writer,'sheet1',index=False)
writer.save()
