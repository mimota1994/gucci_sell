#_*_encoding:utf-8_*_#

import pandas as pd
import numpy as np

Q_table=pd.DataFrame({
    'status':[],
    'A0':[],
    'A1':[],
    'A2':[],
    'A3':[],

})

print(Q_table)

def learn(inial_status,learn_times):
    status=inial_status
    for i in range(learn_times):
        if Q_table['status']!=status:#如果没有此状态，动态添加
            print(Q_table.loc[Q_table.shape[0] + 1])
        # if not Q_table.loc[Q_table['status'] == status]:
        #     print(Q_table.loc[Q_table.shape[0] + 1])

learn([0,0,0],1)
print(Q_table)