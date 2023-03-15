
import pandas as pd


list1 = [1, 2, 3, 4, 5]

list2 = [1, 2, 6, 4, 5]


Column_match_SQL_Excel = pd.DataFrame({"SQL" :list1,
                                "Excel" : list2}, 
                                columns=['SQL', 'Excel'])

# print(Column_match_SQL_Excel)
# print(Column_match_SQL_Excel.shape)

# Column_match_SQL_Excel['Match Status'] = 
M_S = []
match = 'Matched'
N_match = 'Not-Matched'

for i in range(0, Column_match_SQL_Excel.shape[0]):
    if Column_match_SQL_Excel['SQL'][i] == Column_match_SQL_Excel['Excel'][i]:
        M_S.append(match)
        
    else:
        M_S.append(N_match)

#print(M_S)
        
Column_match_SQL_Excel['Match_Status'] = M_S

print(Column_match_SQL_Excel)