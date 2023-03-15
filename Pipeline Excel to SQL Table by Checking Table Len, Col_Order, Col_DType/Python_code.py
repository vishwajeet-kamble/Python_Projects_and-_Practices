# # Installing Required Libraries -- Press ( CTRl + ? ) to Uncomment -

# !pip3 install pymysql --quiet
# !pip3 install ipython-sql --quiet
# !pip3 install mysqlclient --quiet

print("Libraries Installed successfully")

# Imported Libraries

import MySQLdb
import pandas as pd
import openpyxl
# import mysql.connector
import warnings
warnings.simplefilter("ignore", UserWarning)


# Excel Table:
print("Enter_Excel_path -")
path = input()

print("")

# Excel workbook path --
#path = "test_excel.xlsx"

# Open the workbook and define the worksheet
workbook = openpyxl.load_workbook(path)
worksheet = workbook.active

# Table of MySQL database

print("Enter_SQL_Database or Schema_Name -")
schema = input()

print("")

print("Enter_SQL_Table_Name -")
table = input()

print("")


# Establish a MySQL connection
database = MySQLdb.connect (host="localhost", user = "root", passwd = "root", db = schema )

# Get the cursor, which is used to traverse the database, line by line
cursor = database.cursor()


# Defining function to get Table, Column_DType, Column_name, Column_length
def excel_tbl_col_dt_len(rpath): 
    excel_tbl = pd.read_excel(path, index_col=False)

    #print(excel_tbl)

    # print(excel_tbl.dtypes)

    # excel dataframe column
    excel_tbl_coln = excel_tbl.columns
    #print(excel_tbl_coln)

    # excel table column length
    excel_tbl_coln_length = excel_tbl_coln.shape[0]
    #(excel_tbl_coln_length)

    # excel table column datatype
    excel_tbl_coln_dtypes = excel_tbl.dtypes
    #print(excel_tbl_coln_dtypes)
    
    return excel_tbl, excel_tbl_coln_length, excel_tbl_coln, excel_tbl_coln_dtypes

excel_tbl, excel_tbl_coln_length, excel_tbl_coln, excel_tbl_coln_dtypes = excel_tbl_col_dt_len(path)

print("")

# table_name = "orders_int"

sql_tbl = pd.read_sql_query("select * from " + table, database) 

# print(sql_tbl.dtypes)

# sql table columns
sql_column = sql_tbl.columns
# print(sql_column)

#sql table column length
sql_column_len = sql_column.shape[0]
# print(sql_column_len)

#sql table column length
sql_column_dtypes = sql_tbl.dtypes
#print(sql_column_dtypes)

print("SQL connection success")

print("")


print("SQL Table Column Length is - ", sql_column.shape[0] , "\nSo type '%s' with comma seperated", sql_column.shape[0], " times")
print("\nFor example - if Length is 4 then Type like -\n", "%s, %s, %s, %s\n")
value_len = input()

print("")

print("Total rows in SQL Table Before adding Rows - ", sql_tbl.shape[0])

print("")

sql_col_dtype = pd.read_sql_query("select column_name, DATA_TYPE from information_schema.columns where table_schema = '" + schema + "' and table_name = '" + table + "';", database)
# SQL table column data type in sorted order
sql_column_dtype = sql_col_dtype['DATA_TYPE']
sql_column_sort = sql_col_dtype['COLUMN_NAME']

for i in range(0, sql_column_len):
    if sql_column_dtype[i] == "decimal":
        
        sql_column_dtype = [x.replace('decimal', 'float64') for x in sql_column_dtype]
        
    elif sql_column_dtype[i] == "datetime":
        
        sql_column_dtype = [x.replace('datetime', 'datetime64[ns]') for x in sql_column_dtype]
        
    elif sql_column_dtype[i] == "text":
        
        sql_column_dtype = [x.replace('text', 'object') for x in sql_column_dtype]
        
    elif sql_column_dtype[i] == "int":
        
        sql_column_dtype = [x.replace('int', 'int64') for x in sql_column_dtype]
    
    i = i + 1
        
   
# print("sql dtypes", sql_column_dtype)

#Insert Query

query = """INSERT INTO """ + table + """ VALUES (""" +  value_len + """)"""

# "INSERT INTO orders_int (Date, SM, Product, SP, CP) VALUES (%s, %s, %s, %s, %s)"

print("-" * 65)


# Creating DataFrame for Matching SQL & Excel Columns
Col_Match_df = pd.DataFrame({"SQL" :sql_column,
                            "Excel" : excel_tbl_coln}, 
                            columns=['SQL', 'Excel'])

# print(Col_Match_df)
# print(Col_Match_df.shape)

# Column_match_SQL_Excel['Match Status'] Process -
M_S = []
match = 'Matched'
N_match = 'Not-Matched'

for i in range(0, Col_Match_df.shape[0]):
    if Col_Match_df['SQL'][i] == Col_Match_df['Excel'][i]:
        M_S.append(match)
        
    else:
        M_S.append(N_match)

# print(M_S)   
Col_Match_df['Match_Status'] = M_S

# print(Column_match_SQL_Excel)


# Checking Column_length, Column_order, Dtype matching for Excel and SQl if Matching then Dumping Data - 
if excel_tbl_coln_length == sql_column_len:
    
    print("Table Columns Length Matched")

    for i in range(0,excel_tbl_coln_length):


        if excel_tbl_coln[i] == sql_column[i]:

            # print('Column_Name by Order Matched', i , "[", excel_tbl_coln[i], " Excel_table ", "=", sql_column[i], " SQL_table ", "]")
            
            if excel_tbl_coln_dtypes[i] == sql_column_dtype[i]:
                
                # print('Column_Datatype by Order Matched', i , "[", excel_tbl_coln_dtypes[i], " Excel_table ", "=", excel_tbl_coln_dtypes[i], " SQL_table ", "]")
            
                i = i + 1

                if i < excel_tbl_coln_length:

                    continue


                # Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
                # Append each row in the Excel worksheet to the MySQL table

                for row in worksheet.iter_rows(min_row=2, values_only=True):
                    cursor.execute(query, row)

                print("-" * 65)

                sql_tbl_new = pd.read_sql_query("select * from " + table, database) 
                #print("SQL Tabel", sql_tbl_new)
                print("Total rows in SQL Table After Adding Rows - ", sql_tbl_new.shape[0])
                # Commit the changes and close the database connection
                database.commit()
                database.close()
                
                print("")
                print("Rows successfully appended to MySQL table.")
                print("All Conditions Matched - Table Column Length, Column Order, Column DataType")
                
                
            else:
                    
                print('Column_DataType by Order Not Matched', "Index-" , i ,  "[ Excel Column Name - ", excel_tbl_coln[i], "DataType - ", excel_tbl_coln_dtypes[i], " Excel_table ", " <> ",  " SQL Column Name - ", sql_column_sort[i], "DataType - ", sql_column_dtype[i],"]")

                break 


            
        else:
            print('Error - Column_Name by Order Not Matched',  "Index-" , i , "[ Excel_table Column Name - ", excel_tbl_coln[i], " <> ", " SQL_Table Column Name - " , sql_column[i], "]")
            print("")
            print(Col_Match_df)

            break 

else:
    
    print("Table Columns Length not Matched", "[ Excel_Table Length", excel_tbl_coln_length ,"<>", "SQL_Table Length",sql_column_len, "]" )
 
