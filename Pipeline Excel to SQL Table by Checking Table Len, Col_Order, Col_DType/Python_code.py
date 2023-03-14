# # Installing Required Libraries If not installed -- Press ( CTRl + ? ) to Uncomment -

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

# Open the workbook and define the worksheet
workbook = openpyxl.load_workbook(path)
worksheet = workbook.active

# Establish a MySQL connection
# Table of MySQL database
print("Enter SQL DBName & TableName -")

print("\nSQL DB Name")
SQL_DB = input()

print("\nSQL Table Name")
SQL_Table = input()


database = MySQLdb.connect (host="localhost", user = "root", passwd = "root", db = SQL_DB)
print("\nSQL connection success")

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

sql_tbl = pd.read_sql_query("select * from " + SQL_Table, database) 
print("Total rows in SQL Table Before adding Rows - ", sql_tbl.shape[0])

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

#Insert Query
query = """INSERT INTO """ + SQL_Table + """ (Date, SM, Product, SP, CP) VALUES (%s, %s, %s, %s, %s)"""

# "INSERT INTO orders_int (Date, SM, Product, SP, CP) VALUES (%s, %s, %s, %s, %s)"

print("-" * 65,"\n")


# Checking Column_length, Column_order, Dtype matching for Excel and SQl if Matching then Dumping Data - 
if excel_tbl_coln_length == sql_column_len:
    
    print("Table Columns Length Matched\n")

    for i in range(0,excel_tbl_coln_length):


        if excel_tbl_coln[i] == sql_column[i]:

            # print('Column_Name by Order Matched', i , "[", excel_tbl_coln[i], " Excel_table ", "=", sql_column[i], " SQL_table ", "]")
            
            if excel_tbl_coln_dtypes[i] == sql_column_dtypes[i]:
                
                # print('Column_Datatype by Order Matched', i , "[", excel_tbl_coln_dtypes[i], " Excel_table ", "=", excel_tbl_coln_dtypes[i], " SQL_table ", "]")
            
                i = i + 1

                if i < excel_tbl_coln_length:

                    continue


                # Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
                # Append each row in the Excel worksheet to the MySQL table

                for row in worksheet.iter_rows(min_row=2, values_only=True):
                    cursor.execute(query, row)

                print("-" * 65)

                sql_tbl_new = pd.read_sql_query("select * from " + SQL_Table, database) 
                #print("SQL Tabel", sql_tbl_new)
                print("Total rows in SQL Table After Adding Rows - ", sql_tbl_new.shape[0])
                # Commit the changes and close the database connection
                database.commit()
                database.close()
                
                print("")
                print("Rows successfully appended to MySQL table\n")
                print("All Conditions Matched - Table Column Length, Column Order, Column DataType")
                
            else:
                    
                print('Column_DataType by Order Not Matched', "Index-" , i ,  "[ Excel Column Name - ", excel_tbl_coln[i], "DataType - ", excel_tbl_coln_dtypes[i], " Excel_table ", " <> ",  " SQL Column Name - ", sql_column[i], "DataType - ", sql_column_dtypes[i],"]")

                break 


            
        else:
            print('Column_Name by Order Not Matched',  "Index-" , i , "[ Excel_table Column Name - ", excel_tbl_coln[i], " <> ", " SQL_Table Column Name - " , sql_column[i], "]")

            break 

else:
    
    print("Table Columns Length not Matched", "[ Excel_Table Length", excel_tbl_coln_length ,"<>", "SQL_Table Length",sql_column_len, "]" )
 
