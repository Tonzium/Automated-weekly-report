from Connect_SQL_server import ConnectSQL
from SQL_strings import SQL_queries
from PowerPoint import PowerPoint

## Main ##

#Set SQL Query
sql_object = SQL_queries()
SQL_Scrap, SQL_Hold = sql_object.get_variables()

#Connect Oracle
connect_object = ConnectSQL(SQL_Scrap, SQL_Hold)
data_filtered, summary_qty, summary_code, summary_step, summary_product, summary_data, sorted_df, df_top, df_f = connect_object.all_variables()

#Create PowerPoint slide show
pp_object = PowerPoint(data_filtered, summary_qty, summary_code, summary_step, summary_product, summary_data, sorted_df, df_top, df_f)

#Do you want to save Powerpoint?
pp_object.Ask_Save_PP()
