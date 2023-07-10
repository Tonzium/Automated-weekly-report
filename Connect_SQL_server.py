import cx_Oracle
from dotenv import load_dotenv
import pandas as pd
from datetime import date, datetime, timedelta
import os



class ConnectSQL:
    def __init__(self, SQL_Scrap, SQL_Hold):
        self.SQL_Scrap = SQL_Scrap
        self.SQL_Hold = SQL_Hold
        self.today = self.__Current_Time()
        self.number_days = self.__Ask_scrap_days()
        self.ask_hold = self.__Ask_hold()
        self.start_datetime = self.__Set_time_variables()
        self.user_id_pdam, self.password_id_pdam, self.dsn_id_pdam = self.__Client_info()
        self.data_filtered, self.summary_qty, self.summary_code, self.summary_step, self.summary_product, self.summary_data, self.sorted_df = self.Connect_scrap()
        self.df_top, self.df_f = self.Connect_hold()
        

    def __Client_info(self):
        load_dotenv("dotenv.env")
        user_id_pdam = os.getenv('USER_pdam')
        password_id_pdam = os.getenv('PW_pdam')
        dsn_id_pdam = os.getenv('DSN_pdam')
        return user_id_pdam, password_id_pdam, dsn_id_pdam

    def __SQL_query(self):
        load_dotenv()
        Scrap_query = 'Scrap_SQL_query'
        Hold_query = 'HOLD_SQL_query'
        return Scrap_query, Hold_query


    def __Current_Time(self):
        #Print current time
        today = date.today()
        string_time = today.strftime("%d/%m/%Y")
        print("\n\nCurrent date:", string_time, "\n")
        return today

    def __Ask_scrap_days(self):
        while True:
            try:
                #Loop
                number_days = input("How many days of scrap data you want [4-14]:")
                number_days = int(number_days)
                if number_days < 4 or number_days > 14:
                    raise ValueError("Number of days must be between 4 and 14.")
                #Exit loop if this is valid
                break
            except ValueError:
                #Print error message
                print("Invalid input. Please enter an integer between 4 and 14.")
        return number_days

    def __Ask_hold(self):
        ask_hold = input("Do you want to include HOLD's to your report Y/N:")
        if ask_hold.lower() == "yes" or ask_hold.lower() == "y":
            ask_hold = 1
        else:
            ask_hold = 0
        return ask_hold
    
    def __Set_time_variables(self):
        #Filter times beyond X days from SQL data
        start_date = self.today - timedelta(days=self.number_days) #this row cannot be re-positioned
        #Converting 'date' -datatype to datetime
        start_datetime = datetime.combine(start_date, datetime.min.time())
        return start_datetime


    def Connect_scrap(self):
        try:
            #Creating connection to Oracle database
            connection = cx_Oracle.connect(user=self.user_id_pdam, password=self.password_id_pdam, dsn=self.dsn_id_pdam)
                
            #Creating Cursor Object
            cur = connection.cursor()
            print("**Connection to Oracle Successful**")
                
            #Execute SQL
            cur.execute(self.SQL_Scrap)
            
            #Fetch all rows
            rows = cur.fetchall()
            
            #Filter main DataFrame
            df = pd.DataFrame(rows, columns=[desc[0] for desc in cur.description])
            data = df.fillna("NO DATA")
            data_filtered = data.loc[data["ACTION_DATE"] >= self.start_datetime]
            
            
            # DEVICE vs Scrap data
            grouped = data_filtered.groupby("DEVICE_ID")
            summary_qty = grouped['QTY'].sum().sort_values(ascending=False)
            
            # CODE vs Scrap data + add additional column CODES with real names of codes
            data_filtered["CODES"] = data_filtered["CODE"].replace("GSC_02", "x").replace("GSC_03",
            "x").replace("GSC_04", "x").replace("GSC_13","x").replace("GSC_15", "x").replace("GSC_17",
            "x").replace("GSC_20","x").replace("SC_04",
            "x").replace("SC_09", "x").replace("SC_10", "x").replace("GSC_11",
            "x").replace("GSC_16", "x").replace("GSC_10",
            "x").replace("HSC_01","x").replace("HSC_02",
            "x").replace("POKECS_06","x")
            
            grouped_code = data_filtered.groupby("CODES")
            summary_code = grouped_code.QTY.sum().sort_values(ascending=False)
            
            #aggregated df for plotting
            aggregated_df = data_filtered.groupby("CODES", as_index=False).agg({"QTY": "sum"})
            #sort the aggregated values for barplot
            sorted_df = aggregated_df.sort_values(by="QTY", ascending=False)
            
            # CURRENT_STEP data
            grouped_step = data_filtered.groupby("CURRENT_STEP")
            summary_step = grouped_step.QTY.sum().sort_values(ascending=False)
            
            # PRODUCT_ID data
            grouped_product = data_filtered.groupby("PRODUCT_ID")
            summary_product = grouped_product.QTY.sum().sort_values(ascending=False)

            #data filtered
            summary_data = data_filtered[["LOT_ID", "QTY", "COMMENTS", "CODE"]]

            #Print example of this dataframe
            print("\nTotal rows in query:", len(data_filtered))
            print("\nLOT_ID // QTY // COMMENTS")
            #show scraps in command line
            for index, row in data_filtered.iterrows():
                print(index ,row["LOT_ID"], "//", row["QTY"], "//", row["COMMENTS"])

            print("\nTotal amount of scraps:", data_filtered["QTY"].sum())
            
            #close object
            cur.close()

        #Print if Oracle connection success/fail
        except cx_Oracle.Error as error:
            print("\nFailed to connect,", error, "\n")
        finally:
            if connection:
                connection.close()
                print("\n**The Oracle connection is closed **\n")
        
        return data_filtered, summary_qty, summary_code, summary_step, summary_product, summary_data, sorted_df


    def Connect_hold(self):
        if self.ask_hold == 1:
            #Creating connection to Oracle database
            connection = cx_Oracle.connect(user=self.user_id_pdam, password=self.password_id_pdam, dsn=self.dsn_id_pdam)
            
            cur = connection.cursor()

            #Execute HOLD SQL
            cur.execute(self.SQL_Hold)
            
            #Fetch all rows
            rows2 = cur.fetchall()
            
            #Filter main DataFrame
            df_hold_dataframe = pd.DataFrame(rows2, columns=[desc[0] for desc in cur.description])
            df12 = df_hold_dataframe.fillna("NO DATA")
            df13 = df12.loc[df12["LOCATION_ID"] == "LASITUS"]
            df15 = df13.loc[df13["HOLD_DATE"] >= self.start_datetime]
            
            #HOLD Data STEP_ID
            df_hold_grouped = df15.groupby("STEP_ID")
            df_hold = df_hold_grouped.LOT_ID.size().sort_values(ascending=False)
            df_top10 = df_hold.nlargest(15)
            df_top = df_top10.reset_index()
            
            #Freq data by COMMENTS
            df_freq = df15.groupby("COMMENTS")
            df_f = df_freq.LOT_ID.size().sort_values(ascending=False).reset_index()
            
            cur.close()
        else:
            print("\n**HOLD data not included**\n")

        return df_top, df_f
    
    def all_variables(self):
        return self.data_filtered, self.summary_qty, self.summary_code, self.summary_step, self.summary_product, self.summary_data, self.sorted_df, self.df_top, self.df_f