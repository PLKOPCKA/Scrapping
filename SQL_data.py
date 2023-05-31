import pyodbc
import pandas as pd
import glob
import os

def execute_sql_proc():


    # list_of_files = glob.glob('C:/Users/PLKOPCKA/Downloads/*.csv')
    # latest_file = max(list_of_files, key=os.path.getctime)
    # latest_file = latest_file.replace('\\', '/')

    latest_file = 'C:/Users/PLKOPCKA/Downloads/Transportation Lane Products Results.csv'

    lane_products = pd.read_csv(latest_file)
    lane_products.columns = lane_products.columns.str.replace(' ', '_')

    conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                          'Server=ISC-DEV-AS01;'
                          'Database=CUST_ALLOCATION;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()
    cursor.execute('SET NOCOUNT ON; TRUNCATE TABLE [CUST_ALLOCATION].[dbo].[SCS_TransportationLaneProducts_Results];')
    conn.commit()

    for row in lane_products.itertuples():
        cursor.execute('''
                    INSERT INTO [CUST_ALLOCATION].[dbo].[SCS_TransportationLaneProducts_Results]
                    ([Scenario], [From Facility],[To Facility],[Product],[Departure Period No],[Arrival Period No]
                    ,[Mode],[Departure Period],[Arrival Period],[Distance],[Travel Time],[Flow Qty]
                     ,[Product Based Flow Cost],[Pipeline Inventory Cost],[Duty Cost],[Tax],[Sourcing Cost],[CO2e Footprint])
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    '''
                    ,row.Scenario, row.From_Facility, row.To_Facility, row.Product, row.Departure_Period_No, row.Arrival_Period_No
                    ,row.Mode, row.Departure_Period, row.Arrival_Period, row.Distance, row.Travel_Time, row.Flow_Qty
                    ,row.Product_Based_Flow_Cost, row.Pipeline_Inventory_Cost, row.Duty_Cost, row.Tax, row.Sourcing_Cost, row.CO2e_Footprint
                    )

    conn.commit()


execute_sql_proc()