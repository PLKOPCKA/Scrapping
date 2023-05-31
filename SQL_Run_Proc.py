import pyodbc
import pandas as pd
import glob
import os


def execute_sql_proc(cut, band, iteration):

    if iteration == 0:
        rank = 0
    else:
        rank = 3
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
    if iteration == 0:
        cursor.execute('SET NOCOUNT ON; EXEC [fico].[Insert_Customer_Band];')
        conn.commit()

    cursor.execute(f"SET NOCOUNT ON; EXEC [fico].[Single_Part_Customers_Waves] @cut = {cut}, @band = {band}, @rank = {rank}")
    conn.commit()

    cursor.execute('SET NOCOUNT ON; EXEC [fico].[insFix_LaneProducts];')
    conn.commit()

    os.remove(latest_file)


def insert_outputs(scenario, band, date, obj, gap, time):
    conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                          'Server=ISC-DEV-AS01;'
                          'Database=CUST_ALLOCATION;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()

    cursor.execute('''
                INSERT INTO [CUST_ALLOCATION].[dbo].[CUAL_Outputs]
                (
                [Scenario]
                ,[Band]
                ,[Date_time]
                ,[Solving_time]
                ,[Objective_value]
                ,[Gap_Value])
                VALUES (?,?,?,?,?,?)
                '''
                ,scenario, band, date, time, obj, gap
                )

    conn.commit()


def update_lanes_values(scenario, band):
    conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                          'Server=ISC-DEV-AS01;'
                          'Database=CUST_ALLOCATION;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()

    cursor.execute('''
                    with cte as
                    (
                    SELECT round(sum(cast([Flow Qty] as float)),0) as Total
                    FROM [CUST_ALLOCATION].[dbo].[SCS_TransportationLaneProducts_Results]
                    where [From Facility] = 'Dummy'
                    )

                    update [CUST_ALLOCATION].[dbo].[CUAL_Outputs]
                    set [Lost_Sales_Vol] = (select Total from cte)
                    where Band = ?
                    and Scenario = ?
                '''
                   , band, scenario
                   )

    conn.commit()

    cursor.execute('''
                    with cte as
                    (
                    SELECT round(sum(cast([Product Based Flow Cost] as float)) + round(sum(cast([Flow Qty] as float)),0)*999, 0) as Total
                    FROM [CUST_ALLOCATION].[dbo].[SCS_TransportationLaneProducts_Results]
                    where [From Facility] = 'Dummy'
                    )

                    update [CUST_ALLOCATION].[dbo].[CUAL_Outputs]
                    set [Lost_Sales_Value] = (select Total from cte)
                    where Band = ?
                    and Scenario = ?
                    '''
                   , band, scenario
                   )

    conn.commit()

    cursor.execute('''
                    with cte as
                    (
                    SELECT round(sum(cast([Flow Qty] as float)),0) as Total
                    FROM [CUST_ALLOCATION].[dbo].[SCS_TransportationLaneProducts_Results]
                    where [From Facility] = 'RDC_P022_Oswiecim'
                    )

                    update [CUST_ALLOCATION].[dbo].[CUAL_Outputs]
                    set [Oswiecim_Vol] = (select Total from cte)
                    where Band = ?
                    and Scenario = ?
                    '''
                   , band, scenario
                   )

    conn.commit()


def transform_all():

    conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                          'Server=ISC-DEV-AS01;'
                          'Database=CUST_ALLOCATION;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()
    cursor.execute(f"SET NOCOUNT ON; EXEC [fico].[transform_all]")
    conn.commit()


def update_bands(band_low, band_up):
    conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                          'Server=ISC-DEV-AS01;'
                          'Database=CUST_ALLOCATION;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()

    cursor.execute('''
                    UPDATE [CUST_ALLOCATION].[dbo].[Customer_Band]
                    SET Single = null
                    WHERE Cumulative between ? and ?
                '''
                   , band_low, band_up
                   )
    conn.commit()


def new_bands(iteration, cut, band):
    if iteration == 0:
        rank = 0
    else:
        rank = 3

    conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                          'Server=ISC-DEV-AS01;'
                          'Database=CUST_ALLOCATION;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()

    if iteration == 0:
        cursor.execute('SET NOCOUNT ON; EXEC [fico].[Insert_Customer_Band];')
        conn.commit()

    cursor.execute(f"SET NOCOUNT ON; EXEC [fico].[Single_Part_Customers_Waves] @cut = {cut}, @band = {band}, @rank = {rank}")
    conn.commit()

    cursor.execute('SET NOCOUNT ON; EXEC [fico].[insFix_LaneProducts];')
    conn.commit()
