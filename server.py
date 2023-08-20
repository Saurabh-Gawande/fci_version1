import pandas as pd
from pulp import *
from array import *
import json
from flask import Flask, request, session, jsonify, send_file
import pickle
from flask_cors import CORS
import ast



app = Flask(__name__)
app.secret_key = 'aqswdefrgt'
CORS(app, supports_credentials=True)


@app.route('/login',methods = ["POST"])
def login():
    username = request.form['username']
    password = request.form['password']
    data = {}
    if(username=="admin@iitd.com" and password=="admin@321"):
        data['status'] = 1
        session['username'] = username
    else:
        data['status'] = 0

    json_data = json.dumps(data)
    json_object = json.loads(json_data)
    response = jsonify(json_object)
    response.headers.add('Access-Control-Allow-Origin', 'http://localhost:3000')
    response.headers.add('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type, Authorization')
    response.headers.add('Access-Control-Allow-Credentials', 'true') 
    return (json.dumps(json_object, indent = 1))

# @app.route("/upload_Monthly_File_M01",methods = ["POST"])
# def upload_Monthly_File_M01():
#     data = {}
#     try:
#         file = request.files['uploadFile']
#         file.save("Input//Input_Template.xlsx")
#         data['status'] = 1
#     except:
#         data['status'] = 0
    
#     json_data = json.dumps(data)
#     json_object = json.loads(json_data)

#     return(json.dumps(json_object, indent = 1))

# @app.route("/upload_Monthly_File_M02",methods = ["POST"])
# def upload_Monthly_File_M02():
#     data = {}
#     try:
#         file = request.files['uploadFile_M02']
#         file.save("Input//Input_Template_M02.xlsx")
#         data['status'] = 1
#     except:
#         data['status'] = 0
    
#     json_data = json.dumps(data)
#     json_object = json.loads(json_data)

#     return(json.dumps(json_object, indent = 1))

# @app.route("/uploadDailyFile_S2",methods = ["POST"])
# def uploadDailyFile_S2():
#     data = {}
#     try:
#         file = request.files['uploadFile']
#         file.save("Input//Temp_balanced_DPT_scen2.xlsx")
#         data['status'] = 1
#     except:
#         data['status'] = 0
    
#     json_data = json.dumps(data)
#     json_object = json.loads(json_data)

#     return(json.dumps(json_object, indent = 1))

# @app.route("/uploadDailyFile_S1",methods = ["POST"])
# def uploadDailyFile_S1():
#     data = {}
#     try:
#         file = request.files['uploadFile']
#         file.save("Input//Temp_balanced_DPT_scen1.xlsx")
#         data['status'] = 1
#     except:
#         data['status'] = 0
    
#     json_data = json.dumps(data)
#     json_object = json.loads(json_data)

#     return(json.dumps(json_object, indent = 1))
    

@app.route("/read_Monthly_state_table",methods = ["POST","GET"])
def read_Monthly_state_table():
    if request.method == "POST":        
        try: 
            df1 = pd.read_excel('Output\\Monthly_State_To_State_Table.xlsx', sheet_name="Wheat")
            df2 = pd.read_excel('Output\\Monthly_State_To_State_Tablee.xlsx', sheet_name="Rice")    
            json_data1 = df1.to_json(orient='records', indent=1)
            json_data2 = df2.to_json(orient='records', indent=1)
            json_data = {"Wheat": json_data1, "Rice": json_data2}
        except:
            json_data = json.dumps({"Status": 0}, indent=1)

        json_object = json.dumps(json_data)
        return json_object
    else:
        return ("Error")
        

@app.route("/read_Relevant_Result",methods = ["POST","GET"])
def read_Relevant_Result():
    if request.method == "POST":        
        try: 
            df1 = pd.read_excel('Output\\Relevent_Results.xlsx', sheet_name="wheat")
            df2 = pd.read_excel('Output\\Relevent_Results.xlsx', sheet_name="rice")    
            json_data1 = df1.to_json(orient='records', indent=1)
            json_data2 = df2.to_json(orient='records', indent=1)
            json_data = {"wheat": json_data1, "rice": json_data2}
        except:
            json_data = json.dumps({"Status": 0}, indent=1)

        json_object = json.dumps(json_data)
        return json_object
    else:
        return ("error")
    
# @app.route("/read_Result_M02",methods = ["POST","GET"])
# def read_Result_M02():
#     if request.method == "POST":        
#         try: 
#             df1 = pd.read_excel('Output\\Results_tentative.xlsx', sheet_name="r_wheat")
#             df2 = pd.read_excel('Output\\Results_tentative.xlsx', sheet_name="r_rice")  
#             df3 = pd.read_excel('Output\\Results_tentative.xlsx', sheet_name="w_wheat")
#             df4 = pd.read_excel('Output\\Results_tentative.xlsx', sheet_name="w_rice")  
#             df5 = pd.read_excel('Output\\Results_tentative.xlsx', sheet_name="x_wr_wheat")
#             df6 = pd.read_excel('Output\\Results_tentative.xlsx', sheet_name="x_wr_rice")  
#             df7 = pd.read_excel('Output\\Results_tentative.xlsx', sheet_name="x_rw_wheat")
#             df8 = pd.read_excel('Output\\Results_tentative.xlsx', sheet_name="x_rw_rice")    
#             json_data1 = df1.to_json(orient='records', indent=1)
#             json_data2 = df2.to_json(orient='records', indent=1)
#             json_data3 = df3.to_json(orient='records', indent=1)
#             json_data4 = df4.to_json(orient='records', indent=1)
#             json_data5 = df5.to_json(orient='records', indent=1)
#             json_data6 = df6.to_json(orient='records', indent=1)
#             json_data7 = df7.to_json(orient='records', indent=1)
#             json_data8 = df8.to_json(orient='records', indent=1)
#             json_data = {"r_rice": json_data1, "r_wheat": json_data2, "w_rice": json_data3, "w_wheat": json_data4, "x_wr_rice": json_data5, "x_wr_wheat": json_data6, "x_rw_rice": json_data7, "x_rw_wheat": json_data8}
#         except:
#             json_data = json.dumps({"Status": 0}, indent=1)

#         json_object = json.dumps(json_data)
#         return json_object
#     else:
#         return ("error")
    
@app.route("/read_Daily_Planner_S2",methods = ["POST","GET"])
def read_Daily_Planner_S2():
    if request.method == "POST":        
        try: 
            df1 = pd.read_excel('Output\\List_DPT2.xlsx', sheet_name="rice") 
            df2 = pd.read_excel('Output\\List_DPT2.xlsx', sheet_name="wheat") 
            json_data1 = df1.to_json(orient='records', indent=1)
            json_data2 = df2.to_json(orient='records', indent=1)
            json_data = {"rice": json_data1, "wheat": json_data2}
        except:
            json_data = json.dumps({"Status": 0}, indent=1)

        json_object = json.dumps(json_data)
        return json_object
    else:
        return ("error")

@app.route("/read_Daily_Planner_S1",methods = ["POST","GET"])
def read_Daily_Planner_S1():
    if request.method == "POST":        
        try: 
            df1 = pd.read_excel('Output\\List_DPT.xlsx', sheet_name="rice") 
            df2 = pd.read_excel('Output\\List_DPT.xlsx', sheet_name="wheat") 
            json_data1 = df1.to_json(orient='records', indent=1)
            json_data2 = df2.to_json(orient='records', indent=1)
            json_data = {"rice": json_data1, "wheat": json_data2}
        except:
            json_data = json.dumps({"Status": 0}, indent=1)

        json_object = json.dumps(json_data)
        return json_object
    else:
        return ("error")
    
@app.route("/read_Monthly_Template_M1",methods = ["POST","GET"])
def read_Monthly_Template_M1():
    if request.method == "POST":        
        try: 
            df1 = pd.read_excel('Input\\Monthly_Template_M1.xlsx', sheet_name="Surplus_wheat") 
            df2 = pd.read_excel('Input\\Monthly_Template_M1.xlsx', sheet_name="Deficit_wheat")
            df3 = pd.read_excel('Input\\Monthly_Template_M1.xlsx', sheet_name="Surplus_rice")
            df4 = pd.read_excel('Input\\Monthly_Template_M1.xlsx', sheet_name="Deficit_rice")
            df5 = pd.read_excel('Input\\Monthly_Template_M1.xlsx', sheet_name="States_supply")
            df6 = pd.read_excel('Input\\Monthly_Template_M1.xlsx', sheet_name="States_allocation")
            df7 = pd.read_excel('Input\\Monthly_Template_M1.xlsx', sheet_name="Rail_cost_chart")
            json_data1 = df1.to_json(orient='records', indent=1)
            json_data2 = df2.to_json(orient='records', indent=1)
            json_data3 = df3.to_json(orient='records', indent=1)
            json_data4 = df4.to_json(orient='records', indent=1)
            json_data5 = df5.to_json(orient='records', indent=1)
            json_data6 = df6.to_json(orient='records', indent=1)
            json_data7 = df7.to_json(orient='records', indent=1)
            json_data = {"Surplus_wheat": json_data1, "Deficit_wheat": json_data2, "Surplus_rice": json_data3, "Deficit_rice": json_data4, "States_supply": json_data5, "States_allocation": json_data6, "Rail_cost_chart": json_data7}
        except:
            json_data = json.dumps({"Status": 0}, indent=1)

        json_object = json.dumps(json_data)
        return json_object
    else:
        return ("error")
    
@app.route("/read_Daily_Template_S1",methods = ["POST","GET"])
def read_Daily_Template_S1():
    if request.method == "POST":        
        try: 
            df1 = pd.read_excel('Input\\Daily_Template_Scene1.xlsx', sheet_name="Surplus_wheat") 
            df2 = pd.read_excel('Input\\Daily_Template_Scene1.xlsx', sheet_name="Deficit_wheat")
            df3 = pd.read_excel('Input\\Daily_Template_Scene1.xlsx', sheet_name="Surplus_rice")
            df4 = pd.read_excel('Input\\Daily_Template_Scene1.xlsx', sheet_name="Deficit_rice")
            df5 = pd.read_excel('Input\\Daily_Template_Scene1.xlsx', sheet_name="States_supply")
            df6 = pd.read_excel('Input\\Daily_Template_Scene1.xlsx', sheet_name="States_allocation")
            df7 = pd.read_excel('Input\\Daily_Template_Scene1.xlsx', sheet_name="Rail_cost_chart")
            json_data1 = df1.to_json(orient='records', indent=1)
            json_data2 = df2.to_json(orient='records', indent=1)
            json_data3 = df3.to_json(orient='records', indent=1)
            json_data4 = df4.to_json(orient='records', indent=1)
            json_data5 = df5.to_json(orient='records', indent=1)
            json_data6 = df6.to_json(orient='records', indent=1)
            json_data7 = df7.to_json(orient='records', indent=1)
            json_data = {"Surplus_wheat": json_data1, "Deficit_wheat": json_data2, "Surplus_rice": json_data3, "Deficit_rice": json_data4, "States_supply": json_data5, "States_allocation": json_data6, "Rail_cost_chart": json_data7}
        except:
            json_data = json.dumps({"Status": 0}, indent=1)

        json_object = json.dumps(json_data)
        return json_object
    else:
        return ("error")
    
@app.route("/read_Daily_Template_S2",methods = ["POST","GET"])
def read_Daily_Template_S2():
    if request.method == "POST":        
        try: 
            df1 = pd.read_excel('Input\\Daily_Template_Scene2.xlsx', sheet_name="Surplus_wheat") 
            df2 = pd.read_excel('Input\\Daily_Template_Scene2.xlsx', sheet_name="Deficit_wheat")
            df3 = pd.read_excel('Input\\Daily_Template_Scene2.xlsx', sheet_name="Surplus_rice")
            df4 = pd.read_excel('Input\\Daily_Template_Scene2.xlsx', sheet_name="Deficit_rice")
            df5 = pd.read_excel('Input\\Daily_Template_Scene2.xlsx', sheet_name="States_supply")
            df6 = pd.read_excel('Input\\Daily_Template_Scene2.xlsx', sheet_name="States_allocation")
            df7 = pd.read_excel('Input\\Daily_Template_Scene2.xlsx', sheet_name="Rail_cost_chart")
            json_data1 = df1.to_json(orient='records', indent=1)
            json_data2 = df2.to_json(orient='records', indent=1)
            json_data3 = df3.to_json(orient='records', indent=1)
            json_data4 = df4.to_json(orient='records', indent=1)
            json_data5 = df5.to_json(orient='records', indent=1)
            json_data6 = df6.to_json(orient='records', indent=1)
            json_data7 = df7.to_json(orient='records', indent=1)
            json_data = {"Surplus_wheat": json_data1, "Deficit_wheat": json_data2, "Surplus_rice": json_data3, "Deficit_rice": json_data4, "States_supply": json_data5, "States_allocation": json_data6, "Rail_cost_chart": json_data7}
        except:
            json_data = json.dumps({"Status": 0}, indent=1)

        json_object = json.dumps(json_data)
        return json_object
    else:
        return ("error")
    
@app.route("/Download_Template_to_add", methods=["POST", "GET"])
def Download_Template_to_add():
    if request.method == "POST":
        try:
            df1 = pd.read_excel('Input\\Matrices.xlsx', sheet_name="Railhead_cost_matrix_1rake") 
            df2 = pd.read_excel('Input\\Monthly_Template_M1.xlsx', sheet_name="Surplus_wheat") 

            prev_col = df1.columns
            present_col = df2["Railhead"]

            prev_st = set(prev_col)
            add_rh = []
            for rh in present_col:
                if rh not in prev_st:
                    add_rh.append(rh)

            # Create a dictionary with "Railhead" as the first column
            data = {"": add_rh}

            # Add other columns
            for col in present_col:
                data[col] = [""] * len(add_rh)

            Excel_data = pd.DataFrame(data)

            json_data1 = Excel_data.to_json(orient='records', indent=1)

            json_data = {
                "Railhead_cost_matrix_1rake": json_data1,
                "Railhead_cost_matrix": json_data1,
                "Railhead_dist_matrix": json_data1
            }
        except Exception as e:
            json_data = {"Status": 0, "Error": str(e)}
        json_object = json.dumps(json_data, indent=1)
        return json_object
    else:
        return "error"



    
@app.route("/Monthly_readPickle",methods = ["POST","GET"])
def Monthly_readPickle():
    try:
        dbfile = open('Output\\OutputPickle.pkl', 'rb')     
        db = pickle.load(dbfile)
        dbfile.close()
    except:
        db = {}
        db["status"] = 0
    return(json.dumps(db, indent = 1))


@app.route("/Update_matrices",methods = ["POST"])
def Update_matrices():
    data = {}
    try:
        file = request.files['uploadFile']
        file.save("Input//Update_matrices.xlsx")
        Railhead_cost_matrix_1rake_U_data = pd.read_excel("Input/Update_matrices.xlsx", sheet_name="Railhead_cost_matrix_1rake", index_col=0)
        Railhead_cost_matrix_U_data = pd.read_excel("Input/Update_matrices.xlsx", sheet_name="Railhead_cost_matrix", index_col=0)
        Railhead_dist_matrix_U_data = pd.read_excel("Input/Update_matrices.xlsx", sheet_name="Railhead_dist_matrix", index_col=0)
        Railhead_cost_matrix_1rake_M_data = pd.read_excel("Input/Matrices.xlsx", sheet_name="Railhead_cost_matrix_1rake", index_col=0)
        Railhead_cost_matrix_M_data = pd.read_excel("Input/Matrices.xlsx", sheet_name="Railhead_cost_matrix", index_col=0)
        Railhead_dist_matrix_M_data = pd.read_excel("Input/Matrices.xlsx", sheet_name="Railhead_dist_matrix", index_col=0)

        for row in Railhead_cost_matrix_1rake_U_data.index:
            for col in Railhead_cost_matrix_1rake_U_data.columns:
                value = Railhead_cost_matrix_1rake_U_data.loc[row, col]
                Railhead_cost_matrix_1rake_M_data.at[row, col] = value
                Railhead_cost_matrix_1rake_M_data.at[col, row] = value

        for row in Railhead_cost_matrix_U_data.index:
            for col in Railhead_cost_matrix_U_data.columns:
                value = Railhead_cost_matrix_U_data.loc[row, col]
                Railhead_cost_matrix_M_data.at[row, col] = value
                Railhead_cost_matrix_M_data.at[col, row] = value

        for row in Railhead_dist_matrix_U_data.index:
            for col in Railhead_dist_matrix_U_data.columns:
                value = Railhead_dist_matrix_U_data.loc[row, col]
                Railhead_dist_matrix_M_data.at[row, col] = value
                Railhead_dist_matrix_M_data.at[col, row] = value

        with pd.ExcelWriter("Input/Matrices.xlsx",mode='a',engine='openpyxl', if_sheet_exists='replace') as writer:
            Railhead_cost_matrix_1rake_M_data.to_excel(writer,sheet_name="Railhead_cost_matrix_1rake", index=True)
            Railhead_cost_matrix_M_data.to_excel(writer,sheet_name="Railhead_cost_matrix", index=True)
            Railhead_dist_matrix_M_data.to_excel(writer,sheet_name="Railhead_dist_matrix", index=True)

        data['status'] = 1
    except:
        data['status'] = 0
    
    json_data = json.dumps(data)
    json_object = json.loads(json_data)

    return(json.dumps(json_object, indent = 1))


@app.route("/Alternate_Railhead_readPickle",methods = ["POST","GET"])
def Alternate_Railhead_readPickle():
    try:
        dbfile = open('Output\\Alternate_Railhead.pkl', 'rb')     
        db = pickle.load(dbfile)
        dbfile.close()
        blank_data = []
        with open('Output\\Alternate_Railhead.pkl', 'wb') as f:
            pickle.dump(blank_data, f)
    except:
        db = {}
        db["status"] = 0
    return jsonify(db)


@app.route("/Add_Railhead", methods=["POST", "GET"])
def Add_Railhead():
    try:
        Railhead_name = []
        Railhead_State = []
        fetched_data = request.get_json()
        Railhead_name.append(fetched_data["railhead"].upper())
        Railhead_State.append(fetched_data['state'])
        Monthly_Template_M1 = 'Input\\Monthly_Template_M1.xlsx'
        Daily_Template_S1 = 'Input\\Daily_Template_Scene1.xlsx'
        Daily_Template_S2 = 'Input\\Daily_Template_Scene2.xlsx'
        Data_sheet = 'Frontend/public/data/Updated_railhead_list.xlsx'



        # Sheets
        Monthly_Sheets = ["Surplus_wheat", "Deficit_wheat", "Surplus_rice", "Deficit_rice"]
        Daily_Sheets_S1 = ["Surplus_wheat", "Deficit_wheat", "Surplus_rice", "Deficit_rice"]
        Daily_Sheets_S2 = ["Surplus_wheat", "Deficit_wheat", "Surplus_rice", "Deficit_rice"]

        # Initialize lists for DataFrames
        Monthly_df = []
        Daily_S1_df = []
        Daily_S2_df = []
        Data_sheets = pd.read_excel(Data_sheet, sheet_name="RH_Sheet")

        # Read data from Excel files and store in lists
        for sheets in Monthly_Sheets:
            x = pd.read_excel(Monthly_Template_M1, sheet_name=sheets)
            Monthly_df.append(x)

        for sheets in Daily_Sheets_S1:
            x = pd.read_excel(Daily_Template_S1, sheet_name=sheets)
            Daily_S1_df.append(x)

        for sheets in Daily_Sheets_S2:
            x = pd.read_excel(Daily_Template_S2, sheet_name=sheets)
            Daily_S2_df.append(x)

        for i in range(len(Railhead_name)):
                Data_sheets = pd.concat([Data_sheets, pd.DataFrame({"RH_code": [Railhead_name[i]], "State": [Railhead_State[i]]})])
        # Append data to the DataFrames
        for i in range(len(Monthly_Sheets)):
            for j in range(len(Railhead_name)):
                Monthly_df[i] = pd.concat([Monthly_df[i], pd.DataFrame({"Railhead": [Railhead_name[j]], "State": [Railhead_State[j]]})])

        for i in range(len(Daily_Sheets_S1)):
            for j in range(len(Railhead_name)):
                Daily_S1_df[i] = pd.concat([Daily_S1_df[i], pd.DataFrame({"Railhead": [Railhead_name[j]], "State": [Railhead_State[j]]})])

        for i in range(len(Daily_Sheets_S2)):
            for j in range(len(Railhead_name)):
                Daily_S2_df[i] = pd.concat([Daily_S2_df[i], pd.DataFrame({"Railhead": [Railhead_name[j]], "State": [Railhead_State[j]]})])

        # Write modified DataFrames back to Excel files
        with pd.ExcelWriter("Input\\Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            for i in range(len(Monthly_Sheets)):
                Monthly_df[i].to_excel(writer, sheet_name=Monthly_Sheets[i], index=False)

        with pd.ExcelWriter("Input\\Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            for i in range(len(Daily_Sheets_S1)):
                Daily_S1_df[i].to_excel(writer, sheet_name=Daily_Sheets_S1[i], index=False)

        with pd.ExcelWriter("Input\\Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            for i in range(len(Daily_Sheets_S2)):
                Daily_S2_df[i].to_excel(writer, sheet_name=Daily_Sheets_S2[i], index=False)

        with pd.ExcelWriter("Frontend\\public\\data\\Updated_railhead_list.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            Data_sheets.to_excel(writer, sheet_name="RH_Sheet", index=False)

        db = {"status": 1, "message": "Railhead names and states added successfully"}
    except Exception as e:
        db = {"status": 0, "message": str(e)}

    return json.dumps(db, indent=1)

@app.route('/getMonthlyExcelData')
def get_monthly_excel_data():
    Monthly_Template_M1 = 'Input\\Monthly_Template_M1.xlsx'
    excel_path = os.path.join(os.path.dirname(__file__), Monthly_Template_M1)
    return send_file(excel_path, as_attachment=True)

@app.route('/getDaily1ExcelData')
def get_daily_scen1_excel_data():
    Monthly_Template_M1 = 'Input\\Daily_Template_Scene1.xlsx'
    excel_path = os.path.join(os.path.dirname(__file__), Monthly_Template_M1)
    return send_file(excel_path, as_attachment=True)

@app.route('/getDaily2ExcelData')
def get_daily_scen2_excel_data():
    Monthly_Template_M1 = 'Input\\Daily_Template_Scene2.xlsx'
    excel_path = os.path.join(os.path.dirname(__file__), Monthly_Template_M1)
    return send_file(excel_path, as_attachment=True)

@app.route("/Modify_Monthly_Template_M01", methods=["POST", "GET"])
def Modify_Monthly_Template_M01():
    try:
        def try_float(value):
            try:
                return float(value)
            except (ValueError, TypeError):
                return value
        fetched_data = request.get_json()
        sheets = fetched_data['SheetNames']
        # print(type(fetched_data['Sheets']['Surplus_wheat'][f'C{3}']['v']))
        # print(fetched_data)

        for sht in sheets:
            if sht == 'Surplus_wheat':
                columns = ['Railhead', 'State', 'Supply']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                supply = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Supply': supply})
                with pd.ExcelWriter("Input/Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Surplus_wheat", index=False)

            elif sht == 'Deficit_wheat':
                columns = ['Railhead', 'State', 'Demand', 'Capacity']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                Demand = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                Capacity = [try_float(sht_data[f'D{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Demand': Demand, 'Capacity': Capacity})
                with pd.ExcelWriter("Input/Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Deficit_wheat", index=False)

            elif sht == 'Surplus_rice':
                columns = ['Railhead', 'State', 'Supply']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                supply = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Supply': supply})
                with pd.ExcelWriter("Input/Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Surplus_rice", index=False)

            elif sht == 'Deficit_rice':
                columns = ['Railhead', 'State', 'Demand']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                demand = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Demand': demand})
                with pd.ExcelWriter("Input/Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Deficit_rice", index=False)

            elif sht == 'States_supply':
                columns = ['State', 'Supply_wheat', 'Supply_rice']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                State = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                Supply_wheat = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                Supply_rice = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'State': State, 'Supply_wheat': Supply_wheat, 'Supply_rice': Supply_rice})
                with pd.ExcelWriter("Input/Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="States_supply", index=False)

            elif sht == 'States_allocation':
                columns = ['States', 'Alloc_wheat', 'Alloc_rice', 'Capacity']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                States = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                Alloc_wheat = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                Alloc_rice = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                Capacity = [try_float(sht_data[f'D{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'States': States, 'Alloc_wheat': Alloc_wheat, 'Alloc_rice': Alloc_rice, 'Capacity': Capacity})
                with pd.ExcelWriter("Input/Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="States_allocation", index=False)

            elif sht == 'Rail_cost_chart':
                columns = ['From', 'To', 'Rate per Ton']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                From = [try_float(sht_data[f'A{i}']['v']) for i in range(3, length + 1)]
                To = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]
                Rate_per_Ton = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'From': From, 'To': To, 'Rate per Ton': Rate_per_Ton})
                with pd.ExcelWriter("Input/Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Rail_cost_chart", index=False)

        db = {"status": 1, "message": "Railhead names and states added successfully"}
    except Exception as e:
        db = {"status": 0, "message": str(e)}

    return json.dumps(db, indent=1)

@app.route("/Modify_Daily_Template_S01", methods=["POST", "GET"])
def Modify_Daily_Template_S01():
    try:
        def try_float(value):
            try:
                return float(value)
            except (ValueError, TypeError):
                return value
        fetched_data = request.get_json()
        sheets = fetched_data['SheetNames']
        # print(type(fetched_data['Sheets']['Surplus_wheat'][f'C{3}']['v']))
        # print(fetched_data)

        for sht in sheets:
            if sht == 'Surplus_wheat':
                columns = ['Railhead', 'State', 'Supply']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                supply = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Supply': supply})
                with pd.ExcelWriter("Input/Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Surplus_wheat", index=False)

            elif sht == 'Deficit_wheat':
                columns = ['Railhead', 'State', 'Demand']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                Demand = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Demand': Demand})
                with pd.ExcelWriter("Input/Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Deficit_wheat", index=False)

            elif sht == 'Surplus_rice':
                columns = ['Railhead', 'State', 'Supply']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                supply = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Supply': supply})
                with pd.ExcelWriter("Input/Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Surplus_rice", index=False)

            elif sht == 'Deficit_rice':
                columns = ['Railhead', 'State', 'Demand']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                demand = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Demand': demand})
                with pd.ExcelWriter("Input/Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Deficit_rice", index=False)

            elif sht == 'States_supply':
                columns = ['State', 'Supply_wheat', 'Supply_rice']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                State = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                Supply_wheat = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                Supply_rice = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'State': State, 'Supply_wheat': Supply_wheat, 'Supply_rice': Supply_rice})
                with pd.ExcelWriter("Input/Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="States_supply", index=False)

            elif sht == 'States_allocation':
                columns = ['States', 'Alloc_wheat', 'Alloc_rice']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                States = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                Alloc_wheat = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                Alloc_rice = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'States': States, 'Alloc_wheat': Alloc_wheat, 'Alloc_rice': Alloc_rice})
                with pd.ExcelWriter("Input/Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="States_allocation", index=False)

            elif sht == 'Rail_cost_chart':
                columns = ['From', 'To', 'Rate per Ton']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                From = [try_float(sht_data[f'A{i}']['v']) for i in range(3, length + 1)]
                To = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]
                Rate_per_Ton = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'From': From, 'To': To, 'Rate per Ton': Rate_per_Ton})
                with pd.ExcelWriter("Input/Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Rail_cost_chart", index=False)

        db = {"status": 1, "message": "Railhead names and states added successfully"}
    except Exception as e:
        db = {"status": 0, "message": str(e)}

    return json.dumps(db, indent=1)

@app.route("/Modify_Daily_Template_S02", methods=["POST", "GET"])
def Modify_Daily_Template_S02():
    try:
        def try_float(value):
            try:
                return float(value)
            except (ValueError, TypeError):
                return value
        fetched_data = request.get_json()
        sheets = fetched_data['SheetNames']
        # print(type(fetched_data['Sheets']['Surplus_wheat'][f'C{3}']['v']))
        # print(fetched_data)

        for sht in sheets:
            if sht == 'Surplus_wheat':
                columns = ['Railhead', 'State', 'Supply']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                supply = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Supply': supply})
                with pd.ExcelWriter("Input/Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Surplus_wheat", index=False)

            elif sht == 'Deficit_wheat':
                columns = ['Railhead', 'State', 'Demand']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                Demand = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Demand': Demand})
                with pd.ExcelWriter("Input/Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Deficit_wheat", index=False)

            elif sht == 'Surplus_rice':
                columns = ['Railhead', 'State', 'Supply']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                supply = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Supply': supply})
                with pd.ExcelWriter("Input/Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Surplus_rice", index=False)

            elif sht == 'Deficit_rice':
                columns = ['Railhead', 'State', 'Demand']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                Railhead = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                state = [sht_data[f'B{i}']['v'] for i in range(3, length + 1)]
                demand = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'Railhead': Railhead, 'State': state, 'Demand': demand})
                with pd.ExcelWriter("Input/Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Deficit_rice", index=False)

            elif sht == 'States_supply':
                columns = ['State', 'Supply_wheat', 'Supply_rice']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                State = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                Supply_wheat = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                Supply_rice = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'State': State, 'Supply_wheat': Supply_wheat, 'Supply_rice': Supply_rice})
                with pd.ExcelWriter("Input/Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="States_supply", index=False)

            elif sht == 'States_allocation':
                columns = ['States', 'Alloc_wheat', 'Alloc_rice']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                States = [sht_data[f'A{i}']['v'] for i in range(3, length + 1)]
                Alloc_wheat = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                Alloc_rice = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'States': States, 'Alloc_wheat': Alloc_wheat, 'Alloc_rice': Alloc_rice})
                with pd.ExcelWriter("Input/Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="States_allocation", index=False)

            elif sht == 'Rail_cost_chart':
                columns = ['From', 'To', 'Rate per Ton']
                sht_data = fetched_data['Sheets'][sht]
                length = len(sht_data) // len(columns)
                From = [try_float(sht_data[f'A{i}']['v']) for i in range(3, length + 1)]
                To = [try_float(sht_data[f'B{i}']['v']) for i in range(3, length + 1)]
                Rate_per_Ton = [try_float(sht_data[f'C{i}']['v']) for i in range(3, length + 1)]  # Convert to float
                df = pd.DataFrame({'From': From, 'To': To, 'Rate per Ton': Rate_per_Ton})
                with pd.ExcelWriter("Input/Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Rail_cost_chart", index=False)

        db = {"status": 1, "message": "Railhead names and states added successfully"}
    except Exception as e:
        db = {"status": 0, "message": str(e)}

    return json.dumps(db, indent=1)




@app.route("/Remove_Railhead", methods=["POST", "GET"])
def Remove_Railhead():
    try:
        Railhead_name = []
        Railhead_State = []
        fetched_data = request.get_json()
        Railhead_name.append(fetched_data["railhead"].upper())
        Railhead_State.append(fetched_data['state'])
        Monthly_Template_M1 = 'Input\\Monthly_Template_M1.xlsx'
        Daily_Template_S1 = 'Input\\Daily_Template_Scene1.xlsx'
        Daily_Template_S2 = 'Input\\Daily_Template_Scene1.xlsx'
        Data_sheet = 'Frontend\\public\\data\\Updated_railhead_list.xlsx'

        Monthly_Sheets = ["Surplus_wheat", "Deficit_wheat", "Surplus_rice", "Deficit_rice"]
        Daily_Sheets_S1 = ["Surplus_wheat", "Deficit_wheat", "Surplus_rice", "Deficit_rice"]
        Daily_Sheets_S2 = ["Surplus_wheat", "Deficit_wheat", "Surplus_rice", "Deficit_rice"]

        Monthly_df = []
        Daily_S1_df = []
        Daily_S2_df = []
        Data_sheets = pd.read_excel(Data_sheet, sheet_name="RH_Sheet")

        for sheets in Monthly_Sheets:
            x = pd.read_excel(Monthly_Template_M1, sheet_name=sheets)
            Monthly_df.append(x)
        for sheets in Daily_Sheets_S1:
            x = pd.read_excel(Daily_Template_S1, sheet_name=sheets)
            Daily_S1_df.append(x)
        for sheets in Daily_Sheets_S2:
            x = pd.read_excel(Daily_Template_S2, sheet_name=sheets)
            Daily_S2_df.append(x)

        for i in range(len(Monthly_Sheets)):
            for j in range(len(Railhead_name)):
                for df in [Monthly_df[i], Daily_S1_df[i], Daily_S2_df[i]]:
                        df.drop(df[df["Railhead"] == Railhead_name[j]].index, inplace=True)
                        
        for i in range(len(Railhead_name)):
            Data_sheets.drop(Data_sheets[Data_sheets["RH_code"] == Railhead_name[i]].index, inplace=True)

        with pd.ExcelWriter("Input\\Monthly_Template_M1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            for i in range(len(Monthly_Sheets)):
                Monthly_df[i].to_excel(writer, sheet_name=Monthly_Sheets[i], index=False)
        with pd.ExcelWriter("Input\\Daily_Template_Scene1.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            for i in range(len(Daily_Sheets_S1)):
                Daily_S1_df[i].to_excel(writer, sheet_name=Daily_Sheets_S1[i], index=False)
        with pd.ExcelWriter("Input\\Daily_Template_Scene2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            for i in range(len(Daily_Sheets_S2)):
                Daily_S2_df[i].to_excel(writer, sheet_name=Daily_Sheets_S2[i], index=False)
        with pd.ExcelWriter("Frontend\\public\\data\\Updated_railhead_list.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            Data_sheets.to_excel(writer, sheet_name="RH_Sheet", index=False)


        db = {"status": 1, "message": "Railhead names and states added successfully"}
    except Exception as e:
        db = {"status": 0, "message": str(e)}

    return json.dumps(db, indent=1)



@app.route("/Monthly_Solution",methods = ["POST","GET"])
def Monthly_Solution():
    data1 = {}
    if request.method == "POST":
        try:
            r_s = 25
            r_d = 25

            fetched_data = request.get_json()
            r_s_fetched = fetched_data['r_s']
            r_d_fetched = fetched_data['r_d']
            TEFD_fetched = fetched_data['TEFD']

            if (r_s_fetched != ''):
                r_s = int(r_s_fetched)
            if r_d_fetched != '':
                r_d = int(r_d_fetched)


            data=pd.ExcelFile("Input\\Monthly_Template_M1.xlsx")
            surplus_wheat=pd.read_excel(data,sheet_name="Surplus_wheat",index_col=1)
            deficit_wheat=pd.read_excel(data,sheet_name="Deficit_wheat",index_col=1)
            surplus_rice=pd.read_excel(data,sheet_name="Surplus_rice",index_col=1)
            deficit_rice=pd.read_excel(data,sheet_name="Deficit_rice",index_col=1)
            states_alloc=pd.read_excel(data,sheet_name="States_allocation",index_col=0)
            rail_cost = None
            if TEFD_fetched == 'NON-TEFD':
                rail_cost=pd.read_excel("Input\\Non-TEFD.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
            elif TEFD_fetched == 'TEFD':
                rail_cost=pd.read_excel("Input\\TEFD.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
            elif TEFD_fetched == 'Non-TEFD+TC':
                rail_cost=pd.read_excel("Input\\Non_TEFD_TC.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
            else:
                rail_cost=pd.read_excel("Input\\TEFD_TC.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
            


            prob = LpProblem("Output\\FCI_monthly_model_allocation_rr",LpMinimize)
            x_ij_wheat=LpVariable.dicts("x_wheat",[(i,j) for i in surplus_wheat.index for j in deficit_wheat.index],0)
            x_ij_rice=LpVariable.dicts("x_rice",[(i,j) for i in surplus_rice.index for j in deficit_rice.index],0)
            

            prob+=lpSum(x_ij_wheat[(i,j)]*rail_cost.loc[i][j] for i in surplus_wheat.index for j in deficit_wheat.index)+lpSum(x_ij_rice[(i,j)]*rail_cost.loc[i][j] for i in surplus_rice.index for j in deficit_rice.index)            
            
            for i in surplus_wheat.index:
                for j in deficit_wheat.index:
                    if i==j:
                        prob+=x_ij_wheat[(i,j)]==0
                        
            for i in surplus_rice.index:
                for j in deficit_rice.index:
                    if i==j:
                        prob+=x_ij_rice[(i,j)]==0
                        
            for i in surplus_wheat.index:
                prob+=lpSum(x_ij_wheat[(i,j)] for j in deficit_wheat.index)<=surplus_wheat["Supply"][i]
            for i in surplus_rice.index:
                prob+=lpSum(x_ij_rice[(i,j)] for j in deficit_rice.index)<=surplus_rice["Supply"][i]
                
            for i in surplus_wheat.index:
                prob+=lpSum(x_ij_wheat[(i,j)] for j in deficit_wheat.index)+lpSum(x_ij_rice[(i,j)] for j in deficit_rice.index)<=r_s

           
            
            for a in states_alloc.index:
                prob+=lpSum(x_ij_wheat[(i,j)] for i in surplus_wheat.index for j in deficit_wheat.index if deficit_wheat.loc[j]["State"]==a)>=states_alloc.loc[a]["Alloc_wheat"]
                prob+=lpSum(x_ij_wheat[(i,j)] for i in surplus_wheat.index for j in deficit_wheat.index if deficit_wheat.loc[j]["State"]==a)<=states_alloc.loc[a]["Alloc_wheat"]
            for a in states_alloc.index:
                prob+=lpSum(x_ij_rice[(i,j)] for i in surplus_rice.index for j in deficit_rice.index if deficit_rice.loc[j]["State"]==a)>=states_alloc.loc[a]["Alloc_rice"]
                prob+=lpSum(x_ij_rice[(i,j)] for i in surplus_rice.index for j in deficit_rice.index if deficit_rice.loc[j]["State"]==a)<=states_alloc.loc[a]["Alloc_rice"]
            for j in deficit_wheat.index:
                prob+=lpSum(x_ij_wheat[(i,j)] for i in surplus_wheat.index)+lpSum(x_ij_rice[(i,j)] for i in surplus_rice.index)<=deficit_wheat["Capacity"][j]
            for j in deficit_wheat.index:
                prob+=lpSum(x_ij_wheat[(i,j)] for i in surplus_wheat.index)+lpSum(x_ij_rice[(i,j)] for i in surplus_rice.index)<=r_d
            

            prob.writeLP("Output\\FCI_monthly_model_allocation_rr.lp")
            prob.solve()
            print("Status:", LpStatus[prob.status])
            print("Minimum Cost of Transportation = Rs.", value(prob.objective),"Lakh")
            print("Total Number of Variables:",len(prob.variables()))
            print("Total Number of Constraints:",len(prob.constraints)) 
            result_data = {"Minimum Cost of Transportation": value(prob.objective), "Total Number of Variables": len(prob.variables()),"Total Number of Constraints":len(prob.constraints)}
            with open('Output\\OutputPickle.pkl', 'wb') as f:
                pickle.dump(result_data, f)   
            r_wheat={}
            r_wheat=pd.DataFrame([],index=surplus_wheat.index,columns=deficit_wheat.index)
                
            for r in surplus_wheat.index:
                for j in deficit_wheat.index:
                    r_wheat.loc[r][j]=x_ij_wheat[(r,j)].value()
                        
            r_rice={}
            r_rice=pd.DataFrame([],index=surplus_rice.index,columns=deficit_rice.index)

            for r in surplus_rice.index:
                for j in deficit_rice.index:
                    r_rice.loc[r][j]=x_ij_rice[(r,j)].value()
                    
            with pd.ExcelWriter("Output//Total_Results.xlsx",mode='a',engine='openpyxl', if_sheet_exists='replace') as writer:
                r_wheat.to_excel(writer,sheet_name="r_wheat",float_format="%0.3f")
                r_rice.to_excel(writer,sheet_name="r_rice",float_format="%0.3f")
            
            relevant_data=pd.ExcelFile("Output//Total_Results.xlsx")
            relevant_r_wheat=pd.read_excel(relevant_data,sheet_name="r_wheat",index_col=0)
            relevant_r_rice=pd.read_excel(relevant_data,sheet_name="r_rice",index_col=0)
            relevant_Dict_wheat={}
            relevant_Dict_rice={}
            for i in range(len(relevant_r_wheat.index)):
                for j in range(len(relevant_r_wheat.columns)):
                    if relevant_r_wheat.iat[i,j]>0:
                        relevant_Dict_wheat[relevant_r_wheat.index[i],relevant_r_wheat.columns[j]]=relevant_r_wheat.iloc[i][relevant_r_wheat.columns[j]]
                        
            for i in range(len(relevant_r_rice.index)):
                for j in range(len(relevant_r_rice.columns)):
                    if relevant_r_rice.iat[i,j]>0:
                        relevant_Dict_rice[relevant_r_rice.index[i],relevant_r_rice.columns[j]]=relevant_r_rice.iloc[i][relevant_r_rice.columns[j]]
            
            L1=list(relevant_Dict_wheat.keys())
            L2=list(relevant_Dict_wheat.values())
            A=[]
            B=[]
            C=[]

            df_wheat=pd.DataFrame()

            for i in range(len(L1)):
                A.append(L1[i][0])
                B.append(L1[i][1])
                C.append(L2[i])
                

            df_wheat["From"]=A
            df_wheat["To"]=B
            df_wheat["Values"]=C

            From_state=[]
            To_state=[]
            Commodity=[]

            for i in range(len(L1)):
                for j in surplus_wheat.index:
                    if L1[i][0]==j:
                        From_state.append(surplus_wheat.loc[j]["State"])
                        
            for i in range(len(L1)):
                for j in surplus_wheat.index:
                    if L1[i][1]==j:
                        To_state.append(surplus_wheat.loc[j]["State"])
                        
            for i in range(len(L1)):
                Commodity.append("Wheat")
                
                
            df_wheat.insert(1,"From_state",From_state)
            df_wheat.insert(3,"To_state",To_state)
            df_wheat.insert(4,"Commodity",Commodity)

            L3=list(relevant_Dict_rice.keys())
            L4=list(relevant_Dict_rice.values())

            D=[]
            E=[]
            F=[]

            df_rice=pd.DataFrame()

            for i in range(len(L3)):
                D.append(L3[i][0])
                E.append(L3[i][1])
                F.append(L4[i])
                
            df_rice["From"]=D
            df_rice["To"]=E
            df_rice["Values"]=F

            From_state_rice=[]
            To_state_rice=[]
            Commodity_rice=[]

            for i in range(len(L3)):
                for j in surplus_wheat.index:
                    if L3[i][0]==j:
                        From_state_rice.append(surplus_wheat.loc[j]["State"])
                        
            for i in range(len(L3)):
                for j in surplus_wheat.index:
                    if L3[i][1]==j:
                        To_state_rice.append(surplus_wheat.loc[j]["State"])
                        
            for i in range(len(L3)):
                Commodity_rice.append("Rice")
                
            df_rice.insert(1,"From_state",From_state_rice)
            df_rice.insert(3,"To_state",To_state_rice)
            df_rice.insert(4,"Commodity",Commodity_rice)

            with pd.ExcelWriter("Output//Relevent_Results.xlsx",mode='a',engine='openpyxl', if_sheet_exists='replace') as writer:
                df_wheat.to_excel(writer,sheet_name="wheat")
                df_rice.to_excel(writer,sheet_name="rice")
            
            table = pd.ExcelFile("Output//Relevent_Results.xlsx")
            table_data_w = pd.read_excel(table,sheet_name="wheat",index_col=0)
            table_data_r = pd.read_excel(table,sheet_name="rice",index_col=0)
            tab_A_w = list(table_data_w["From_state"])
            tab_B_w = list(table_data_w["To_state"])
            tab_A_r = list(table_data_r["From_state"])
            tab_B_r = list(table_data_r["To_state"])
            A_main_w = []
            B_main_w = []
            C_main_w = []
            table_set_w = set()

            A_main_r = []
            B_main_r = []
            C_main_r = []
            table_set_r = set()

            for i in range(len(tab_A_w)):
                table_set_w.add((tab_A_w[i],tab_B_w[i]))
            for i in range(len(tab_A_r)):
                table_set_r.add((tab_A_r[i],tab_B_r[i]))

            for i in table_set_w:
                source = i[0]
                dest = i[1]
                val = 0
                for _, row in table_data_w.iterrows():
                    if (row["From_state"] == source) and (row["To_state"] == dest):
                        val += int(row["Values"])
                A_main_w.append(source)
                B_main_w.append(dest)
                C_main_w.append(val)

            for i in table_set_r:
                source = i[0]
                dest = i[1]
                val = 0
                for _, row in table_data_r.iterrows():
                    if (row["From_state"] == source) and (row["To_state"] == dest):
                        val += int(row["Values"])
                A_main_r.append(source)
                B_main_r.append(dest)
                C_main_r.append(val)

            df_table_w = pd.DataFrame()
            df_table_w["From_State"] = A_main_w
            df_table_w["To_State"] = B_main_w
            df_table_w["Values"] = C_main_w

            df_table_r = pd.DataFrame()
            df_table_r["From_State"] = A_main_r
            df_table_r["To_State"] = B_main_r
            df_table_r["Values"] = C_main_r

            pivot_table_w = df_table_w.pivot_table(index="From_State", columns="To_State", values="Values", fill_value=0)
            pivot_table_r = df_table_r.pivot_table(index="From_State", columns="To_State", values="Values", fill_value=0)

            with pd.ExcelWriter("Output//Monthly_State_To_State_Table.xlsx",mode='a',engine='openpyxl', if_sheet_exists='replace') as writer:
                pivot_table_w.to_excel(writer,sheet_name="Wheat")
                pivot_table_r.to_excel(writer,sheet_name="Rice")
            
            data1["status"] = 1
                  
        except Exception as e:
            print(e)
            data1["status"] = 0
        json_data = json.dumps(data1)
        json_object = json.loads(json_data)

        return(json.dumps(json_object, indent = 1))
    else:
        return ("error")

    

@app.route("/Daily_Planner",methods = ["POST","GET"])
def Daily_Planner():
    data1 = {}
    if request.method == "POST":
        try:
            blocked_org_rhcode = []
            blocked_dest_rhcode = []
            blocked_org_state = []
            blocked_dest_state = []

            confirmed_org_rhcode = []
            confirmed_dest_rhcode = []
            confirmed_org_state = []
            confirmed_dest_state = []
            confirmed_railhead_value = []
            confirmed_railhead_commodities = []
            
           
            fetched_data = request.get_json()
            blocked_data = fetched_data['block_data']
            confirmed_data = fetched_data['confirmed_data']
            Scenerio = fetched_data["Scenerio"]
            TEFD_fetched = fetched_data['TEFD']

            for i in range(len(blocked_data)):
                blocked_org_rhcode.append(blocked_data[i]["origin_railhead"])
                blocked_dest_rhcode.append(blocked_data[i]["destination_railhead"])
                blocked_org_state.append(blocked_data[i]["origin_state"])
                blocked_dest_state.append(blocked_data[i]["destination_state"])
            
            for i in range(len(confirmed_data)):
                confirmed_org_rhcode.append(confirmed_data[i]["origin_railhead"])
                confirmed_dest_rhcode.append(confirmed_data[i]["destination_railhead"])
                confirmed_org_state.append(confirmed_data[i]["origin_state"])
                confirmed_dest_state.append(confirmed_data[i]["destination_state"])
                confirmed_railhead_value.append(confirmed_data[i]["value_fixed"])
                confirmed_railhead_commodities.append(confirmed_data[i]["commodity_fixed"])
            
            if Scenerio == "Scenerio 2":
                data=pd.ExcelFile("Input\\Temp_balanced_DPT_scen2.xlsx")
                matrices_data = pd.ExcelFile("Input\\Matrices.xlsx")
                # blocking_data = pd.ExcelFile("Input\\Route_blocker_DPT.xlsx")

                surplus_wheat=pd.read_excel(data,sheet_name="Surplus_wheat",index_col=1)
                deficit_wheat=pd.read_excel(data,sheet_name="Deficit_wheat",index_col=1)
                surplus_rice=pd.read_excel(data,sheet_name="Surplus_rice",index_col=1)
                deficit_rice=pd.read_excel(data,sheet_name="Deficit_rice",index_col=1)
                states_alloc=pd.read_excel(data,sheet_name="States_allocation",index_col=0)
                rail_cost = None
                if TEFD_fetched == 'NON-TEFD':
                    rail_cost=pd.read_excel("Input\\Non-TEFD.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
                elif TEFD_fetched == 'TEFD':
                    rail_cost=pd.read_excel("Input\\TEFD.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
                elif TEFD_fetched == 'Non-TEFD+TC':
                    rail_cost=pd.read_excel("Input\\Non_TEFD_TC.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
                else:
                    rail_cost=pd.read_excel("Input\\TEFD_TC.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)

                prob = LpProblem("Output\\FCI_monthly_model_allocation_rr5",LpMinimize)

                x_ij_wheat=LpVariable.dicts("x_wheat",[(i,j) for i in surplus_wheat.index for j in deficit_wheat.index],0)
                x_ij_rice=LpVariable.dicts("x_rice",[(i,j) for i in surplus_rice.index for j in deficit_rice.index],0)



                prob+=lpSum(x_ij_wheat[(i,j)]*rail_cost.loc[i][j] for i in surplus_wheat.index for j in deficit_wheat.index)+lpSum(x_ij_rice[(i,j)]*rail_cost.loc[i][j] for i in surplus_rice.index for j in deficit_rice.index)
                
                for i in range(len(blocked_org_rhcode)):
                    prob += x_ij_rice[(blocked_org_rhcode[i], blocked_dest_rhcode[i])] == 0
                    prob += x_ij_wheat[(blocked_org_rhcode[i], blocked_dest_rhcode[i])] == 0
                
                for i in range(len(confirmed_org_rhcode)):
                    if confirmed_railhead_commodities[i] == "RICE":
                        prob += x_ij_rice[(confirmed_org_rhcode[i], confirmed_dest_rhcode[i])] == int(confirmed_railhead_value[i])
                    else:
                        prob += x_ij_wheat[(confirmed_org_rhcode[i], confirmed_dest_rhcode[i])] == int(confirmed_railhead_value[i])
                
                for i in surplus_wheat.index:
                    for j in deficit_wheat.index:
                        if i==j:
                            prob+=x_ij_wheat[(i,j)]==0
                            # print(x_ij_wheat[(i,j)]==0)
                            
                for i in surplus_rice.index:
                    for j in deficit_rice.index:
                        if i==j:
                            prob+=x_ij_rice[(i,j)]==0
                            # print(x_ij_rice[(i,j)]==0)

                for i in surplus_wheat.index:
                    prob+=lpSum(x_ij_wheat[(i,j)] for j in deficit_wheat.index)<=surplus_wheat["Supply"][i]
                    
                for i in surplus_rice.index:
                    prob+=lpSum(x_ij_rice[(i,j)] for j in deficit_rice.index)<=surplus_rice["Supply"][i]

                for j in deficit_wheat.index:
                    prob+=lpSum(x_ij_wheat[(i,j)] for i in surplus_wheat.index)+lpSum(x_ij_rice[(i,j)] for i in surplus_rice.index)<=1
                for j in deficit_wheat.index:
                    prob+=lpSum(x_ij_wheat[(i,j)] for i in surplus_wheat.index)+lpSum(x_ij_rice[(i,j)] for i in surplus_rice.index)<=deficit_wheat["Capacity"][j]

                for a in states_alloc.index:
                    prob+=lpSum(x_ij_wheat[(i,j)] for i in surplus_wheat.index for j in deficit_wheat.index if deficit_wheat.loc[j]["State"]==a)>=states_alloc.loc[a]["Alloc_wheat"]
                    prob+=lpSum(x_ij_wheat[(i,j)] for i in surplus_wheat.index for j in deficit_wheat.index if deficit_wheat.loc[j]["State"]==a)<=states_alloc.loc[a]["Alloc_wheat"]
                    
                for a in states_alloc.index:
                    prob+=lpSum(x_ij_rice[(i,j)] for i in surplus_rice.index for j in deficit_rice.index if deficit_rice.loc[j]["State"]==a)>=states_alloc.loc[a]["Alloc_rice"]
                    prob+=lpSum(x_ij_rice[(i,j)] for i in surplus_rice.index for j in deficit_rice.index if deficit_rice.loc[j]["State"]==a)<=states_alloc.loc[a]["Alloc_rice"]




                prob.writeLP("Output\\FCI_monthly_model_allocation_rr5.lp")
                prob.solve()
                print("Status:", LpStatus[prob.status])
                print("Minimum Cost of Transportation = Rs.", value(prob.objective),"Lakh")
                print("Total Number of Variables:",len(prob.variables()))
                print("Total Number of Constraints:",len(prob.constraints))


                r_wheat={}
                r_wheat=pd.DataFrame([],index=surplus_wheat.index,columns=deficit_wheat.index)
                    
                for r in surplus_wheat.index:
                    for j in deficit_wheat.index:
                        r_wheat.loc[r][j]=x_ij_wheat[(r,j)].value()
                        
                r_rice={}
                r_rice=pd.DataFrame([],index=surplus_rice.index,columns=deficit_rice.index)
                    
                for r in surplus_rice.index:
                    for j in deficit_rice.index:
                        r_rice.loc[r][j]=x_ij_rice[(r,j)].value()
                        
                with pd.ExcelWriter("Output\\Results_DPT2.xlsx",mode='a',engine='openpyxl', if_sheet_exists='replace') as writer:
                    r_wheat.to_excel(writer,sheet_name="r_wheat",float_format="%0.3f")
                    r_rice.to_excel(writer,sheet_name="r_rice",float_format="%0.3f")


                relevant_data = pd.ExcelFile("Output//Results_DPT2.xlsx")
                relevant_r_wheat = pd.read_excel(relevant_data, sheet_name="r_wheat", index_col=0)
                relevant_r_rice = pd.read_excel(relevant_data, sheet_name="r_rice", index_col=0)
                relevant_Dict_wheat = {}
                relevant_Dict_rice = {}
                Rice_cost = []
                Wheat_cost = []

                # x_ij_wheat[(i,j)]*rail_cost.loc[i][j]
                for i in range(len(relevant_r_wheat.index)):
                    for j in range(len(relevant_r_wheat.columns)):
                        if relevant_r_wheat.iat[i, j] > 0:
                            relevant_Dict_wheat[relevant_r_wheat.index[i], relevant_r_wheat.columns[j]] = relevant_r_wheat.iloc[i][relevant_r_wheat.columns[j]]
                            Wheat_cost.append((x_ij_wheat[(relevant_r_wheat.index[i],relevant_r_wheat.columns[j])]*rail_cost.loc[relevant_r_wheat.index[i]][relevant_r_wheat.columns[j]]))

                for i in range(len(relevant_r_rice.index)):
                    for j in range(len(relevant_r_rice.columns)):
                        if relevant_r_rice.iat[i, j] > 0:
                            relevant_Dict_rice[relevant_r_rice.index[i], relevant_r_rice.columns[j]] = relevant_r_rice.iloc[i][relevant_r_rice.columns[j]]
                            Rice_cost.append((x_ij_wheat[(relevant_r_rice.index[i],relevant_r_rice.columns[j])]*rail_cost.loc[relevant_r_rice.index[i]][relevant_r_rice.columns[j]]))



                L1 = list(relevant_Dict_wheat.keys())
                L2 = list(relevant_Dict_wheat.values())
                A = []
                B = []
                C = []

                df_wheat = pd.DataFrame()

                for i in range(len(L1)):
                    A.append(L1[i][0])
                    B.append(L1[i][1])
                    C.append(L2[i])

                df_wheat["From"] = A
                df_wheat["To"] = B
                df_wheat["Values"] = C

                From_state = []
                To_state = []
                Commodity = []

                for i in range(len(L1)):
                    for j in surplus_wheat.index:
                        if L1[i][0]==j:
                            From_state.append(surplus_wheat.loc[j]["State"])
                                            
                for i in range(len(L1)):
                    for j in surplus_wheat.index:
                        if L1[i][1]==j:
                            To_state.append(surplus_wheat.loc[j]["State"])


                for i in range(len(L1)):
                    Commodity.append("Wheat")

                df_wheat.insert(1,"From_state",From_state)
                df_wheat.insert(3,"To_state",To_state)
                df_wheat.insert(4,"Commodity",Commodity)
                df_wheat["Cost"] = Wheat_cost



                L3 = list(relevant_Dict_rice.keys())
                L4 = list(relevant_Dict_rice.values())

                D = []
                E = []
                F = []

                df_rice = pd.DataFrame()

                for i in range(len(L3)):
                    D.append(L3[i][0])
                    E.append(L3[i][1])
                    F.append(L4[i])

                df_rice["From"] = D
                df_rice["To"] = E
                df_rice["Values"] = F

                From_state_rice = []
                To_state_rice = []
                Commodity_rice = []

                for i in range(len(L3)):
                    for j in surplus_wheat.index:
                        if L3[i][0]==j:
                            From_state_rice.append(surplus_wheat.loc[j]["State"])
                                            
                for i in range(len(L3)):
                    for j in surplus_wheat.index:
                        if L3[i][1]==j:
                            To_state_rice.append(surplus_wheat.loc[j]["State"])



                for i in range(len(L3)):
                    Commodity_rice.append("Rice")

                df_rice.insert(1,"From_state",From_state_rice)
                df_rice.insert(3,"To_state",To_state_rice)
                df_rice.insert(4,"Commodity",Commodity_rice)
                df_rice["Cost"] = Rice_cost


                with pd.ExcelWriter("Output//List_DPT2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df_wheat.to_excel(writer, sheet_name="wheat")
                    df_rice.to_excel(writer, sheet_name="rice")

                List_data = pd.ExcelFile("Output//List_DPT2.xlsx")
                List_rice = pd.read_excel(List_data, sheet_name="rice", index_col=0)
                List_wheat = pd.read_excel(List_data, sheet_name="wheat", index_col=0)

                def extract_value_before_asterisk(data):
                    return float(data.split('*')[0])


                List_wheat["Cost"] = List_wheat["Cost"].apply(lambda x: extract_value_before_asterisk(x))
                List_rice["Cost"] = List_rice["Cost"].apply(lambda x: extract_value_before_asterisk(x))

                # Write the updated DataFrame back to the Excel file
                with pd.ExcelWriter("Output//List_DPT2.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    List_rice.to_excel(writer, sheet_name="rice")
                    List_wheat.to_excel(writer, sheet_name="wheat")
            
            else:

                data=pd.ExcelFile("Input\\Temp_balanced_DPT_scen1.xlsx")
                matrices_data = pd.ExcelFile("Input\\Matrices.xlsx")

                surplus_wheat=pd.read_excel(data,sheet_name="Surplus_wheat",index_col=1)
                deficit_wheat=pd.read_excel(data,sheet_name="Deficit_wheat",index_col=1)
                surplus_rice=pd.read_excel(data,sheet_name="Surplus_rice",index_col=1)
                deficit_rice=pd.read_excel(data,sheet_name="Deficit_rice",index_col=1)
                states_alloc=pd.read_excel(data,sheet_name="States_allocation",index_col=0)
                rail_cost = None
                if TEFD_fetched == 'NON-TEFD':
                    rail_cost=pd.read_excel("Input\\Non-TEFD.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
                elif TEFD_fetched == 'TEFD':
                    rail_cost=pd.read_excel("Input\\TEFD.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
                elif TEFD_fetched == 'Non-TEFD+TC':
                    rail_cost=pd.read_excel("Input\\Non_TEFD_TC.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)
                else:
                    rail_cost=pd.read_excel("Input\\TEFD_TC.xlsx",sheet_name="Railhead_cost_matrix",index_col=0)

                prob = LpProblem("Output\\FCI_monthly_model_allocation_rr5",LpMinimize)

                x_ij_wheat=LpVariable.dicts("x_wheat",[(i,j) for i in surplus_wheat.index for j in deficit_wheat.index],0)
                x_ij_rice=LpVariable.dicts("x_rice",[(i,j) for i in surplus_rice.index for j in deficit_rice.index],0)

                prob+=lpSum(x_ij_wheat[(i,j)]*rail_cost.loc[i][j] for i in surplus_wheat.index for j in deficit_wheat.index)+lpSum(x_ij_rice[(i,j)]*rail_cost.loc[i][j] for i in surplus_rice.index for j in deficit_rice.index)

                for i in range(len(blocked_org_rhcode)):
                    prob += x_ij_rice[(blocked_org_rhcode[i], blocked_dest_rhcode[i])] == 0
                    prob += x_ij_wheat[(blocked_org_rhcode[i], blocked_dest_rhcode[i])] == 0
                
                for i in range(len(confirmed_org_rhcode)):
                    if confirmed_railhead_commodities[i] == "RICE":
                        prob += x_ij_rice[(confirmed_org_rhcode[i], confirmed_dest_rhcode[i])] == int(confirmed_railhead_value[i])
                    else:
                        prob += x_ij_wheat[(confirmed_org_rhcode[i], confirmed_dest_rhcode[i])] == int(confirmed_railhead_value[i])

                for i in surplus_wheat.index:
                    for j in deficit_wheat.index:
                        if i==j:
                            prob+=x_ij_wheat[(i,j)]==0
                            # print(x_ij_wheat[(i,j)]==0)
                            
                for i in surplus_rice.index:
                    for j in deficit_rice.index:
                        if i==j:
                            prob+=x_ij_rice[(i,j)]==0
                            # print(x_ij_rice[(i,j)]==0)

                for i in surplus_wheat.index:
                    prob+=lpSum(x_ij_wheat[(i,j)] for j in deficit_wheat.index)<=surplus_wheat["Supply"][i]
                    
                for i in surplus_rice.index:
                    prob+=lpSum(x_ij_rice[(i,j)] for j in deficit_rice.index)<=surplus_rice["Supply"][i]



                for i in deficit_wheat.index:
                    prob+=lpSum(x_ij_wheat[(j,i)] for j in surplus_wheat.index)>=deficit_wheat["Demand"][i]
                    prob+=lpSum(x_ij_wheat[(j,i)] for j in surplus_wheat.index)<=deficit_wheat["Demand"][i]
                    
                for i in deficit_rice.index:
                    prob+=lpSum(x_ij_rice[(j,i)] for j in surplus_rice.index)>=deficit_rice["Demand"][i]
                    prob+=lpSum(x_ij_rice[(j,i)] for j in surplus_rice.index)<=deficit_rice["Demand"][i]




                prob.writeLP("FCI_monthly_model_allocation_rr.lp")
                prob.solve()
                print("Status:", LpStatus[prob.status])
                print("Minimum Cost of Transportation = Rs.", value(prob.objective),"Lakh")
                print("Total Number of Variables:",len(prob.variables()))
                print("Total Number of Constraints:",len(prob.constraints))


                r_wheat={}
                r_wheat=pd.DataFrame([],index=surplus_wheat.index,columns=deficit_wheat.index)
                    
                for r in surplus_wheat.index:
                    for j in deficit_wheat.index:
                        r_wheat.loc[r][j]=x_ij_wheat[(r,j)].value()
                        
                r_rice={}
                r_rice=pd.DataFrame([],index=surplus_rice.index,columns=deficit_rice.index)
                    
                for r in surplus_rice.index:
                    for j in deficit_rice.index:
                        r_rice.loc[r][j]=x_ij_rice[(r,j)].value()
                        
                with pd.ExcelWriter("Output\\Results_DPT.xlsx",mode='a',engine='openpyxl', if_sheet_exists='replace') as writer:
                    r_wheat.to_excel(writer,sheet_name="r_wheat",float_format="%0.3f")
                    r_rice.to_excel(writer,sheet_name="r_rice",float_format="%0.3f")


                relevant_data = pd.ExcelFile("Output//Results_DPT.xlsx")
                relevant_r_wheat = pd.read_excel(relevant_data, sheet_name="r_wheat", index_col=0)
                relevant_r_rice = pd.read_excel(relevant_data, sheet_name="r_rice", index_col=0)
                relevant_Dict_wheat = {}
                relevant_Dict_rice = {}
                Rice_cost = []
                Wheat_cost = []

                for i in range(len(relevant_r_wheat.index)):
                    for j in range(len(relevant_r_wheat.columns)):
                        if relevant_r_wheat.iat[i, j] > 0:
                            relevant_Dict_wheat[relevant_r_wheat.index[i], relevant_r_wheat.columns[j]] = relevant_r_wheat.iloc[i][relevant_r_wheat.columns[j]]
                            Wheat_cost.append((x_ij_wheat[(relevant_r_wheat.index[i],relevant_r_wheat.columns[j])]*rail_cost.loc[relevant_r_wheat.index[i]][relevant_r_wheat.columns[j]]))

                for i in range(len(relevant_r_rice.index)):
                    for j in range(len(relevant_r_rice.columns)):
                        if relevant_r_rice.iat[i, j] > 0:
                            relevant_Dict_rice[relevant_r_rice.index[i], relevant_r_rice.columns[j]] = relevant_r_rice.iloc[i][relevant_r_rice.columns[j]]
                            Rice_cost.append((x_ij_wheat[(relevant_r_rice.index[i],relevant_r_rice.columns[j])]*rail_cost.loc[relevant_r_rice.index[i]][relevant_r_rice.columns[j]]))



                L1 = list(relevant_Dict_wheat.keys())
                L2 = list(relevant_Dict_wheat.values())
                A = []
                B = []
                C = []

                df_wheat = pd.DataFrame()

                for i in range(len(L1)):
                    A.append(L1[i][0])
                    B.append(L1[i][1])
                    C.append(L2[i])

                df_wheat["From"] = A
                df_wheat["To"] = B
                df_wheat["Values"] = C

                From_state = []
                To_state = []
                Commodity = []

                for i in range(len(L1)):
                    for j in surplus_wheat.index:
                        if L1[i][0]==j:
                            From_state.append(surplus_wheat.loc[j]["State"])
                                            
                for i in range(len(L1)):
                    for j in surplus_wheat.index:
                        if L1[i][1]==j:
                            To_state.append(surplus_wheat.loc[j]["State"])


                for i in range(len(L1)):
                    Commodity.append("Wheat")

                df_wheat.insert(1,"From_state",From_state)
                df_wheat.insert(3,"To_state",To_state)
                df_wheat.insert(4,"Commodity",Commodity)
                df_wheat["Cost"] = Wheat_cost



                L3 = list(relevant_Dict_rice.keys())
                L4 = list(relevant_Dict_rice.values())

                D = []
                E = []
                F = []

                df_rice = pd.DataFrame()

                for i in range(len(L3)):
                    D.append(L3[i][0])
                    E.append(L3[i][1])
                    F.append(L4[i])

                df_rice["From"] = D
                df_rice["To"] = E
                df_rice["Values"] = F

                From_state_rice = []
                To_state_rice = []
                Commodity_rice = []

                for i in range(len(L3)):
                    for j in surplus_wheat.index:
                        if L3[i][0]==j:
                            From_state_rice.append(surplus_wheat.loc[j]["State"])
                                            
                for i in range(len(L3)):
                    for j in surplus_wheat.index:
                        if L3[i][1]==j:
                            To_state_rice.append(surplus_wheat.loc[j]["State"])



                for i in range(len(L3)):
                    Commodity_rice.append("Rice")

                df_rice.insert(1,"From_state",From_state_rice)
                df_rice.insert(3,"To_state",To_state_rice)
                df_rice.insert(4,"Commodity",Commodity_rice)
                df_rice["Cost"] = Rice_cost


                with pd.ExcelWriter("Output//List_DPT.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df_wheat.to_excel(writer, sheet_name="wheat")
                    df_rice.to_excel(writer, sheet_name="rice")

                List_data = pd.ExcelFile("Output//List_DPT.xlsx")
                List_rice = pd.read_excel(List_data, sheet_name="rice", index_col=0)
                List_wheat = pd.read_excel(List_data, sheet_name="wheat", index_col=0)

                def extract_value_before_asterisk(data):
                    return float(data.split('*')[0])


                List_wheat["Cost"] = List_wheat["Cost"].apply(lambda x: extract_value_before_asterisk(x))
                List_rice["Cost"] = List_rice["Cost"].apply(lambda x: extract_value_before_asterisk(x))

                # Write the updated DataFrame back to the Excel file
                with pd.ExcelWriter("Output//List_DPT.xlsx", mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    List_rice.to_excel(writer, sheet_name="rice")
                    List_wheat.to_excel(writer, sheet_name="wheat")

                    

               
            data1["status"] = 1
                  
        except Exception as e:
            print(e)
            data1["status"] = 0
        json_data = json.dumps(data1)
        json_object = json.loads(json_data)

        return(json.dumps(json_object, indent = 1))
    else:
        return ("error")



@app.route("/Alternate_Railhead_Solve",methods = ["POST","GET"])
def Alternate_Railhead_Solve():
    data = request.get_json()
    rh_source = data['rh_source']
    rh_dest = data['rh_dest']
    # zone = data['zone']
    # n = data['n']
    Alternate_Railhead_source = rh_source.upper()
    Alternate_Railhead_Destination = rh_dest.upper()
    # Alternate_Railhead_zone = zone
    # Alternate_Railhead_increment = 0.8
    data1 = {}
    if request.method == "POST":
        try:
            file = pd.ExcelFile("Input\\Daily_Template_Scene1.xlsx")
            matrices_data = pd.ExcelFile("Input\\Non-TEFD.xlsx")
            surplus_wheat = pd.read_excel(file, sheet_name="Surplus_wheat", index_col=0)
            rail_cost = pd.read_excel(matrices_data, sheet_name="Railhead_cost_matrix", index_col=0)
            alt_rh_state = surplus_wheat.loc[Alternate_Railhead_Destination]["State"]

            lst1 = []

            for index, row in surplus_wheat.iterrows():
                if row["State"] == alt_rh_state:
                    lst1.append(index)

            lst2 = []

            for j in lst1:
                lst2.append(rail_cost.loc[Alternate_Railhead_source, j])

            keys = lst1
            values = lst2

            dict_altrh = dict(zip(keys, values))

            threshold = rail_cost.loc[Alternate_Railhead_source, Alternate_Railhead_Destination]
            filt_dict_altrh = {k: v for k, v in dict_altrh.items() if k != Alternate_Railhead_Destination and v >= threshold}
            sort_dict_altrh = dict(sorted(filt_dict_altrh.items(), key=lambda item: item[1]))
            top_3_elements = list(sort_dict_altrh.items())[:3]
            result_altrh = []

            for i in range(len(top_3_elements)):
                result_altrh.append(top_3_elements[i][0])

            with open('Output\\Alternate_Railhead.pkl', 'wb') as f:
                pickle.dump(result_altrh, f)
                        
            data1["status"] = 1
                  
        except Exception as e:
            print(e)
            data1["status"] = 0
        json_data = json.dumps(data1)
        json_object = json.loads(json_data)

        return(json.dumps(json_object, indent = 1))
    else:
        return ("error")




if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
