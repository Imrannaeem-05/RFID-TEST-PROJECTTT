from flask import Flask, render_template, request, redirect, url_for, flash, session
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = 'ambalabu'

EXCEL_FILE = "rfidexcelnew.xlsx"


if not os.path.exists(EXCEL_FILE):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        pd.DataFrame(columns=["CardID", "Name"]).to_excel(writer, sheet_name='UserList', index=False)
        pd.DataFrame(columns=["CardID", "Name", "ToolID", "ToolName", "Action", "Condition", "Timestamp"]).to_excel(writer, sheet_name='UserLog', index=False)
        pd.DataFrame(columns=["ToolID", "ToolName"]).to_excel(writer, sheet_name='ToolList', index=False)


@app.route("/") #this the place where da system give options to borrow or return fr
def welcome():
    return render_template("welcome.html")


@app.route("/borrow", methods=["GET", "POST"]) #after press borrow button gang
def borrow():
    message = None
    color = None
    if request.method == "POST":
        user_id = request.form.get("user_id", "").strip().upper()

        
        df_users = pd.read_excel(EXCEL_FILE, sheet_name="UserList")
        df_users.columns = df_users.columns.str.strip()
        df_users["CardID"] = df_users["CardID"].astype(str).str.strip()

        
        match = df_users[df_users["CardID"] == user_id]

        if match.empty:
            message = "Access Denied"
            color = "red"
            return render_template("home.html", message=message, color=color)

        user_name = match.iloc[0]["Name"]
        return redirect(url_for("tool", user_id=user_id, name=user_name))

    return render_template("home.html", message=message, color=color)


@app.route("/tool", methods=["GET", "POST"])
def tool():
    user_id = request.args.get("user_id", "")
    user_name = request.args.get("name", "")
    message = None
    color = None

    if request.method == "POST":
        tool_id = request.form.get("tool_id", "").strip().upper()

        if not tool_id:
            message = "Please scan a tool first."
            color = "red"
            return render_template("tool.html", user_id=user_id, name=user_name, message=message, color=color)

        df_tools = pd.read_excel(EXCEL_FILE, sheet_name="ToolList")
        df_tools.columns = df_tools.columns.str.strip()
        df_tools["ToolID"] = df_tools["ToolID"].astype(str).str.strip()

        tool_match = df_tools[df_tools["ToolID"] == tool_id]

        if tool_match.empty:
            message = "ToolID not found!"
            color = "red"
            return render_template("tool.html", user_id=user_id, name=user_name, message=message, color=color)

        tool_name = tool_match.iloc[0]["ToolName"]

        df_log = pd.read_excel(EXCEL_FILE, sheet_name="UserLog")
        df_log.columns = df_log.columns.str.strip()

        last_action = df_log[df_log["ToolID"] == tool_id].tail(1)

        if not last_action.empty and last_action.iloc[0]["Action"] == "BORROW":
            message = f"Tool has not been returned!"
            color = "red"
            return render_template("tool.html", user_id=user_id, name=user_name, message=message, color=color)

        new_entry = pd.DataFrame([{
            "CardID": user_id,
            "Name": user_name,
            "ToolID": tool_id,
            "ToolName": tool_name,
            "Action": "BORROW",
            "Condition": "-",
            "Timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        }])

        df_log = pd.concat([df_log, new_entry], ignore_index=True)
        with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
            df_log.to_excel(writer, sheet_name="UserLog", index=False)

        message = f"{tool_name} borrowed successfully!"
        color = "green"
        return render_template("tool.html", user_id=user_id, name=user_name, message=message, color=color)

    return render_template("tool.html", user_id=user_id, name=user_name, message=message, color=color)



@app.route("/return", methods=["GET", "POST"])
def return_tool():
    message = None

    if request.method == "POST":
        tool_id = request.form.get("tool_id", "").strip().upper()
        condition = request.form.get("condition", "Good")

        if not tool_id:
            message = "Please scan a tool first."
            color = "red"
            return render_template("return.html", message=message , color=color)

        df_tools = pd.read_excel(EXCEL_FILE, sheet_name="ToolList")
        df_tools.columns = df_tools.columns.str.strip()
        df_tools["ToolID"] = df_tools["ToolID"].astype(str).str.strip()

        tool_match = df_tools[df_tools["ToolID"] == tool_id]

        if tool_match.empty:
            message = "ToolID not found!"
            color = "red"
            return render_template("return.html", message=message , color=color)

        tool_name = tool_match.iloc[0]["ToolName"]

        df_log = pd.read_excel(EXCEL_FILE, sheet_name="UserLog")
        df_log.columns = df_log.columns.str.strip()

        last_action = df_log[df_log["ToolID"] == tool_id].tail(1)

        if last_action.empty:
            message = "This tool has not been borrowed."
            color = "red"
            return render_template("return.html", message=message, color=color)

        if last_action.iloc[0]["Action"] == "RETURN":
            message = "This tool has already been returned."
            color = "red"
            return render_template("return.html", message=message, color=color)

        borrower_id = last_action.iloc[0]["CardID"]
        borrower_name = last_action.iloc[0]["Name"]

        new_entry = pd.DataFrame([{
            "CardID": borrower_id,
            "Name": borrower_name,
            "ToolID": tool_id,
            "ToolName": tool_name,
            "Action": "RETURN",
            "Condition": condition,
            "Timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        }])

        df_log = pd.concat([df_log, new_entry], ignore_index=True)
        with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
            df_log.to_excel(writer, sheet_name="UserLog", index=False)

        message = f"{tool_name} returned successfully!"
        color = "green"
        return render_template("return.html", message=message, color=color)

    return render_template("return.html", message=message)

@app.route("/logs") #log table la apa lagi
def show_logs():
    df_log = pd.read_excel(EXCEL_FILE, sheet_name="UserLog")
    df_log.columns = df_log.columns.str.strip()
    
    df_log = df_log.drop(columns=["ToolID"]) #dflog.drop is to hide from the flask, (but is saved in excel)
    df_log = df_log.drop(columns=["CardID"] )
    df_log = df_log[["Name", "ToolName", "Action", "Condition", "Timestamp"]]
    
    
    for col in df_log.columns:
        if df_log[col].dtype == 'object':
            df_log[col] = df_log[col].astype(str).str.strip()
    
    logs = df_log.to_dict('records')
    
    return render_template("log.html", logs=logs) 


@app.route("/clear_logs")
def clear_logs():
    df_log = pd.DataFrame(columns=["CardID", "Name", "ToolID", "ToolName", "Action", "Condition", "Timestamp"])
    with pd.ExcelWriter(EXCEL_FILE, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        df_log.to_excel(writer, sheet_name="UserLog", index=False)
    flash("All logs cleared successfully.")
    return redirect(url_for("show_logs"))

@app.route("/tooltab") #when press tools on the navigation bar
def tool_tab():
    df_tools = pd.read_excel(EXCEL_FILE, sheet_name="ToolList")
    df_tools.columns = df_tools.columns.str.strip()

    df_log= pd.read_excel(EXCEL_FILE, sheet_name="UserLog")
    df_log.columns = df_log.columns.str.strip()

    tools_list =[]

    for idx, tool in df_tools.iterrows():
        tool_id = str(tool["ToolID"]).strip()
        tool_name = tool["ToolName"]

        last_action = df_log[df_log["ToolID"] == tool_id].tail(1)

        if last_action.empty:
            status = "IN STORAGE"
            operator = "-"
            condition = "-"

        elif last_action.iloc[0]["Action"] == "RETURN":
            status = "IN STORAGE"
            operator = "-"
            condition = last_action.iloc[0]["Condition"]
        else:
            status = "BORROWED"
            operator = last_action.iloc[0]["Name"]
            condition = "-"

        tools_list.append({
            "number": idx + 1,
            "tool_id": tool_id,
            "tool_name": tool_name,
            "status": status,
            "operator": operator,
            "condition": condition
        })    

    return render_template("tooltab.html", tools=tools_list)



@app.route("/adminlogin", methods=["GET", "POST"])
def admin_login():
    message = None
    color = None

    if request.method == "POST":
        card_id = request.form.get("user_id", "").strip().upper()
        
        if card_id == "E20047124741662E032E737C": #card putih (user card)
            session['admin'] = True
            flash("Welcome, Admin!", "success")
            return redirect("/logs")
        else:
            message = "Access Denied — Admin Only"
            color = "red"

    return render_template("adminlogin.html", message=message, color=color)

    


if __name__ == "__main__":
    app.run(debug=True)