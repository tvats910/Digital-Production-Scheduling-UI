import os
import pandas as pd
import numpy as np
import json
import os
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse, FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from datetime import date
from pydantic import BaseModel
from aps.aps_engine import run_aps
from fastapi.staticfiles import StaticFiles

app = FastAPI()

base_dir = os.path.dirname(os.path.abspath(__file__))
frontend_dir = os.path.join(base_dir, "frontend")

# Mount the frontend folder
app.mount("/frontend", StaticFiles(directory=frontend_dir), name="frontend")


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- MODELS ---
class MachineUpdate(BaseModel):
    machine: str
    last_part: str

# --- GLOBAL STORAGE ---
# Persists data in memory so Analytics page can access it
LAST_RESULT = None
CURRENT_SECTION = "VT" 

@app.post("/upload")
async def upload_files(
    section: str = Form(...),
    book: UploadFile = File(...),
    inventory: UploadFile = File(...),
    changeover: UploadFile = File(...),
    matrix: UploadFile = File(...),
    terminal: UploadFile = File(...),
    demand: UploadFile = File(...)
):
    global LAST_RESULT, CURRENT_SECTION
    CURRENT_SECTION = section
    
    os.makedirs("uploads", exist_ok=True)
    os.makedirs("outputs", exist_ok=True)
    
    def save(f, custom_name=None):
        name = custom_name if custom_name else f.filename
        p = os.path.join("uploads", name)
        with open(p, "wb") as buffer:
            buffer.write(f.file.read())
        return p

    try:
        # 1. Save files
        # We rename them slightly to ensure the engine recognizes CSV vs Excel
        b_path = save(book)
        i_path = save(inventory)
        m_path = save(matrix)
        c_path = save(changeover)
        t_path = save(terminal)
        d_path = save(demand)

        # 2. PRE-FLIGHT FIX: Fix the Demand Column Header automatically
        # This prevents the find_col error we identified earlier
        if d_path.endswith('.csv'):
            temp_df = pd.read_csv(d_path)
            if 'Demand' in temp_df.columns and 'Daily_Demand' not in temp_df.columns:
                temp_df.rename(columns={'Demand': 'Daily_Demand'}, inplace=True)
                temp_df.to_csv(d_path, index=False)
        elif d_path.endswith(('.xlsx', '.xls')):
            # If it's an Excel, we try to fix the specific sheet
            with pd.ExcelWriter(d_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                temp_df = pd.read_excel(d_path, sheet_name=f"{section}_Demand")
                if 'Demand' in temp_df.columns:
                    temp_df.rename(columns={'Demand': 'Daily_Demand'}, inplace=True)
                    temp_df.to_excel(writer, sheet_name=f"{section}_Demand", index=False)

        # 3. Call the Engine
        result = run_aps(
            planning_date=date.today(),
            indent_month=date.today().replace(day=1),
            book_path=b_path,
            inventory_path=i_path,
            matrix_path=m_path,
            changeover_path=c_path,
            terminal_path=t_path,
            demand_path=d_path,
            machine_state_path="machine_state.json",
            output_path=f"outputs/aps_{section}_output.xlsx",
            section=section
        )

        LAST_RESULT = result

        if not result.get("success"):
            return JSONResponse(content={"error": result.get("error")})

        # --- DYNAMIC KEY EXTRACTION ---
        # If section is VT, extracts 'VT_Plan'. If HZ, extracts 'HZ_Plan'
        plan_key = f"{section}_Plan"
        skip_key = f"{section}_Not_Planned"

        plan_data = result["sheets"].get(plan_key, [])
        deferred_data = result["sheets"].get(skip_key, [])

        # Fallback in case of naming mismatch in engine
        if not plan_data and "Plan" in result["sheets"]:
            plan_data = result["sheets"].get("Plan", [])

        df = pd.DataFrame(plan_data)
        
        if df.empty:
            return JSONResponse(content={
                "plan": [], 
                "deferred": deferred_data, 
                "section": section,
                "message": f"Optimization finished for {section}, but 0 production rows were generated."
            })

        # Standard grouping logic for Dashboard Table
        grouped = df.groupby(["Machine", "Part"], as_index=False).agg({"Production_Qty": "sum"})
        grouped.columns = ["Machine", "Part", "Qty_Planned"]
        grouped["Total_Machine_Qty"] = grouped.groupby("Machine")["Qty_Planned"].transform("sum")

        return JSONResponse(content={
            "plan": grouped.replace({np.nan: None}).to_dict(orient="records"),
            "deferred": deferred_data,
            "section": section
        })

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return JSONResponse(content={"error": str(e)}, status_code=500)

@app.get("/utilization-last")
async def get_utilization_last():
    global LAST_RESULT, CURRENT_SECTION
    if not LAST_RESULT:
        return JSONResponse(content={"error": "No run data found. Run Optimizer first."})

    try:
        # Dynamically fetch (e.g., HZ_Machine_Util or VT_Machine_Util)
        util_key = f"{CURRENT_SECTION}_Machine_Util"
        util_rows = LAST_RESULT["sheets"].get(util_key, [])
        df_util = pd.DataFrame(util_rows)

        if df_util.empty:
            return JSONResponse(content={"utilization": [], "idle_machines": [], "section": CURRENT_SECTION})

        # Map 'Utilization_%' to a standard 'Utilization_Pct' for JavaScript
        if "Utilization_%" in df_util.columns:
            df_util["Utilization_Pct"] = df_util["Utilization_%"]
        
        active = df_util[df_util["Used_Hours"] > 0].replace({np.nan: 0}).to_dict(orient="records")
        idle = df_util[df_util["Used_Hours"] <= 0]["Machine"].tolist()

        return JSONResponse(content={
            "utilization": active,
            "idle_machines": idle,
            "section": CURRENT_SECTION
        })
    except Exception as e:
        return JSONResponse(content={"error": str(e)})

@app.post("/update-bulk-state")
async def update_bulk_state(payload: dict):
    file_path = "machine_state.json"
    try:
        existing_state = {}
        if os.path.exists(file_path):
            with open(file_path, "r") as f:
                try:
                    existing_state = json.load(f)
                except:
                    existing_state = {}
        
        # Merge new states from UI setup grid into JSON
        existing_state.update(payload)

        with open(file_path, "w") as f:
            json.dump(existing_state, f, indent=4)
            
        return {"status": "success"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download")
def download():
    # Downloads the last generated file for the active section
    target_file = f"outputs/aps_{CURRENT_SECTION}_output.xlsx"
    if os.path.exists(target_file):
        return FileResponse(target_file)
    return FileResponse("outputs/aps_output.xlsx")

@app.post("/update-state")
async def update_state(data: MachineUpdate):
    file_path = "machine_state.json"
    state = {}
    if os.path.exists(file_path):
        with open(file_path, "r") as f:
            try: state = json.load(f)
            except: state = {}
    state[data.machine] = data.last_part
    with open(file_path, "w") as f:
        json.dump(state, f, indent=4)
    return {"status": "success", "updated": data.machine}

import io

@app.get("/download-sheet")
async def download_sheet(sheet_name: str):
    global LAST_RESULT, CURRENT_SECTION
    if not LAST_RESULT:
        raise HTTPException(status_code=400, detail="No data found")

    # Construct the internal key (e.g., VT_Plan or HZ_Plan)
    internal_key = f"{CURRENT_SECTION}_{sheet_name}"
    data = LAST_RESULT["sheets"].get(internal_key, [])
    
    if not data:
        # Fallback if the engine didn't use a prefix for some reason
        data = LAST_RESULT["sheets"].get(sheet_name, [])

    df = pd.DataFrame(data)
    
    # Create CSV in memory
    stream = io.StringIO()
    df.to_csv(stream, index=False)
    
    return StreamingResponse(
        iter([stream.getvalue()]),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={internal_key}.csv"}
    )
# Also keep your main Excel download
@app.get("/download-main")
def download_main():
    target_file = f"outputs/aps_{CURRENT_SECTION}_output.xlsx"
    if os.path.exists(target_file):
        return FileResponse(target_file, filename=f"Smart_APS_{CURRENT_SECTION}_Full_Report.xlsx")
    raise HTTPException(status_code=404, detail="File not found")