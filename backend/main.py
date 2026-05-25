from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/api/hello")
def read_root():
    return {"message": "Python API is running!"}

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        filename = file.filename.lower()
        
        if filename.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(contents))
        elif filename.endswith('.xlsx') or filename.endswith('.xls'):
            df = pd.read_excel(io.BytesIO(contents))
        elif filename.endswith('.json'):
            df = pd.read_json(io.BytesIO(contents))
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format. Please upload CSV, Excel, or JSON.")
            
        # Clean up column names by stripping whitespace
        df.columns = df.columns.str.strip()
        
        # Replace NaN with empty strings
        df = df.fillna("")
        
        # Convert to list of dictionaries
        records = df.to_dict(orient="records")
        columns = list(df.columns)
        
        return {
            "filename": file.filename,
            "rows": len(records),
            "cols": len(columns),
            "columns": columns,
            "data": records
        }
        
    except Exception as e:
        print(f"Error parsing file: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error parsing data: {str(e)}")
