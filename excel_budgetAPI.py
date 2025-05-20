from fastapi import FastAPI, File, UploadFile
import shutil
from excel_budget import process_budget_excel

app = FastAPI()

@app.post("/upload-budget/")
async def upload_budget(file: UploadFile = File(...)):
    temp_file = f"temp_{file.filename}"
    with open(temp_file, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    output_file = process_budget_excel(temp_file)
    return {"message": "Traitement terminé", "output": output_file}
