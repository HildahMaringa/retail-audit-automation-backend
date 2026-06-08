from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pathlib import Path
from datetime import datetime
import shutil
import uuid
import pandas as pd

app = FastAPI(title="Retail Audit Automation Backend")

# Frontend dev + deployed frontend URLs
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "https://retail-audit-automation-ui.vercel.app",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


@app.get("/")
def health_check():
    return {
        "status": "ok",
        "service": "Retail Audit Automation Backend",
    }


@app.post("/api/runs")
async def create_run(
    query_family: str = Form(...),
    project: str = Form(...),
    action: str = Form(...),
    month: str = Form(...),
    year: str = Form(...),
    batch: str = Form(...),
    data_file: UploadFile = File(...),
    feedback_file: UploadFile | None = File(None),
):
    run_id = str(uuid.uuid4())

    run_upload_dir = UPLOAD_DIR / run_id
    run_output_dir = OUTPUT_DIR / run_id

    run_upload_dir.mkdir(parents=True, exist_ok=True)
    run_output_dir.mkdir(parents=True, exist_ok=True)

    data_file_path = run_upload_dir / data_file.filename

    with data_file_path.open("wb") as buffer:
        shutil.copyfileobj(data_file.file, buffer)

    feedback_file_path = None
    if feedback_file:
        feedback_file_path = run_upload_dir / feedback_file.filename
        with feedback_file_path.open("wb") as buffer:
            shutil.copyfileobj(feedback_file.file, buffer)

    # Dummy output for now.
    # Later this is where we will call the real Python scripts.
    output_file_name = f"{project}-{action.replace(' ', '-')}-{month}{year}-Batch{batch}-Output.xlsx"
    output_file_path = run_output_dir / output_file_name

    summary_df = pd.DataFrame(
        [
            {
                "Run ID": run_id,
                "Query Family": query_family,
                "Project": project,
                "Action": action,
                "Month": month,
                "Year": year,
                "Batch": batch,
                "Data File": data_file.filename,
                "Feedback File": feedback_file.filename if feedback_file else "",
                "Status": "Success",
                "Created At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        ]
    )

    summary_df.to_excel(output_file_path, index=False)

    return {
        "run_id": run_id,
        "status": "success",
        "message": "Run completed successfully.",
        "output_file": output_file_name,
        "download_url": f"/api/runs/{run_id}/download",
    }


@app.get("/api/runs/{run_id}/download")
def download_output(run_id: str):
    run_output_dir = OUTPUT_DIR / run_id

    if not run_output_dir.exists():
        return {
            "status": "error",
            "message": "Output folder not found.",
        }

    output_files = list(run_output_dir.glob("*.xlsx"))

    if not output_files:
        return {
            "status": "error",
            "message": "No output file found for this run.",
        }

    output_file = output_files[0]

    return FileResponse(
        path=output_file,
        filename=output_file.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )