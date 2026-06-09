from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pathlib import Path
from datetime import datetime
import shutil
import uuid
import zipfile
import traceback
import pandas as pd

from scripts.data_queries_engine import run_data_queries
from scripts.data_queries_feedback_correction import apply_feedback_corrections_ai_agent
from scripts.unified_cooler_pos_queries import run_unified_cooler_pos_queries
from scripts.unified_cooler_pos_feedback_correction import apply_unified_feedback_corrections


app = FastAPI(title="Retail Audit Automation Backend")

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


def package_generated_outputs(
    run_output_dir: Path,
    before_files: set,
    project: str,
    action: str,
    month: str,
    year: str,
    batch: str,
):
    after_files = set(run_output_dir.glob("*"))

    new_files = sorted(
        [path for path in list(after_files - before_files) if path.is_file()],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )

    if not new_files:
        raise RuntimeError("The script completed but did not generate an output file.")

    excel_files = [
        path
        for path in new_files
        if path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]
    ]

    if len(excel_files) > 1:
        output_file_name = (
            f"{project}-{action.replace(' ', '-')}-{month}{year}-Batch{batch}-Outputs.zip"
        )
        output_file_path = run_output_dir / output_file_name

        with zipfile.ZipFile(output_file_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file_path in excel_files:
                zip_file.write(file_path, arcname=file_path.name)

        return output_file_path, output_file_name

    if len(excel_files) == 1:
        output_file_path = excel_files[0]
        return output_file_path, output_file_path.name

    output_file_path = new_files[0]
    return output_file_path, output_file_path.name


def package_specific_files(
    run_output_dir: Path,
    files: list,
    project: str,
    action: str,
    month: str,
    year: str,
    batch: str,
):
    existing_files = [
        Path(file_path)
        for file_path in files
        if file_path and Path(file_path).exists()
    ]

    if not existing_files:
        raise RuntimeError("The script completed but did not return any existing output files.")

    if len(existing_files) > 1:
        output_file_name = (
            f"{project}-{action.replace(' ', '-')}-{month}{year}-Batch{batch}-Outputs.zip"
        )
        output_file_path = run_output_dir / output_file_name

        with zipfile.ZipFile(output_file_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file_path in existing_files:
                zip_file.write(file_path, arcname=file_path.name)

        return output_file_path, output_file_name

    output_file_path = existing_files[0]
    return output_file_path, output_file_path.name


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

    try:
        if query_family == "Data Queries" and action == "Run Queries":
            before_files = set(run_output_dir.glob("*"))

            run_data_queries(
                project_name=project,
                data_path=str(data_file_path),
                output_dir=str(run_output_dir),
                prev_feedback_path=str(feedback_file_path) if feedback_file_path else None,
            )

            output_file_path, output_file_name = package_generated_outputs(
                run_output_dir=run_output_dir,
                before_files=before_files,
                project=project,
                action=action,
                month=month,
                year=year,
                batch=batch,
            )

        elif query_family == "Data Queries" and action == "Correct Feedback":
            if not feedback_file_path:
                raise RuntimeError("Feedback workbook is required for Data Queries feedback correction.")

            corrected_data_file, corrected_feedback_file = apply_feedback_corrections_ai_agent(
                project_name=project,
                data_path=str(data_file_path),
                feedback_path=str(feedback_file_path),
                output_dir=str(run_output_dir),
            )

            output_file_path, output_file_name = package_specific_files(
                run_output_dir=run_output_dir,
                files=[corrected_data_file, corrected_feedback_file],
                project=project,
                action=action,
                month=month,
                year=year,
                batch=batch,
            )

        elif query_family == "POS and Cooler Queries" and action == "Run Queries":
            before_files = set(run_output_dir.glob("*"))

            run_unified_cooler_pos_queries(
                project_name=project,
                input_file=str(data_file_path),
                output_dir=str(run_output_dir),
            )

            output_file_path, output_file_name = package_generated_outputs(
                run_output_dir=run_output_dir,
                before_files=before_files,
                project=project,
                action=action,
                month=month,
                year=year,
                batch=batch,
            )

        elif query_family == "POS and Cooler Queries" and action == "Correct Feedback":
            if not feedback_file_path:
                raise RuntimeError("Feedback workbook is required for POS and Cooler feedback correction.")

            before_files = set(run_output_dir.glob("*"))

            corrected_file = apply_unified_feedback_corrections(
                project_name=project,
                data_path=str(data_file_path),
                feedback_path=str(feedback_file_path),
                output_dir=str(run_output_dir),
            )

            if corrected_file:
                output_file_path = Path(corrected_file)
                output_file_name = output_file_path.name
            else:
                output_file_path, output_file_name = package_generated_outputs(
                    run_output_dir=run_output_dir,
                    before_files=before_files,
                    project=project,
                    action=action,
                    month=month,
                    year=year,
                    batch=batch,
                )

        else:
            output_file_name = (
                f"{project}-{action.replace(' ', '-')}-{month}{year}-Batch{batch}-Output.xlsx"
            )
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

    except Exception as exc:
        print("❌ RUN FAILED")
        traceback.print_exc()

        raise HTTPException(
            status_code=500,
            detail=f"Run failed: {str(exc)}",
        )

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
        raise HTTPException(
            status_code=404,
            detail="Output folder not found.",
        )

    output_files = list(run_output_dir.glob("*.zip"))

    if not output_files:
        output_files = list(run_output_dir.glob("*.xlsx"))

    if not output_files:
        raise HTTPException(
            status_code=404,
            detail="No output file found for this run.",
        )

    output_file = sorted(
        output_files,
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )[0]

    return FileResponse(
        path=output_file,
        filename=output_file.name,
        media_type=(
            "application/zip"
            if output_file.suffix.lower() == ".zip"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )