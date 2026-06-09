from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pathlib import Path
from datetime import datetime
from typing import Optional
import shutil
import uuid
import zipfile
import traceback
import pandas as pd

from scripts.data_queries_engine import run_data_queries
from scripts.data_queries_feedback_correction import apply_feedback_corrections_ai_agent
from scripts.unified_cooler_pos_queries import run_unified_cooler_pos_queries
from scripts.unified_cooler_pos_feedback_correction import apply_unified_feedback_corrections
import scripts.feedback_batch_merger as feedback_batch_merger


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

# Simple in-memory run tracker.
# This is okay for now on one Render instance.
RUNS = {}


@app.get("/")
def health_check():
    return {
        "status": "ok",
        "service": "Retail Audit Automation Backend",
    }


def save_upload_file(upload_file: UploadFile, destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)

    with destination.open("wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)

    return destination


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


def process_run(
    run_id: str,
    query_family: str,
    project: str,
    action: str,
    month: str,
    year: str,
    batch: str,
    data_file_path: Optional[str],
    feedback_file_path: Optional[str],
    query_folder_path: Optional[str],
    feedback_folder_path: Optional[str],
    data_file_name: Optional[str],
    feedback_file_name: Optional[str],
):
    run_output_dir = OUTPUT_DIR / run_id
    run_output_dir.mkdir(parents=True, exist_ok=True)

    RUNS[run_id]["status"] = "running"
    RUNS[run_id]["message"] = "Processing started."
    RUNS[run_id]["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        data_path = Path(data_file_path) if data_file_path else None
        feedback_path = Path(feedback_file_path) if feedback_file_path else None

        if query_family == "Data Queries" and action == "Run Queries":
            if not data_path:
                raise RuntimeError("Data workbook is required for Data Queries.")

            before_files = set(run_output_dir.glob("*"))

            run_data_queries(
                project_name=project,
                data_path=str(data_path),
                output_dir=str(run_output_dir),
                prev_feedback_path=str(feedback_path) if feedback_path else None,
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
            if not data_path:
                raise RuntimeError("Data workbook is required for Data Queries feedback correction.")
            if not feedback_path:
                raise RuntimeError("Feedback workbook is required for Data Queries feedback correction.")

            corrected_data_file, corrected_feedback_file = apply_feedback_corrections_ai_agent(
                project_name=project,
                data_path=str(data_path),
                feedback_path=str(feedback_path),
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
            if not data_path:
                raise RuntimeError("Data workbook is required for POS and Cooler Queries.")

            before_files = set(run_output_dir.glob("*"))

            run_unified_cooler_pos_queries(
                project_name=project,
                input_file=str(data_path),
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
            if not data_path:
                raise RuntimeError("Data workbook is required for POS and Cooler feedback correction.")
            if not feedback_path:
                raise RuntimeError("Feedback workbook is required for POS and Cooler feedback correction.")

            before_files = set(run_output_dir.glob("*"))

            corrected_file = apply_unified_feedback_corrections(
                project_name=project,
                data_path=str(data_path),
                feedback_path=str(feedback_path),
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

        elif action == "Merge Feedback":
            if not query_folder_path:
                raise RuntimeError("Query files are required for Merge Feedback.")
            if not feedback_folder_path:
                raise RuntimeError("Feedback files are required for Merge Feedback.")

            query_folder = Path(query_folder_path)
            feedback_folder = Path(feedback_folder_path)

            query_files_found = list(query_folder.glob("*.xlsx")) + list(query_folder.glob("*.xlsm")) + list(query_folder.glob("*.xls"))
            feedback_files_found = list(feedback_folder.glob("*.xlsx")) + list(feedback_folder.glob("*.xlsm")) + list(feedback_folder.glob("*.xls"))

            if not query_files_found:
                raise RuntimeError("No query Excel files were uploaded for Merge Feedback.")
            if not feedback_files_found:
                raise RuntimeError("No feedback Excel files were uploaded for Merge Feedback.")

            output_file_name_setting = f"{project}-Merged-Feedback-{month}{year}-Batch{batch}.xlsx"

            # The merger script is folder-based, so we set its config dynamically
            # for this one backend run.
            feedback_batch_merger.PROJECT_NAME = project
            feedback_batch_merger.QUERY_FOLDER = str(query_folder)
            feedback_batch_merger.FEEDBACK_FOLDER = str(feedback_folder)
            feedback_batch_merger.OUTPUT_DIR = str(run_output_dir)
            feedback_batch_merger.OUTPUT_FILE_NAME = output_file_name_setting

            # Let the script check common batches and skip missing ones.
            feedback_batch_merger.BATCH_NUMBERS = [1, 2, 3, 4, 5]

            merged_file = feedback_batch_merger.merge_feedback_batches()

            output_file_path = Path(merged_file)
            output_file_name = output_file_path.name

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
                        "Data File": data_file_name or "",
                        "Feedback File": feedback_file_name or "",
                        "Status": "Success",
                        "Created At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                ]
            )

            summary_df.to_excel(output_file_path, index=False)

        RUNS[run_id].update(
            {
                "status": "success",
                "message": "Run completed successfully.",
                "output_file": output_file_name,
                "download_url": f"/api/runs/{run_id}/download",
                "finished_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        )

    except Exception as exc:
        print("❌ RUN FAILED")
        traceback.print_exc()

        RUNS[run_id].update(
            {
                "status": "failed",
                "message": "Run failed.",
                "error": str(exc),
                "finished_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        )


@app.post("/api/runs")
async def create_run(
    background_tasks: BackgroundTasks,
    query_family: str = Form(...),
    project: str = Form(...),
    action: str = Form(...),
    month: str = Form(...),
    year: str = Form(...),
    batch: str = Form(...),
    data_file: UploadFile | None = File(None),
    feedback_file: UploadFile | None = File(None),
    query_files: list[UploadFile] | None = File(None),
    feedback_files: list[UploadFile] | None = File(None),
):
    run_id = str(uuid.uuid4())

    run_upload_dir = UPLOAD_DIR / run_id
    run_output_dir = OUTPUT_DIR / run_id

    run_upload_dir.mkdir(parents=True, exist_ok=True)
    run_output_dir.mkdir(parents=True, exist_ok=True)

    data_file_path = None
    feedback_file_path = None
    data_file_name = None
    feedback_file_name = None
    query_folder_path = None
    feedback_folder_path = None

    if action == "Merge Feedback":
        query_upload_dir = run_upload_dir / "query_files"
        feedback_upload_dir = run_upload_dir / "feedback_files"

        query_upload_dir.mkdir(parents=True, exist_ok=True)
        feedback_upload_dir.mkdir(parents=True, exist_ok=True)

        uploaded_query_files = query_files or []
        uploaded_feedback_files = feedback_files or []

        # Backward compatibility: if frontend sends only data_file/feedback_file,
        # still save them into the merge folders.
        if not uploaded_query_files and data_file:
            uploaded_query_files = [data_file]

        if not uploaded_feedback_files and feedback_file:
            uploaded_feedback_files = [feedback_file]

        if not uploaded_query_files:
            raise HTTPException(
                status_code=400,
                detail="Please upload at least one query file for Merge Feedback.",
            )

        if not uploaded_feedback_files:
            raise HTTPException(
                status_code=400,
                detail="Please upload at least one feedback file for Merge Feedback.",
            )

        for upload in uploaded_query_files:
            saved_path = save_upload_file(upload, query_upload_dir / upload.filename)
            if not data_file_path:
                data_file_path = str(saved_path)
                data_file_name = upload.filename

        for upload in uploaded_feedback_files:
            saved_path = save_upload_file(upload, feedback_upload_dir / upload.filename)
            if not feedback_file_path:
                feedback_file_path = str(saved_path)
                feedback_file_name = upload.filename

        query_folder_path = str(query_upload_dir)
        feedback_folder_path = str(feedback_upload_dir)

    else:
        if not data_file:
            raise HTTPException(
                status_code=400,
                detail="Please upload the data workbook.",
            )

        saved_data_path = save_upload_file(data_file, run_upload_dir / data_file.filename)
        data_file_path = str(saved_data_path)
        data_file_name = data_file.filename

        if feedback_file:
            saved_feedback_path = save_upload_file(feedback_file, run_upload_dir / feedback_file.filename)
            feedback_file_path = str(saved_feedback_path)
            feedback_file_name = feedback_file.filename

    RUNS[run_id] = {
        "run_id": run_id,
        "status": "queued",
        "message": "Run queued for processing.",
        "query_family": query_family,
        "project": project,
        "action": action,
        "month": month,
        "year": year,
        "batch": batch,
        "data_file": data_file_name or "",
        "feedback_file": feedback_file_name or "",
        "output_file": None,
        "download_url": None,
        "error": None,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "started_at": None,
        "finished_at": None,
    }

    background_tasks.add_task(
        process_run,
        run_id,
        query_family,
        project,
        action,
        month,
        year,
        batch,
        data_file_path,
        feedback_file_path,
        query_folder_path,
        feedback_folder_path,
        data_file_name,
        feedback_file_name,
    )

    return {
        "run_id": run_id,
        "status": "running",
        "message": "Run has started. Check status until it is complete.",
        "status_url": f"/api/runs/{run_id}/status",
        "download_url": f"/api/runs/{run_id}/download",
    }


@app.get("/api/runs/{run_id}/status")
def get_run_status(run_id: str):
    run = RUNS.get(run_id)

    if not run:
        run_output_dir = OUTPUT_DIR / run_id

        if run_output_dir.exists():
            output_files = list(run_output_dir.glob("*.zip"))
            if not output_files:
                output_files = list(run_output_dir.glob("*.xlsx"))

            if output_files:
                output_file = sorted(
                    output_files,
                    key=lambda path: path.stat().st_mtime,
                    reverse=True,
                )[0]

                return {
                    "run_id": run_id,
                    "status": "success",
                    "message": "Run completed successfully.",
                    "output_file": output_file.name,
                    "download_url": f"/api/runs/{run_id}/download",
                }

        raise HTTPException(
            status_code=404,
            detail="Run not found.",
        )

    return run


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
            detail="No output file found for this run yet.",
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