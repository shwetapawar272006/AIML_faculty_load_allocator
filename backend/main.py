from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import io, pandas as pd
from processor import process_subject_allotment

app = FastAPI(title="AIML Faculty Load Allocator")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def _excel_to_preview(output_bytes: bytes) -> dict:
    """Parse the output Excel into JSON for the frontend preview."""
    xls = pd.ExcelFile(io.BytesIO(output_bytes))

    # --- subject allocation ---
    sa = pd.read_excel(xls, sheet_name="subject allocation", header=None)
    faculty = []
    for i, row in sa.iterrows():
        if i < 8:
            continue
        idx = row.iloc[0]
        if idx is None or (isinstance(idx, float) and pd.isna(idx)):
            continue
        def _v(col):
            v = row.iloc[col] if len(row) > col else None
            return None if (v is None or (isinstance(v, float) and pd.isna(v))) else float(v) if isinstance(v, (int, float)) else v
        faculty.append({
            "idx": int(idx), "name": str(row.iloc[1]).strip(),
            "designation": str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
            "theory": _v(3), "lab": _v(4), "tu": _v(5) or 0,
            "r": _v(6), "c": _v(7), "p": _v(8),
            "i": _v(9), "e": _v(10), "admin": _v(11),
            "total": _v(13),
        })

    # --- core1-4 break-up ---
    c4 = pd.read_excel(xls, sheet_name="Core1-4 Break-Up", header=None)
    cores = []
    for i, row in c4.iterrows():
        if i == 0:
            continue
        idx = row.iloc[0]
        if idx is None or (isinstance(idx, float) and pd.isna(idx)):
            continue
        core_slots = []
        for col_t, col_d in [(3,4),(5,6),(7,8),(9,10)]:
            t = str(row.iloc[col_t]).strip() if (len(row) > col_t and pd.notna(row.iloc[col_t])) else None
            d = str(row.iloc[col_d]).strip() if (len(row) > col_d and pd.notna(row.iloc[col_d])) else None
            core_slots.append({"type": t, "detail": d} if t else None)
        cores.append({
            "idx": int(idx), "name": str(row.iloc[1]).strip(),
            "designation": str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
            "cores": core_slots,
        })

    return {"subject_allocation": faculty, "core_breakup": cores}


@app.post("/generate")
async def generate(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "Only Excel files allowed")
    content = await file.read()
    try:
        output_bytes = process_subject_allotment(content)
        return JSONResponse(_excel_to_preview(output_bytes))
    except Exception as e:
        raise HTTPException(500, f"Processing error: {e}")


@app.post("/download")
async def download(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "Only Excel files allowed")
    content = await file.read()
    try:
        output_bytes = process_subject_allotment(content)
        return StreamingResponse(
            io.BytesIO(output_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=CMRIT_TRCPIE_Faculty_Core_Load_AIML_2025-26_Even_Sem.xlsx"}
        )
    except Exception as e:
        raise HTTPException(500, f"Processing error: {e}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
