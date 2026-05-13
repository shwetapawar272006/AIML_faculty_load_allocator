# AIML Faculty Load Allocator

## Project Structure
```
faculty_load_app/
├── backend/
│   ├── main.py           # FastAPI server
│   ├── processor.py      # Core logic (reads input Excel, produces output)
│   └── requirements.txt
└── frontend/
    ├── src/
    │   ├── components/
    │   │   └── FileUpload.tsx
    │   ├── App.tsx
    │   ├── main.tsx
    │   └── index.css
    ├── index.html
    ├── package.json
    ├── tsconfig.json
    └── vite.config.ts
```

## How to Run

### Step 1 — Backend (Python / FastAPI)

Open a terminal and run:

```bash
cd backend
pip install -r requirements.txt
python main.py
```

The API will start at **http://localhost:8000**

### Step 2 — Frontend (React / Vite)

Open a second terminal and run:

```bash
cd frontend
npm install
npm run dev
```

The UI will open at **http://localhost:5173**

### Step 3 — Use the App

1. Open http://localhost:5173 in your browser
2. Upload the **Subject Allotment Even Semester** Excel file
3. Click **Generate & Download Report**
4. The output Excel file will be downloaded automatically

## Output File

The generated Excel contains **3 sheets**:

| Sheet | Description |
|-------|-------------|
| **TRCPIE** | Full TRCPIE Allocation sheet (Even Semester version) |
| **subject allocation** | Faculty-wise Theory, Lab, TRCPIE unit breakdown |
| **Core1-4 Break-Up** | Detailed core slot assignment per faculty |

## How it Works

All values are computed purely from the input file:

- **Theory (T) and Lab (L) units** → read directly from the "TRCPIE Allocation" sheet
- **R, C, P, I, E, Admin values** → read from the "TRCPIE Allocation" sheet  
- **Core 1-4 details** → derived from "Subject List" sheet (section, course code, subject name)
- **Name matching** → fuzzy algorithm handles abbreviations, prefixes, and reversed name tokens
  (e.g. "Swasthisudha Punyathuya" ↔ "Swasti Sudha", "Prabhakara S" ↔ "Prabhakar")
