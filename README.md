# RIS Full Workflow - Usage Guide

## Overview
`ris_full_workflow.py` - Complete end-to-end automation for RIS radiology reporting

## What It Does
1. **Date Selection**: Prompts you to enter a date (or press Enter for today)
2. **Search**: Sets filters (Category=一般攝影, Location=台大總院, Exam=32001CXM) and performs search
3. **Open**: Selects the first result in the grid and opens it
4. **Extract**: Waits for PACS viewer, sends Ctrl+I to anonymize, then takes a screenshot of the viewer window
5. **Analyze**: Sends image to MedGemma 1.5 (via HF Inference Endpoint) for structured radiology report generation
6. **Report**: Enters AI-generated report into RIS EXAM text field

## How to Run

### Prerequisites
- Python installed on your system
- RIS application must be open (main window visible)
- Close any existing PACS viewer windows
- Internet connection for HF Inference Endpoint
- Hugging Face token and deployed MedGemma 1.5 endpoint URL

### Method 1: Double-click `run.bat`

1. Deploy MedGemma 1.5 at https://endpoints.huggingface.co/
2. Open `run.bat` in Notepad
3. Replace `your-hf-token-here` with your Hugging Face token
4. Replace `https://your-endpoint.endpoints.huggingface.cloud` with your endpoint URL
5. Save and close
6. Double-click `run.bat` to run

### Method 2: Command Line (PowerShell)

```powershell
$env:HF_TOKEN = "your-hf-token-here"
$env:HF_ENDPOINT_URL = "https://your-endpoint.endpoints.huggingface.cloud"
python "ris_full_workflow.py"
```

## Configuration

- **HF Token**: Set via environment variable `HF_TOKEN`
- **Endpoint URL**: Set via environment variable `HF_ENDPOINT_URL`
- **Output Files** (saved in the same folder as the script):
  - `extracted_xray.png` - Captured X-ray screenshot
  - `report.txt` - AI-generated report
  - `workflow_log.txt` - Execution log (no patient data)

## Date Input Format

When prompted, you can enter:
- Press **Enter** for today's date
- `YYYY/MM/DD` (e.g., `2026/01/15`)
- `MM/DD` (e.g., `01/15` - uses current year)
- Also accepts `-` separator (e.g., `2026-01-15`)

## Troubleshooting

- **"HF_TOKEN not set"**: Edit `run.bat` and add your HF token and endpoint URL
- **Date not setting**: Check "Final Date Values" output before search
- **No results in grid**: Verify filters and date are correct
- **PACS not detected**: Window title must start with `[總院]`
- **Screenshot failed**: PACS viewer window may be minimized or off-screen
- **Report not entered**: Ensure EXAM text box is visible

## Files

```
RRG_medgemma/
├── ris_full_workflow.py   - Main automation script
├── RIS_WORKFLOW_GUIDE.md  - This documentation
└── run.bat                - Click to run (edit HF token first)
```
