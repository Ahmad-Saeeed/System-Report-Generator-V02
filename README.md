# System Report Generator

This is a Claude-skill converted to a Standalone Windows desktop application for generating the System-based testing progress report and Outlook-ready email from a test-results DOCX to be shared between my colleagues.

## Runtime requirements

The final `TMS_Report_Generator.exe` does **not** require Python to be installed.
It does not use an LLM, API key, or tokens, and its report generation logic is local.

## Build the Windows EXE without installing Python locally

This repository includes a GitHub Actions workflow:

`.github/workflows/build-windows-exe.yml`

1. Create a GitHub repository.
2. Upload this project to it.
3. Push to `main` (or `master`), or open **Actions → Build Windows EXE → Run workflow**.
4. When the workflow finishes, open the workflow run and download the artifact named `TMS_Report_Generator-Windows`.
5. Extract it and run `TMS_Report_Generator.exe` on Windows 10/11.

## Build locally

If Python is installed on a Windows build machine, run `build_windows_exe.bat`.

## Use

1. Open the EXE.
2. Select the test-report `.docx`.
3. Choose System, Environment, and Report Type.
4. Add observations if needed.
5. Click **Generate Report + Email**.
6. The HTML report and email are saved beside the selected DOCX.


## Windows packaging / antivirus false positives

The Windows build uses PyInstaller OneDir mode, disables UPX compression, embeds
Windows version metadata, and pins PyInstaller to 6.22.3 for reproducible builds.
Distribute the complete `dist/TMS_Report_Generator/` folder rather than extracting
only the EXE from it.
