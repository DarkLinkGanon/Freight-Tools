Vercel-ready Freight Extract Tools

Files included:
- app.py
- templates/index.html
- requirements.txt
- vercel.json

What this version adds:
- Drag and drop uploads
- File size shown next to each selected file
- Files larger than 4 MB highlighted in red
- File size validation before upload
- Upload progress bar instead of spinner
- Single file export keeps original file name + " extracted data"
- Multiple files export uses "combined extracted data.xlsx"

How to run locally:
1. Open a terminal in this folder
2. Run:

py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
py app.py

3. Open:
http://127.0.0.1:5000

How to deploy to Vercel:
1. Put this project in a GitHub repository
2. Import the repo into Vercel
3. Vercel should detect the Python build from vercel.json
4. Deploy

Notes:
- This version blocks files over 4 MB in the browser before upload.
- Vercel request limits still apply, so keeping uploads below 4 MB is a good safety margin.
