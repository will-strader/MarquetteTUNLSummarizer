# Rat Performance Analyzer

This web-based tool allows researchers to upload TUNL test data from CSV or Excel files and receive performance summaries broken down by custom-defined distance ranges. It's built using Streamlit for ease of use, even for those with no programming experience.

## Features

- Upload one or more `.csv` or `.xlsx` files
- Customizable column inputs (e.g., animal ID, trial number, distance, correctness)
- Define distance bins like `1-4`, `5-8`, `9-13`
- View a summary table in-browser
- Download processed results as an Excel file

## Running Locally

1. Clone the repository:
   git clone https://github.com/jxiulu/TUNLTestProgram.git
   cd TUNLTestProgram

2. Install Requirements
   pip install -r requirements.txt

3. Launch the Streamlit app
   streamlit run webapp.py
   
This will open the useable app in your browser at http://localhost:8501.

## Deploying for Others (No Coding Needed)

1. Push this repo to GitHub
2. Go to Streamlit Cloud and sign in with GitHub
3. Click “New App” and select your repository and webapp.py as the app file
4. Click “Deploy” — Streamlit will host it and give you a shareable link

 *Make sure your requirements include: pandas, streamlit, openpxyl*
 (Optional requirement if you need to support GUI: wxPython)

## License
This project is licensed under the MIT License.

## Authors
Jerry Lu and Will Strader 2025
