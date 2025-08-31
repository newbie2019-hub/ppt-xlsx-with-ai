## PPT → XLSX Converter

This project is a **Node.js application** that parses `.pptx` files, extracts structured data, and generates an `.xlsx` spreadsheet.
It was built to analyze **slides and notes** from PowerPoint presentations and transform them into a **clean, AI-processed JSON** and finally an **Excel file** for easy review and analysis.

---

## Features
- 🔍 **Manual PPTX Parsing** – `.pptx` files are internally converted to `.zip` and analyzed by reading raw XML for slides and notes.
- 🤖 **AI Processing** – Raw extracted data is cleaned, structured, and finalized using AI.
- 📑 **JSON Output** – A `final_slides.json` file is generated containing structured content.
- 📊 **Excel Export** – Converts the processed JSON into an `.xlsx` file.
- 🌐 **Simple Web UI** – A minimal HTML page allows users to upload a `.pptx` file and download the processed `.xlsx` result.

---

## Installation
```bash
# 1. Clone the repository:
git clone https://github.com/your-username/ppt-xlsx-app.git
cd ppt-xlsx-app
```

```bash
# Install dependencies
npm install

# Run the app:
npm start
```

### 📖 Usage
1. Open `localhost:3000` in your browser.
2. Upload a .pptx file.
3. Wait for processing – the system will:
    - Extract slides & notes.
    - Process content into structured JSON.
    - Generate and return an .xlsx file.
4. Download the Excel sheet and review your data.

### 📝 Example Workflow
- Input: lecture-anatomy.pptx
  - Extracted:
  - output/lecture-anatomy/final_slides.json
  - output/lecture-anatomy/final_slides.xlsx

### 🛠 Tech Stack
- Node.js – Backend processing
- fs/admZip – File handling & extraction
- XML Parsing – Reading PPTX XML content
- OpenAI / Gemini – AI cleanup and structuring
- exceljs – Excel file generation
- HTML + JavaScript – Upload interface

### 🤝 Contributing
Pull requests and improvements are welcome! If you find issues, please open an issue.