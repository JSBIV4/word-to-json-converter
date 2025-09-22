# Word to JSON Converter - Complete Setup Guide

## Quick Setup with Virtual Environment

### 1. Create Project Directory and Transfer the files shared. 
```bash
mkdir word-to-json-converter
cd word-to-json-converter
```

### 2. Create Virtual Environment

**For Windows:**
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate

# You should see (venv) in your command prompt
```

**For Linux/MacOS:**
```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# You should see (venv) in your terminal
```

### 3. Install Required Packages

**Option A: Using requirements.txt (Recommended)**
```bash
pip install -r requirements.txt
```

**Option B: Manual Installation**
```bash
pip install streamlit python-docx
```

### 4. Create Your Files
Save the converter class as `convert_docx_to_json.py` and the Streamlit UI as `streamlit_app.py` in your project folder.

### 5. Run the Application
```bash
streamlit run streamlit_app.py
```

The app will open in your browser at `http://localhost:8501`

## Project Structure
```
word-to-json-converter/
│
├── venv/                     # Virtual environment (created automatically)
├── convert_docx_to_json.py   # Pure converter class
├── streamlit_app.py          # Streamlit web interface
├── requirements.txt          # Package dependencies
├── README.md                 # This file
└── sample_data/              # Your Word documents (optional)
    └── test_documents/
```

## Requirements.txt
Create a `requirements.txt` file with:
```
streamlit>=1.28.0
python-docx>=0.8.11
```

## Environment Management

### Deactivating Virtual Environment
When you're done working:
```bash
deactivate
```

### Reactivating Virtual Environment
Next time you work on the project:

**Windows:**
```bash
cd word-to-json-converter
venv\Scripts\activate
streamlit run streamlit_app.py
```

**Linux/MacOS:**
```bash
cd word-to-json-converter
source venv/bin/activate
streamlit run streamlit_app.py
```

### Updating Dependencies
To update packages:
```bash
pip install --upgrade streamlit python-docx
pip freeze > requirements.txt  # Update requirements file
```

## Making it Accessible to Others

### Option 1: Local Network Access
```bash
streamlit run streamlit_app.py --server.address 0.0.0.0 --server.port 8501
```
Others on your network can access it at `http://YOUR_IP_ADDRESS:8501`

**Find your IP address:**
- **Windows:** `ipconfig`
- **Linux/MacOS:** `ifconfig` or `ip addr show`

### Option 2: Deploy to Streamlit Cloud (Free)
1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repository
4. Deploy your app

### Option 3: Run with Custom Port
```bash
streamlit run streamlit_app.py --server.port 8080
```

## Usage Examples

### Command Line Usage
```python
# In convert_docx_to_json.py or separate script
from convert_docx_to_json import WordToJsonConverter

converter = WordToJsonConverter()

# Convert single file
converter.convert_single_file("document.docx", "output.json")

# Convert entire folder
converter.convert_folder("input_folder/", "output_folder/")
```

### Input Word Document Format:
```
Patient Name: John Smith
Age: 45
Diagnosis: Type 2 Diabetes
Treatment: Metformin 500mg
Notes: Patient responds well to treatment
```

### Output JSON:
```json
{
  "metadata": {
    "source_file": "patient_record.docx",
    "converted_at": "2024-01-20T10:30:00",
    "total_paragraphs": 5,
    "file_type": "docx"
  },
  "content": {
    "Patient Name": "John Smith",
    "Age": "45", 
    "Diagnosis": "Type 2 Diabetes",
    "Treatment": "Metformin 500mg",
    "Notes": "Patient responds well to treatment"
  }
}
```

## Troubleshooting

### Common Issues:

**1. "python is not recognized" (Windows)**
```bash
# Install Python from python.org or use:
py -m venv venv
py -m pip install -r requirements.txt
```

**2. "No module named 'docx'"**
```bash
# Make sure virtual environment is activated, then:
pip install python-docx
```

**3. "Permission denied" (Linux/MacOS)**
```bash
# Use python3 and pip3:
python3 -m venv venv
pip3 install -r requirements.txt
```

**4. "Port already in use"**
```bash
# Try a different port:
streamlit run streamlit_app.py --server.port 8080
```

**5. "ModuleNotFoundError: No module named 'convert_docx_to_json'"**
```bash
# Make sure both files are in the same directory
# Make sure virtual environment is activated
```

### Virtual Environment Issues:

**Recreate virtual environment:**
```bash
# Remove old environment
rm -rf venv  # Linux/MacOS
rmdir /s venv  # Windows

# Create new one
python -m venv venv
# Activate and install requirements again
```

## Security & Performance Notes

### Security:
- Files are processed in memory and not stored permanently
- No data is sent to external servers (runs locally)
- Safe to use with sensitive documents
- Virtual environment isolates dependencies

### Performance Tips:
- Process files in smaller batches for large documents
- The app automatically handles multiple file uploads
- Use SSD storage for faster file processing
- Close other applications for better performance with large files

## Additional Resources

- [Streamlit Documentation](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/)
- [python-docx Documentation](https://python-docx.readthedocs.io/)
- [Python Virtual Environments Guide](https://docs.python.org/3/tutorial/venv.html)

## Getting Help

If you encounter issues:
1. Check that your virtual environment is activated
2. Ensure all files are in the correct directory
3. Verify your Word documents are proper .docx format (not text files with .docx extension)
4. Check the terminal for detailed error messages