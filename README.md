# Excel to RAG JSON Converter

Transform Excel work order sheets into searchable JSON chunks for vector databases and RAG pipelines.

## 🎯 Overview

This project converts complex Excel work order templates into clean, searchable JSON format optimized for vector databases like AstraDB. Perfect for internal chatbots and document search systems.

## 📁 Repository Structure

```
excel-rag-converter/
├── README.md                    # This file
├── demo/
│   └── excel-rag-demo.html     # Interactive web demo
├── src/
│   └── work_order_extractor.py # Production Python code
├── examples/
│   ├── sample_input.xlsx       # Example work order file
│   └── sample_output.json      # Example extracted JSON
├── docs/
│   ├── analysis.md             # Technical analysis
│   └── astradb_integration.md  # Vector DB setup guide
└── requirements.txt            # Python dependencies

```

## 🚀 Quick Start

### Demo (No Installation Required)
1. Open `demo/excel-rag-demo.html` in any browser
2. Drag & drop your Excel file
3. See instant JSON output!

### Production Use
```bash
pip install pandas openpyxl
python src/work_order_extractor.py your_file.xlsx
```

## 💡 Use Cases

- **Internal Chatbots**: "Find projects with metallic substrates"
- **Project Search**: "Show me all Digital Flexo work orders"
- **Specification Lookup**: "What color targets did we use for client X?"
- **Cross-Reference**: "Similar projects to this one?"

## 📊 What Gets Extracted

### Project Metadata
- Project ID, Manager, Salesperson
- Notes and special requirements

### SKU Specifications  
- Production type (POA)
- Materials and substrates
- Color requirements
- Proof requirements
- Special effects and finishes

### Output Format (AstraDB Ready)
```json
{
  "document_id": "PG26794_SKU_1",
  "content": "Project PG26794 - SKU 1\n\nDescription: Premium Coffee Pouch...",
  "metadata": {
    "project_id": "PG26794",
    "substrate": "Metallized PET/PE",
    "has_metallic": true,
    "complexity_level": "high"
  }
}
```

## 🎨 Features

- **Zero Backend Required**: Demo runs entirely in browser
- **Privacy First**: No data uploaded anywhere
- **AstraDB Optimized**: Perfect JSON format for vector search
- **Fuzzy Search Ready**: Includes synonyms and descriptive text
- **Mobile Friendly**: Works on any device

## 🔧 Integration

### AstraDB Vector Database
```python
from astrapy.db import AstraDB

# Initialize AstraDB
db = AstraDB(token="your-token", api_endpoint="your-endpoint")
collection = db.collection("work_orders")

# Insert extracted chunks
for chunk in extracted_chunks:
    collection.insert_one(chunk)
```

### Custom RAG Pipeline
```python
from src.work_order_extractor import WorkOrderExtractor

# Extract from Excel
extractor = WorkOrderExtractor("work_order.xlsx")
chunks = extractor.extract_all_sheets()

# Send to your vector DB
your_vector_db.bulk_insert(chunks)
```

## 📈 Statistics

Example extraction from a 30-SKU work order:
- **62 total chunks** generated
- **30 SKU specifications** + **2 project overviews**
- **~400 words per chunk** (optimal for vector search)
- **Rich metadata** for precise filtering

## 🎯 Next Steps

1. **Demo to stakeholders** using the web interface
2. **Integrate with your vector database**
3. **Customize extraction rules** for your specific needs
4. **Add fuzzy search enhancements** based on user feedback

## 🤝 Contributing

This is an internal tool, but feel free to:
- Add new extraction patterns
- Improve the demo interface  
- Enhance metadata extraction
- Add export formats

## 📝 License

Internal use only - [Your Company Name]

---

**Built for materials printing companies who can't remember if they've done things before** 😄
