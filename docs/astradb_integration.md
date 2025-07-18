# AstraDB Integration Guide

Complete setup guide for integrating Excel work order extraction with AstraDB vector database.

## 🚀 Quick Setup

### 1. Install Dependencies
```bash
pip install astrapy pandas openpyxl
```

### 2. Set Up AstraDB Collection
```python
from astrapy.db import AstraDB

# Initialize connection
db = AstraDB(
    token="your-application-token",
    api_endpoint="your-database-endpoint"
)

# Create collection for work orders
collection = db.create_collection(
    "work_orders",
    dimension=1536,  # For OpenAI embeddings
    metric="cosine"
)
```

### 3. Extract and Upload
```python
from src.work_order_extractor import WorkOrderExtractor

# Extract chunks from Excel
extractor = WorkOrderExtractor("work_order.xlsx")
chunks = extractor.extract_all_sheets()

# Insert into AstraDB
for chunk in chunks:
    collection.insert_one({
        "document_id": chunk["document_id"],
        "content": chunk["content"],
        "metadata": chunk["metadata"]
    })
```

## 🔍 Search Examples

### Basic Text Search
```python
# Find projects with metallic substrates
results = collection.vector_find(
    "metallic substrate projects",
    limit=5
)

for result in results:
    print(f"Project: {result['metadata']['project_id']}")
    print(f"Content: {result['content'][:200]}...")
```

### Metadata Filtering
```python
# Find all projects by specific manager
results = collection.find(
    filter={"metadata.project_manager": "John Smith"},
    limit=10
)
```

### Complex Queries
```python
# Find flexible packaging with special effects
results = collection.vector_find(
    "flexible packaging holographic foil effects",
    filter={
        "metadata.material_type": "flexible_packaging",
        "metadata.has_special_effects": True
    },
    limit=5
)
```

## 📊 Optimal Chunking Strategy

### Chunk Types We Generate

1. **SKU Specifications** (`sku_specification`)
   - Individual product details
   - ~400 words per chunk
   - Rich technical metadata

2. **Project Overviews** (`project_overview`)
   - High-level project summaries
   - Cross-SKU relationships
   - Business context

### Metadata Schema
```python
sku_metadata = {
    "project_id": str,           # "PG26794"
    "sku_number": int,           # 1, 2, 3...
    "substrate": str,            # "Metallized PET/PE"
    "poa_type": str,            # "Digital Flexo"
    "color_target": str,         # "4C + Metallic Silver"
    "project_manager": str,      # "John Smith"
    "salesperson": str,          # "Sarah Johnson"
    "has_metallic": bool,        # True/False
    "has_special_effects": bool, # True/False
    "material_type": str,        # "flexible_packaging"
    "complexity_level": str      # "low", "medium", "high"
}
```

## 🎯 Search Optimization

### Employee Query Patterns

**Exact Searches:**
- "Project PG26794 specifications"
- "Digital Flexo substrate options"
- "John Smith managed projects"

**Fuzzy Searches:**
- "that shiny coffee pouch thing"
- "metallic flexible packaging"
- "holographic effect projects"

### Embedding Strategy
```python
import openai

def create_enhanced_content(chunk):
    """Add searchable synonyms for better retrieval"""
    
    content = chunk["content"]
    
    # Add technical synonyms
    if "Digital Flexo" in content:
        content += "\nAlso known as: digital flexographic printing, high-quality flexible printing"
    
    if "Metallized" in content:
        content += "\nVisual: shiny, reflective, metallic finish, premium look"
    
    if "UV" in content:
        content += "\nFinish: glossy coating, protective layer, enhanced durability"
    
    return content

# Enhanced insertion
for chunk in chunks:
    enhanced_content = create_enhanced_content(chunk)
    
    collection.insert_one({
        "document_id": chunk["document_id"],
        "content": enhanced_content,
        "metadata": chunk["metadata"]
    })
```

## 🔧 Production Pipeline

### Batch Processing
```python
def process_work_order_batch(excel_files):
    """Process multiple Excel files efficiently"""
    
    all_chunks = []
    
    for file_path in excel_files:
        try:
            extractor = WorkOrderExtractor(file_path)
            chunks = extractor.extract_all_sheets()
            all_chunks.extend(chunks)
            print(f"✅ Processed {file_path}: {len(chunks)} chunks")
            
        except Exception as e:
            print(f"❌ Failed {file_path}: {e}")
    
    # Bulk insert for efficiency
    if all_chunks:
        collection.insert_many(all_chunks)
        print(f"🚀 Uploaded {len(all_chunks)} total chunks")
    
    return all_chunks
```

### Error Handling
```python
def safe_insert_chunk(collection, chunk, max_retries=3):
    """Insert chunk with retry logic"""
    
    for attempt in range(max_retries):
        try:
            collection.insert_one(chunk)
            return True
            
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"❌ Failed to insert {chunk['document_id']}: {e}")
                return False
            else:
                print(f"⚠️ Retry {attempt + 1} for {chunk['document_id']}")
                time.sleep(1)
    
    return False
```

## 📈 Performance Tips

### Chunking Best Practices
- **~400 words per chunk** for optimal retrieval
- **Rich metadata** for precise filtering
- **Natural language content** over raw data
- **Cross-references** between related chunks

### Search Optimization
- Use **hybrid search** (semantic + keyword)
- **Filter by metadata** before vector search
- **Limit results** to manageable sets (5-10)
- **Include context** in chunk content

### Scaling Considerations
- **Batch uploads** for large datasets
- **Incremental updates** for new work orders  
- **Deduplication** for template variations
- **Monitoring** for search quality

## 🎨 Custom Integration

### Your RAG Pipeline
```python
class WorkOrderRAG:
    def __init__(self, astra_collection):
        self.collection = astra_collection
    
    def search(self, query, filters=None):
        """Search work orders with business logic"""
        
        # Add common synonyms
        enhanced_query = self.enhance_query(query)
        
        # Search with filters
        results = self.collection.vector_find(
            enhanced_query,
            filter=filters,
            limit=5
        )
        
        return self.format_results(results)
    
    def enhance_query(self, query):
        """Add search synonyms for better matching"""
        synonyms = {
            "shiny": "metallic reflective",
            "flexible": "flexible packaging PET PE",
            "colorful": "4C process color multicolor"
        }
        
        enhanced = query
        for term, expansion in synonyms.items():
            if term in query.lower():
                enhanced += f" {expansion}"
        
        return enhanced
    
    def format_results(self, results):
        """Format results for chatbot responses"""
        formatted = []
        
        for result in results:
            formatted.append({
                "project": result["metadata"]["project_id"],
                "sku": result["metadata"]["sku_number"],
                "summary": result["content"][:200] + "...",
                "relevance": result["similarity"],
                "metadata": result["metadata"]
            })
        
        return formatted
```

## 🚀 Go Live Checklist

- [ ] AstraDB collection created with correct dimensions
- [ ] Test extraction on sample Excel files
- [ ] Verify chunk quality and metadata completeness
- [ ] Test search queries with expected employee language
- [ ] Set up batch processing for historical files
- [ ] Configure monitoring and error handling
- [ ] Train users on search patterns and capabilities

## 📞 Support

For AstraDB specific issues:
- [AstraDB Documentation](https://docs.datastax.com/en/astra/docs/)
- [Python Client Reference](https://github.com/datastax/astrapy)

For extraction issues:
- Check Excel file format compatibility
- Verify sheet naming conventions
- Review extraction logs for errors