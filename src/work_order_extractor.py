"""
Work Order Excel to Vector Database - MVP Scaffold
Quick implementation to demo concept, easily expandable based on requirements
"""

import pandas as pd
import openpyxl
from typing import List, Dict, Any, Optional
import json

class WorkOrderExtractor:
    """Basic extractor that can be easily expanded"""
    
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.workbook = openpyxl.load_workbook(excel_path, data_only=True)
        
    def extract_all_sheets(self) -> List[Dict[str, Any]]:
        """Extract data from all SKU sheets"""
        all_chunks = []
        
        # Process all SKU sheets
        sku_sheets = [name for name in self.workbook.sheetnames 
                     if 'SKU' in name and name != 'Copy_Paste Rows']
        
        for sheet_name in sku_sheets:
            try:
                chunks = self.extract_sheet(sheet_name)
                all_chunks.extend(chunks)
            except Exception as e:
                print(f"Warning: Could not process sheet {sheet_name}: {e}")
                
        return all_chunks
    
    def extract_sheet(self, sheet_name: str) -> List[Dict[str, Any]]:
        """Extract chunks from a single sheet"""
        sheet = self.workbook[sheet_name]
        chunks = []
        
        # Extract project metadata (basic version)
        project_info = self._extract_project_info(sheet)
        
        # Determine number of SKUs from sheet name
        sku_count = self._get_sku_count(sheet_name)
        
        # Extract SKU data blocks
        sku_blocks = self._extract_sku_blocks(sheet, sku_count)
        
        # Create chunks (start simple, can expand)
        for sku_data in sku_blocks:
            if sku_data:  # Only process non-empty SKUs
                chunks.append(self._create_sku_chunk(project_info, sku_data, sheet_name))
        
        # Add project overview chunk
        if sku_blocks:
            chunks.append(self._create_project_chunk(project_info, sku_blocks, sheet_name))
            
        return chunks
    
    def _extract_project_info(self, sheet) -> Dict[str, str]:
        """Extract basic project information"""
        try:
            return {
                'project_id': self._get_cell_value(sheet, 'B2', ''),
                'project_manager': self._get_cell_value(sheet, 'B3', ''),
                'salesperson': self._get_cell_value(sheet, 'P3', ''),
                'notes': self._get_cell_value(sheet, 'B5', '')
            }
        except:
            return {'project_id': '', 'project_manager': '', 'salesperson': '', 'notes': ''}
    
    def _get_sku_count(self, sheet_name: str) -> int:
        """Extract SKU count from sheet name"""
        try:
            if sheet_name == '1 SKU':
                return 1
            else:
                return int(sheet_name.split()[0])
        except:
            return 1
    
    def _extract_sku_blocks(self, sheet, sku_count: int) -> List[Dict[str, Any]]:
        """Extract SKU specification blocks"""
        sku_blocks = []
        
        # SKU blocks start at row 8 and repeat every ~9 rows
        base_row = 8
        
        for sku_num in range(1, sku_count + 1):
            try:
                # Calculate row offset for this SKU
                if sku_num == 1:
                    sku_row = base_row
                else:
                    sku_row = base_row + ((sku_num - 1) * 9)  # Approximate spacing
                
                sku_data = {
                    'sku_number': sku_num,
                    'description': self._get_cell_value(sheet, f'D{sku_row+1}', ''),
                    'poa_type': self._get_cell_value(sheet, f'A{sku_row+2}', ''),
                    'file_location': self._get_cell_value(sheet, f'D{sku_row+2}', ''),
                    'color_target': self._get_cell_value(sheet, f'G{sku_row+2}', ''),
                    'proof_requirements': self._get_cell_value(sheet, f'J{sku_row+2}', ''),
                    'substrate': self._get_cell_value(sheet, f'M{sku_row+2}', ''),
                    'varnishes': self._get_cell_value(sheet, f'P{sku_row+2}', ''),
                    'special_fx': self._get_cell_value(sheet, f'U{sku_row+2}', ''),
                }
                
                # Only add if we have some actual data
                if any(sku_data.values()):
                    sku_blocks.append(sku_data)
                    
            except Exception as e:
                print(f"Warning: Could not extract SKU {sku_num}: {e}")
                continue
                
        return sku_blocks
    
    def _get_cell_value(self, sheet, cell_ref: str, default: str = '') -> str:
        """Safely get cell value"""
        try:
            cell = sheet[cell_ref]
            return str(cell.value) if cell.value is not None else default
        except:
            return default
    
    def _create_sku_chunk(self, project_info: Dict, sku_data: Dict, sheet_name: str) -> Dict[str, Any]:
        """Create a searchable chunk for a single SKU - BASIC VERSION"""
        
        # Basic content format (can expand based on Mark's feedback)
        content = f"""
        Project {project_info['project_id']} - SKU {sku_data['sku_number']}
        
        Description: {sku_data['description']}
        Production Type: {sku_data['poa_type']}
        Material/Substrate: {sku_data['substrate']}
        Color Requirements: {sku_data['color_target']}
        Proof Requirements: {sku_data['proof_requirements']}
        Varnishes: {sku_data['varnishes']}
        Special Effects: {sku_data['special_fx']}
        File Location: {sku_data['file_location']}
        
        Project Manager: {project_info['project_manager']}
        Salesperson: {project_info['salesperson']}
        """.strip()
        
        return {
            'chunk_id': f"{project_info['project_id']}_SKU_{sku_data['sku_number']}",
            'chunk_type': 'sku_specification',
            'content': content,
            'metadata': {
                'project_id': project_info['project_id'],
                'sku_number': sku_data['sku_number'],
                'sheet_name': sheet_name,
                'substrate': sku_data['substrate'],
                'poa_type': sku_data['poa_type'],
                'color_target': sku_data['color_target'],
                'project_manager': project_info['project_manager'],
                'salesperson': project_info['salesperson']
            }
        }
    
    def _create_project_chunk(self, project_info: Dict, sku_blocks: List[Dict], sheet_name: str) -> Dict[str, Any]:
        """Create a project overview chunk"""
        
        sku_summaries = []
        for sku in sku_blocks:
            sku_summaries.append(f"SKU {sku['sku_number']}: {sku['description']} ({sku['substrate']})")
        
        content = f"""
        Project {project_info['project_id']} Overview
        
        Project Manager: {project_info['project_manager']}
        Salesperson: {project_info['salesperson']}
        Total SKUs: {len(sku_blocks)}
        
        SKU Summary:
        {chr(10).join(sku_summaries)}
        
        Notes: {project_info['notes']}
        """.strip()
        
        return {
            'chunk_id': f"{project_info['project_id']}_overview",
            'chunk_type': 'project_overview', 
            'content': content,
            'metadata': {
                'project_id': project_info['project_id'],
                'sku_count': len(sku_blocks),
                'sheet_name': sheet_name,
                'project_manager': project_info['project_manager'],
                'salesperson': project_info['salesperson']
            }
        }

# EXPANSION HOOKS - Easy to modify based on Mark's feedback

class WorkOrderEnhancer:
    """Add-on class for enhanced features based on requirements"""
    
    @staticmethod
    def add_fuzzy_search_terms(chunk: Dict[str, Any]) -> Dict[str, Any]:
        """Add fuzzy search capabilities if needed"""
        # TODO: Add based on Mark's feedback about search patterns
        enhanced_content = chunk['content']
        
        # Example expansions (activate if needed):
        # enhanced_content += WorkOrderEnhancer._add_synonym_expansions(chunk)
        # enhanced_content += WorkOrderEnhancer._add_visual_descriptors(chunk)
        
        chunk['enhanced_content'] = enhanced_content
        return chunk
    
    @staticmethod
    def _add_synonym_expansions(chunk: Dict[str, Any]) -> str:
        """Add technical term synonyms for fuzzy matching"""
        synonyms = {
            'Flexo': 'flexographic printing, flexible',
            'Digital': 'digital printing, high quality',
            'PET': 'clear plastic, transparent film',
            'PP': 'flexible plastic, soft packaging',
            # Add more based on actual data patterns
        }
        
        expansions = []
        for term, description in synonyms.items():
            if term.lower() in chunk['content'].lower():
                expansions.append(f"{term}: {description}")
        
        if expansions:
            return f"\n\nAlso known as: {', '.join(expansions)}"
        return ""

# Usage Example
def demo_extraction(excel_path: str):
    """Demo function to show Mark the concept"""
    extractor = WorkOrderExtractor(excel_path)
    chunks = extractor.extract_all_sheets()
    
    print(f"Extracted {len(chunks)} chunks from work order")
    print(f"Sample chunk:")
    print("="*50)
    print(chunks[0]['content'])
    print("="*50)
    print(f"Metadata: {json.dumps(chunks[0]['metadata'], indent=2)}")
    
    return chunks

# Quick test function
if __name__ == "__main__":
    # Test with your file
    chunks = demo_extraction("WO PG 26794.xlsx")
    
    # Save output for demo
    with open("extracted_chunks.json", "w") as f:
        json.dump(chunks, f, indent=2)
    
    print(f"\nSaved {len(chunks)} chunks to extracted_chunks.json")
    print("Ready to demo to Mark!")