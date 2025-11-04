import json
import os
from datetime import datetime
from typing import Dict, List, Any
import hashlib

class MappingLearningSystem:
    def __init__(self, learning_file: str = "mapping_history.json"):
        self.learning_file = learning_file
        self.history = self._load_history()
    
    def _load_history(self) -> Dict[str, Any]:
        """Load existing mapping history from file"""
        if os.path.exists(self.learning_file):
            try:
                with open(self.learning_file, 'r') as f:
                    return json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                pass
        
        return {
            "mappings": [],
            "patterns": {},
            "statistics": {
                "total_mappings": 0,
                "successful_mappings": 0,
                "last_updated": None
            }
        }
    
    def _save_history(self):
        """Save mapping history to file"""
        self.history["statistics"]["last_updated"] = datetime.now().isoformat()
        with open(self.learning_file, 'w') as f:
            json.dump(self.history, f, indent=2)
    
    def _generate_file_signature(self, column_names: List[str], sample_data: str) -> str:
        """Generate a unique signature for the file structure"""
        # Focus on core census columns for more flexible matching
        core_census_columns = [
            'Employee  Name', 'First', 'Coverage Level', 'Gender', 'DOB', 
            'ZIP CODE', 'Home State', 'Job Title', 'W/C Code', 'W/C State', 
            'Healthcare', 'F/T or P/T', 'Annual Pay'
        ]
        
        # Filter to only include core census columns that exist in this file
        relevant_columns = [col for col in column_names if col in core_census_columns]
        
        # Create signature based on relevant columns and data patterns
        signature_data = {
            "core_columns": sorted(relevant_columns),
            "sample_pattern": sample_data[:200]  # First 200 chars of sample data
        }
        signature_str = json.dumps(signature_data, sort_keys=True)
        return hashlib.md5(signature_str.encode()).hexdigest()[:16]
    
    def store_successful_mapping(self, 
                                original_mapping: Dict[str, List[str]], 
                                corrected_mapping: Dict[str, List[str]], 
                                column_names: List[str], 
                                sample_data: str,
                                file_name: str = "unknown"):
        """Store a successful mapping for future learning"""
        
        file_signature = self._generate_file_signature(column_names, sample_data)
        
        mapping_record = {
            "timestamp": datetime.now().isoformat(),
            "file_name": file_name,
            "file_signature": file_signature,
            "column_names": column_names,
            "original_mapping": original_mapping,
            "corrected_mapping": corrected_mapping,
            "corrections_made": self._analyze_corrections(original_mapping, corrected_mapping),
            "success": True
        }
        
        self.history["mappings"].append(mapping_record)
        self.history["statistics"]["total_mappings"] += 1
        self.history["statistics"]["successful_mappings"] += 1
        
        # Update patterns
        self._update_patterns(mapping_record)
        
        # Save to file
        self._save_history()
        
        print(f"🧠 Learning: Stored successful mapping for file signature {file_signature}")
    
    def _analyze_corrections(self, original: Dict, corrected: Dict) -> Dict[str, Any]:
        """Analyze what corrections were made"""
        corrections = {
            "fields_corrected": [],
            "common_patterns": {},
            "total_corrections": 0
        }
        
        for field, corrected_cols in corrected.items():
            original_cols = original.get(field, [])
            if corrected_cols != original_cols:
                corrections["fields_corrected"].append(field)
                corrections["total_corrections"] += 1
                
                # Track common correction patterns
                for orig_col in original_cols:
                    if orig_col not in corrected_cols:
                        corrections["common_patterns"][f"removed_{orig_col}"] = corrections["common_patterns"].get(f"removed_{orig_col}", 0) + 1
                
                for corr_col in corrected_cols:
                    if corr_col not in original_cols:
                        corrections["common_patterns"][f"added_{corr_col}"] = corrections["common_patterns"].get(f"added_{corr_col}", 0) + 1
        
        return corrections
    
    def _update_patterns(self, mapping_record: Dict):
        """Update learned patterns from successful mappings"""
        corrections = mapping_record["corrections_made"]
        
        for field in corrections["fields_corrected"]:
            if field not in self.history["patterns"]:
                self.history["patterns"][field] = {
                    "common_mappings": {},
                    "avoid_mappings": {},
                    "success_count": 0
                }
            
            pattern = self.history["patterns"][field]
            pattern["success_count"] += 1
            
            # Track successful mappings
            corrected_cols = mapping_record["corrected_mapping"][field]
            for col in corrected_cols:
                pattern["common_mappings"][col] = pattern["common_mappings"].get(col, 0) + 1
            
            # Track mappings to avoid
            original_cols = mapping_record["original_mapping"].get(field, [])
            for col in original_cols:
                if col not in corrected_cols:
                    pattern["avoid_mappings"][col] = pattern["avoid_mappings"].get(col, 0) + 1
    
    def get_learning_context(self, column_names: List[str], sample_data: str) -> str:
        """Generate learning context for the LLM prompt"""
        if not self.history["mappings"]:
            return ""
        
        context_parts = []
        
        # Find similar file structures
        current_signature = self._generate_file_signature(column_names, sample_data)
        similar_mappings = []
        
        for mapping in self.history["mappings"]:
            if mapping["file_signature"] == current_signature:
                similar_mappings.append(mapping)
        
        if similar_mappings:
            context_parts.append("🎯 SIMILAR FILE STRUCTURE DETECTED!")
            context_parts.append("Based on previous successful mappings for similar files:")
            
            for mapping in similar_mappings[-3:]:  # Last 3 similar mappings
                context_parts.append(f"✅ File: {mapping['file_name']} (Success)")
                for field, cols in mapping["corrected_mapping"].items():
                    if cols:  # Only show non-empty mappings
                        context_parts.append(f"   {field}: {', '.join(cols)}")
                context_parts.append("")
        
        # Add learned patterns
        if self.history["patterns"]:
            context_parts.append("🧠 LEARNED PATTERNS:")
            context_parts.append("⚠️  CRITICAL: Follow these patterns based on previous successful mappings!")
            context_parts.append("")
            
            for field, pattern in self.history["patterns"].items():
                if pattern["success_count"] >= 1:  # Show patterns with 1+ successes
                    context_parts.append(f"   {field}:")
                    
                    # Show common successful mappings
                    common = sorted(pattern["common_mappings"].items(), key=lambda x: x[1], reverse=True)[:3]
                    if common:
                        context_parts.append(f"     ✅ Often successful: {', '.join([col for col, count in common])}")
                    
                    # Show mappings to avoid
                    avoid = sorted(pattern["avoid_mappings"].items(), key=lambda x: x[1], reverse=True)[:3]
                    if avoid:
                        context_parts.append(f"     ❌ Often incorrect: {', '.join([col for col, count in avoid])}")
                    
                    context_parts.append("")
            
            context_parts.append("🚨 IMPORTANT: Use the 'Often successful' mappings and AVOID the 'Often incorrect' ones!")
            context_parts.append("")
        
        return "\n".join(context_parts)
    
    def get_statistics(self) -> Dict[str, Any]:
        """Get learning system statistics"""
        stats = self.history["statistics"].copy()
        stats["patterns_learned"] = len(self.history["patterns"])
        
        # Count recent mappings (last 24 hours) - fix timestamp comparison
        recent_count = 0
        cutoff_time = datetime.now().timestamp() - 86400  # 24 hours ago
        
        for mapping in self.history["mappings"]:
            try:
                # Convert ISO timestamp to Unix timestamp for comparison
                mapping_time = datetime.fromisoformat(mapping["timestamp"]).timestamp()
                if mapping_time > cutoff_time:
                    recent_count += 1
            except (ValueError, KeyError):
                # Skip invalid timestamps
                continue
        
        stats["recent_mappings"] = recent_count
        
        return stats

# Global learning system instance
learning_system = MappingLearningSystem()
