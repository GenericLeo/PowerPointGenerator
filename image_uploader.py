"""
Image Uploader and Indexer for PowerPoint Generator
Supports multiple image formats including TIFF files
Includes intelligent sorting and grouping based on filename identifiers
"""

import os
import json
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from PIL import Image
import hashlib


class ImageIdentifier:
    """Handles extraction and classification of image identifiers from filenames"""
    
    # Primary identifiers that should be grouped by numerical prefix
    GROUPABLE_IDENTIFIERS = ['UD', 'LD', 'MD', 'UVD', 'PDBSE', 'PDBSE1', 'BSE', 'SE', 'ABF', 'ADF']
    
    # Secondary identifiers that are not grouped by number
    NON_GROUPABLE_IDENTIFIERS = ['Spectrum', 'Spectra', 'Map', 'Maps', 'Electron Image']
    
    # Identifiers that should extract trailing numbers for grouping
    TRAILING_NUMBER_IDENTIFIERS = ['Map', 'Maps', 'Electron Image', 'Spectrum', 'Spectra']
    
    # Spectrum identifiers (for special horizontal layout)
    SPECTRUM_IDENTIFIERS = ['Spectrum', 'Spectra']
    
    ALL_IDENTIFIERS = GROUPABLE_IDENTIFIERS + NON_GROUPABLE_IDENTIFIERS
    
    @staticmethod
    def format_group_label(numerical_prefix: Optional[str], identifier: Optional[str]) -> str:
        """
        Format the group label for display.
        
        Map/Electron Image groups with numeric prefix show as MAP1, MAP2, etc.
        Spectrum groups with numeric prefix show as SPEC1, SPEC2, etc.
        Regular groups with numeric prefix show as 0001, 0002, etc.
        Custom labels are displayed as-is.
        
        Returns formatted group label or empty string if no group
        """
        if not numerical_prefix:
            return ""
        
        # Check if it's a 4-digit numeric code (auto-formatted groups)
        if numerical_prefix.isdigit() and len(numerical_prefix) == 4:
            if identifier in ['Spectrum', 'Spectra']:
                # Spectrum groups: format as SPEC1, SPEC2, etc.
                try:
                    spec_num = int(numerical_prefix)
                    return f"SPEC{spec_num}"
                except (ValueError, TypeError):
                    return numerical_prefix
            elif identifier in ['Map', 'Maps', 'Electron Image']:
                # Map/Electron Image groups: format as MAP1, MAP2, etc.
                try:
                    map_num = int(numerical_prefix)
                    return f"MAP{map_num}"
                except (ValueError, TypeError):
                    return numerical_prefix
            else:
                # Regular groups: return as-is (0001, 0002, etc.)
                return numerical_prefix
        else:
            # Custom label: return as-is
            return numerical_prefix
    
    @staticmethod
    def extract_identifier_and_number(filename: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """
        Extract identifier and numerical prefix from filename.
        
        Returns:
            Tuple of (numerical_prefix, identifier, full_match)
            Example: "0001 PDBSE image.tif" -> ("0001", "PDBSE", "0001 PDBSE")
                    "Spectrum_Analysis.jpg" -> (None, "Spectrum", "Spectrum")
        """
        # Remove file extension
        name_without_ext = os.path.splitext(filename)[0]
        
        # Try to find any identifier in the filename
        # Sort identifiers by length (longest first) to match more specific ones first
        sorted_identifiers = sorted(ImageIdentifier.ALL_IDENTIFIERS, key=len, reverse=True)
        
        for identifier in sorted_identifiers:
            # Case-insensitive search with flexible word boundaries (space, underscore, dash, or string boundaries)
            # Also allow for characters like parentheses or alphanumeric suffixes after the identifier
            pattern = re.compile(r'(?:^|[\s_\-])(' + re.escape(identifier) + r')(?:[\s_\-\(\d]|[a-z]+\d*|$)', re.IGNORECASE)
            match = pattern.search(name_without_ext)
            
            if match:
                # Found an identifier
                found_identifier = match.group(1)
                
                # Check if this is a groupable identifier
                if identifier in ImageIdentifier.GROUPABLE_IDENTIFIERS:
                    # Look for numerical prefix before the identifier
                    # Pattern: digits followed by optional separators (space, dash, underscore)
                    # Also look for patterns like: prefix_0001_1_PDBSE or 0001_PDBSE
                    num_pattern = re.compile(r'(\d{4})[\s\-_]+\d*[\s\-_]*' + re.escape(identifier) + r'(?:[\s_\-\(\d]|[a-z]+\d*|$)', re.IGNORECASE)
                    num_match = num_pattern.search(name_without_ext)
                    
                    if num_match:
                        numerical_prefix = num_match.group(1)  # Already 4 digits
                        full_match = num_match.group(0).strip()
                        return (numerical_prefix, identifier, full_match)
                    
                    # Simpler pattern for adjacent numbers
                    num_pattern2 = re.compile(r'(\d+)[\s\-_]+' + re.escape(identifier) + r'(?:[\s_\-\(\d]|[a-z]+\d*|$)', re.IGNORECASE)
                    num_match2 = num_pattern2.search(name_without_ext)
                    
                    if num_match2:
                        numerical_prefix = num_match2.group(1).zfill(4)  # Pad with zeros to 4 digits
                        full_match = num_match2.group(0).strip()
                        return (numerical_prefix, identifier, full_match)
                    
                    # Also try to find number after the identifier (e.g., "ABF 0100")
                    num_pattern_after = re.compile(re.escape(identifier) + r'[\s\-_]+(\d+)', re.IGNORECASE)
                    num_match_after = num_pattern_after.search(name_without_ext)
                    
                    if num_match_after:
                        numerical_prefix = num_match_after.group(1).zfill(4)
                        full_match = num_match_after.group(0)
                        return (numerical_prefix, identifier, full_match)
                    
                    # Identifier found but no number - still valid
                    return (None, identifier, found_identifier)
                else:
                    # Non-groupable identifier (Spectrum, Map, etc.)
                    # Keep original identifier - don't normalize
                    # For Map/Maps/Electron Image, look for a number after the identifier or at the end
                    if identifier in ['Map', 'Maps', 'Electron Image']:
                        # First, try to find number right after the identifier: "Map Data 1_1" or "Electron Image 3_1"
                        # Pattern: identifier followed by optional words, then space/separator and number (before any underscore suffix)
                        after_id_pattern = re.compile(
                            re.escape(identifier) + r'(?:\s+\w+)*?[_\s\-]+(\d+)(?:_\d+)?(?:\s|$)',
                            re.IGNORECASE
                        )
                        after_match = after_id_pattern.search(name_without_ext)
                        
                        if after_match:
                            numerical_prefix = after_match.group(1).zfill(4)
                            return (numerical_prefix, identifier, found_identifier)
                        
                        # Fallback: look for the last number before any trailing underscore suffix (like "_1")
                        # This catches patterns like "xyz 3_1" where 3 is the group number
                        fallback_pattern = re.compile(r'[_\s\-](\d+)(?:_\d+)?\s*$')
                        fallback_match = fallback_pattern.search(name_without_ext)
                        
                        if fallback_match:
                            numerical_prefix = fallback_match.group(1).zfill(4)
                            return (numerical_prefix, identifier, found_identifier)
                    
                    elif identifier in ['Spectrum', 'Spectra']:
                        # Spectrum files - look for trailing number
                        trailing_num_pattern = re.compile(r'[_\s\-]?(\d+)(?:_\d+)?\s*$')
                        trailing_match = trailing_num_pattern.search(name_without_ext)
                        
                        if trailing_match:
                            numerical_prefix = trailing_match.group(1).zfill(4)
                            return (numerical_prefix, identifier, found_identifier)
                    
                    # No grouping if no trailing number found
                    return (None, identifier, found_identifier)
        
        # No identifier found
        return (None, None, None)
    
    @staticmethod
    def get_sort_key(img_info: Dict) -> Tuple:
        """
        Generate a sort key for an image based on its metadata.
        
        Sort order:
        1. Identifier category (main groupable vs Map groups vs ungrouped)
        2. Numerical prefix (grouped images together within category)
        3. Identifier type (predefined order)
        4. Filename (alphabetical)
        """
        metadata = img_info.get('metadata', {})
        numerical_prefix = metadata.get('numerical_prefix', '')
        identifier = metadata.get('identifier', '')
        filename = img_info.get('filename', '')
        
        # Determine the category for sorting
        # Category 0: Main groupable identifiers (UD, LD, PDBSE, etc.)
        # Category 1: Map-like types (Map, Maps, Electron Image)
        # Category 2: Spectrum types (Spectrum, Spectra)
        # Category 3: Ungrouped items (unknown files)
        if identifier in ImageIdentifier.GROUPABLE_IDENTIFIERS:
            category = 0
        elif identifier in ['Map', 'Maps', 'Electron Image']:
            category = 1  # Map-like groups are separate
        elif identifier in ['Spectrum', 'Spectra']:
            category = 2  # Spectrum groups are separate
        else:
            category = 3  # Ungrouped
        
        # Define priority order for identifiers within their category
        identifier_priority = {id_name: idx for idx, id_name in enumerate(ImageIdentifier.GROUPABLE_IDENTIFIERS)}
        identifier_priority.update({id_name: 100 + idx for idx, id_name in enumerate(ImageIdentifier.NON_GROUPABLE_IDENTIFIERS)})
        
        # Get priority (default to 999 if not found)
        priority = identifier_priority.get(identifier, 999)
        
        # If no numerical prefix, put it at the end within its category
        if not numerical_prefix:
            numerical_prefix = 'zzzzz'
        
        return (category, numerical_prefix, priority, filename.lower())


class ImageIndex:
    """Manages the indexing of uploaded images"""
    
    def __init__(self, index_file: str = "image_index.json"):
        self.index_file = index_file
        self.images: List[Dict] = []
        self.load_index()
    
    def load_index(self):
        """Load existing index from file"""
        if os.path.exists(self.index_file):
            try:
                with open(self.index_file, 'r') as f:
                    self.images = json.load(f)
            except json.JSONDecodeError:
                self.images = []
    
    def save_index(self):
        """Save index to file"""
        with open(self.index_file, 'w') as f:
            json.dump(self.images, f, indent=2)
    
    def add_image(self, filepath: str, metadata: Optional[Dict] = None) -> Dict:
        """Add an image to the index with metadata"""
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Image file not found: {filepath}")
        
        # Get image information
        try:
            with Image.open(filepath) as img:
                width, height = img.size
                format_type = img.format
                mode = img.mode
        except Exception as e:
            raise ValueError(f"Error reading image: {e}")
        
        # Calculate file hash for uniqueness
        file_hash = self._calculate_hash(filepath)
        
        # Check if image already indexed
        for img in self.images:
            if img['hash'] == file_hash:
                return img
        
        # Extract identifier and numerical prefix from filename
        filename = os.path.basename(filepath)
        numerical_prefix, identifier, full_match = ImageIdentifier.extract_identifier_and_number(filename)
        
        # Merge provided metadata with extracted metadata
        combined_metadata = metadata.copy() if metadata else {}
        combined_metadata['identifier'] = identifier
        combined_metadata['numerical_prefix'] = numerical_prefix
        combined_metadata['identifier_match'] = full_match
        
        # Create image entry
        image_entry = {
            'id': len(self.images) + 1,
            'filename': filename,
            'filepath': os.path.abspath(filepath),
            'hash': file_hash,
            'format': format_type,
            'mode': mode,
            'width': width,
            'height': height,
            'size_bytes': os.path.getsize(filepath),
            'added_date': datetime.now().isoformat(),
            'metadata': combined_metadata
        }
        
        self.images.append(image_entry)
        self.save_index()
        return image_entry
    
    def _calculate_hash(self, filepath: str) -> str:
        """Calculate SHA256 hash of file"""
        sha256_hash = hashlib.sha256()
        with open(filepath, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    
    def get_image(self, image_id: int) -> Optional[Dict]:
        """Retrieve image by ID"""
        for img in self.images:
            if img['id'] == image_id:
                return img
        return None
    
    def get_all_images(self) -> List[Dict]:
        """Get all indexed images, sorted by group and identifier"""
        sorted_images = sorted(self.images, key=ImageIdentifier.get_sort_key)
        return sorted_images
    
    def get_images_by_group(self, numerical_prefix: str) -> List[Dict]:
        """Get all images with a specific numerical prefix"""
        return [img for img in self.images 
                if img.get('metadata', {}).get('numerical_prefix') == numerical_prefix]
    
    def get_images_by_identifier(self, identifier: str) -> List[Dict]:
        """Get all images with a specific identifier"""
        return [img for img in self.images 
                if img.get('metadata', {}).get('identifier') == identifier]
    
    def get_all_groups(self) -> List[str]:
        """Get all unique numerical prefixes (groups)"""
        groups = set()
        for img in self.images:
            prefix = img.get('metadata', {}).get('numerical_prefix')
            if prefix:
                groups.add(prefix)
        return sorted(list(groups))
    
    def get_grouped_images(self) -> Dict[str, List[Dict]]:
        """
        Get images organized by numerical prefix groups.
        Returns a dictionary where keys are numerical prefixes and values are lists of images.
        """
        grouped = {}
        for img in self.images:
            prefix = img.get('metadata', {}).get('numerical_prefix')
            if prefix:
                if prefix not in grouped:
                    grouped[prefix] = []
                grouped[prefix].append(img)
            else:
                # Images without a numerical prefix go into 'ungrouped'
                if 'ungrouped' not in grouped:
                    grouped['ungrouped'] = []
                grouped['ungrouped'].append(img)
        
        # Sort images within each group by identifier
        for group in grouped.values():
            group.sort(key=ImageIdentifier.get_sort_key)
        
        return grouped
    
    def remove_image(self, image_id: int) -> bool:
        """Remove image from index"""
        for i, img in enumerate(self.images):
            if img['id'] == image_id:
                self.images.pop(i)
                self.save_index()
                return True
        return False
    
    def search_images(self, query: str) -> List[Dict]:
        """Search images by filename or metadata"""
        results = []
        query_lower = query.lower()
        for img in self.images:
            if (query_lower in img['filename'].lower() or 
                query_lower in str(img['metadata']).lower()):
                results.append(img)
        return results
    
    def get_stats(self) -> Dict:
        """Get statistics about indexed images"""
        if not self.images:
            return {
                'total_images': 0,
                'total_size_mb': 0,
                'formats': {}
            }
        
        total_size = sum(img['size_bytes'] for img in self.images)
        formats = {}
        for img in self.images:
            fmt = img['format']
            formats[fmt] = formats.get(fmt, 0) + 1
        
        return {
            'total_images': len(self.images),
            'total_size_mb': round(total_size / (1024 * 1024), 2),
            'formats': formats
        }


class ImageUploader:
    """Handles uploading and organizing images"""
    
    SUPPORTED_FORMATS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.webp'}
    
    def __init__(self, upload_dir: str = "uploaded_images"):
        self.upload_dir = upload_dir
        self.index = ImageIndex()
        self._ensure_upload_dir()
    
    def _ensure_upload_dir(self):
        """Create upload directory if it doesn't exist"""
        os.makedirs(self.upload_dir, exist_ok=True)
    
    def is_supported_format(self, filepath: str) -> bool:
        """Check if file format is supported"""
        ext = Path(filepath).suffix.lower()
        return ext in self.SUPPORTED_FORMATS
    
    def upload_image(self, filepath: str, copy_to_upload_dir: bool = False, 
                    metadata: Optional[Dict] = None) -> Dict:
        """
        Upload and index an image
        
        Args:
            filepath: Path to the image file
            copy_to_upload_dir: If True, copy image to upload directory
            metadata: Optional metadata dictionary
        
        Returns:
            Dictionary with image information
        """
        if not self.is_supported_format(filepath):
            raise ValueError(f"Unsupported image format. Supported: {self.SUPPORTED_FORMATS}")
        
        target_path = filepath
        
        if copy_to_upload_dir:
            filename = os.path.basename(filepath)
            target_path = os.path.join(self.upload_dir, filename)
            
            # Handle duplicate filenames
            counter = 1
            while os.path.exists(target_path):
                name, ext = os.path.splitext(filename)
                target_path = os.path.join(self.upload_dir, f"{name}_{counter}{ext}")
                counter += 1
            
            # Copy file
            import shutil
            shutil.copy2(filepath, target_path)
        
        # Add to index
        return self.index.add_image(target_path, metadata)
    
    def upload_multiple_images(self, filepaths: List[str], 
                              copy_to_upload_dir: bool = False,
                              metadata: Optional[Dict] = None) -> List[Dict]:
        """Upload multiple images at once"""
        results = []
        for filepath in filepaths:
            try:
                result = self.upload_image(filepath, copy_to_upload_dir, metadata)
                results.append(result)
            except Exception as e:
                print(f"Error uploading {filepath}: {e}")
        return results
    
    def list_images(self) -> List[Dict]:
        """List all uploaded images"""
        return self.index.get_all_images()
    
    def get_image_info(self, image_id: int) -> Optional[Dict]:
        """Get information about a specific image"""
        return self.index.get_image(image_id)
    
    def remove_image(self, image_id: int, delete_file: bool = False) -> bool:
        """Remove image from index and optionally delete file"""
        img = self.index.get_image(image_id)
        if img and delete_file:
            try:
                os.remove(img['filepath'])
            except Exception as e:
                print(f"Error deleting file: {e}")
        
        return self.index.remove_image(image_id)
    
    def get_statistics(self) -> Dict:
        """Get statistics about uploaded images"""
        return self.index.get_stats()


def main():
    """Example usage of the ImageUploader"""
    uploader = ImageUploader()
    
    print("=" * 60)
    print("Image Uploader for PowerPoint Generator")
    print("=" * 60)
    print("\nSupported formats:", ", ".join(ImageUploader.SUPPORTED_FORMATS))
    print("\nCommands:")
    print("  upload <filepath> - Upload an image")
    print("  list - List all uploaded images")
    print("  info <id> - Get info about an image")
    print("  remove <id> - Remove image from index")
    print("  stats - Show statistics")
    print("  search <query> - Search images")
    print("  quit - Exit")
    print("=" * 60)
    
    while True:
        try:
            command = input("\n> ").strip().split(maxsplit=1)
            
            if not command:
                continue
            
            cmd = command[0].lower()
            
            if cmd == 'quit':
                break
            
            elif cmd == 'upload':
                if len(command) < 2:
                    print("Usage: upload <filepath>")
                    continue
                filepath = command[1].strip()
                result = uploader.upload_image(filepath, copy_to_upload_dir=True)
                print(f"✓ Uploaded: {result['filename']} (ID: {result['id']})")
                print(f"  Format: {result['format']}, Size: {result['width']}x{result['height']}")
            
            elif cmd == 'list':
                images = uploader.list_images()
                if not images:
                    print("No images uploaded yet.")
                else:
                    print(f"\nTotal images: {len(images)}")
                    for img in images:
                        print(f"  [{img['id']}] {img['filename']} - {img['format']} "
                              f"({img['width']}x{img['height']})")
            
            elif cmd == 'info':
                if len(command) < 2:
                    print("Usage: info <id>")
                    continue
                img_id = int(command[1])
                img = uploader.get_image_info(img_id)
                if img:
                    print(f"\nImage Details:")
                    for key, value in img.items():
                        if key != 'metadata' or value:
                            print(f"  {key}: {value}")
                else:
                    print(f"Image with ID {img_id} not found.")
            
            elif cmd == 'remove':
                if len(command) < 2:
                    print("Usage: remove <id>")
                    continue
                img_id = int(command[1])
                if uploader.remove_image(img_id):
                    print(f"✓ Removed image {img_id}")
                else:
                    print(f"Image with ID {img_id} not found.")
            
            elif cmd == 'stats':
                stats = uploader.get_statistics()
                print(f"\nStatistics:")
                print(f"  Total images: {stats['total_images']}")
                print(f"  Total size: {stats['total_size_mb']} MB")
                print(f"  Formats: {stats['formats']}")
            
            elif cmd == 'search':
                if len(command) < 2:
                    print("Usage: search <query>")
                    continue
                query = command[1]
                results = uploader.index.search_images(query)
                if results:
                    print(f"\nFound {len(results)} image(s):")
                    for img in results:
                        print(f"  [{img['id']}] {img['filename']}")
                else:
                    print("No images found.")
            
            else:
                print(f"Unknown command: {cmd}")
        
        except Exception as e:
            print(f"Error: {e}")


if __name__ == "__main__":
    main()
