# PowerPoint Generator - Image Uploader

A Python application for uploading and indexing images (including TIFF files) for PowerPoint generation.

## Features

- ✅ Upload multiple image formats (JPG, PNG, TIFF, GIF, BMP, WEBP)
- ✅ **Intelligent sorting and grouping** based on filename identifiers
- ✅ **Automatic identifier detection** (UD, LD, MD, UVD, PDBSE, BSE, SE, ABF, ADF, Spectrum, Spectra, Map, Maps)
- ✅ **Numerical grouping** - images with same prefix are grouped together (e.g., 0001 UD, 0001 LD, 0001 PDBSE)
- ✅ Automatic image indexing with metadata
- ✅ TIFF file support
- ✅ Image search and filtering functionality
- ✅ Duplicate detection (via SHA256 hashing)
- ✅ Statistics and reporting
- ✅ JSON-based index storage
- ✅ Modern GUI interface with preview

## Installation

1. **Set up Python virtual environment** (recommended):
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On macOS/Linux
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### GUI Application (Recommended)

Run the graphical interface:
```bash
python gui_app.py
```

The GUI provides:
- 📁 Drag & drop or file browser for uploading images
- 📂 Bulk upload from folders
- 🔍 Real-time search functionality
- 🏷️ **Filter by group number** (e.g., show only 0001 group)
- 🔖 **Filter by identifier type** (e.g., show only BSE images)
- 👁️ Image preview thumbnails
- 📊 Statistics dashboard with group counts
- 🗑️ Easy image management
- 📋 Sortable columns (Group, ID, Type, Filename, Format, etc.)

### Command Line Interface

Run the interactive CLI:
```bash
python image_uploader.py
```

**Available Commands:**
- `upload <filepath>` - Upload an image file
- `list` - List all uploaded images
- `info <id>` - Get detailed information about an image
- `remove <id>` - Remove image from index
- `stats` - Show statistics about uploaded images
- `search <query>` - Search images by filename or metadata
- `quit` - Exit the application

### Python API

```python
from image_uploader import ImageUploader

# Initialize uploader
uploader = ImageUploader(upload_dir="uploaded_images")

# Upload a single image
result = uploader.upload_image(
    filepath="path/to/image.tiff",
    copy_to_upload_dir=True,
    metadata={"description": "My TIFF image", "category": "medical"}
)

# Upload multiple images
filepaths = ["image1.jpg", "image2.tiff", "image3.png"]
results = uploader.upload_multiple_images(
    filepaths=filepaths,
    copy_to_upload_dir=True
)

# List all images
images = uploader.list_images()
for img in images:
    print(f"{img['id']}: {img['filename']} - {img['format']}")

# Get image info
image_info = uploader.get_image_info(image_id=1)

# Search images
results = uploader.index.search_images("medical")

# Get statistics
stats = uploader.get_statistics()
print(f"Total images: {stats['total_images']}")
print(f"Total size: {stats['total_size_mb']} MB")
```

## Image Index Structure

Each uploaded image is indexed with the following information:

```json
{
  "id": 1,
  "filename": "example.tiff",
  "filepath": "/absolute/path/to/example.tiff",
  "hash": "sha256_hash_of_file",
  "format": "TIFF",
  "mode": "RGB",
  "width": 1920,
  "height": 1080,
  "size_bytes": 2048576,
  "added_date": "2025-11-13T10:30:00",
  "metadata": {
    "description": "Custom metadata",
    "tags": ["tag1", "tag2"]
  }
}
```

## Image Naming Convention

The system automatically detects and groups images based on their filenames:

### Groupable Identifiers
These identifiers are grouped by numerical prefix:
- **UD** - Upper Detector
- **LD** - Lower Detector  
- **MD** - Middle Detector
- **UVD** - UV Detector
- **PDBSE** - Primary Detector BSE
- **BSE** - Backscattered Electrons
- **SE** - Secondary Electrons
- **ABF** - Annular Bright Field
- **ADF** - Annular Dark Field

**Example grouping:**
```
0001 UD.tif      → Group: 0001, Type: UD
0001 LD.tif      → Group: 0001, Type: LD
0001 PDBSE.tif   → Group: 0001, Type: PDBSE
0002 UD.jpg      → Group: 0002, Type: UD
0002 LD.jpg      → Group: 0002, Type: LD
```

### Non-Groupable Identifiers
These identifiers are catalogued but not grouped by number:
- **Spectrum** / **Spectra** - Spectroscopy data
- **Map** / **Maps** - Mapping data

**Supported filename formats:**
- `0001 UD.tif` (space separator)
- `0025_BSE.tif` (underscore separator)
- `0010-UVD.tif` (dash separator)
- `Sample ABF 0100.tif` (number after identifier)

- JPEG (.jpg, .jpeg)
- PNG (.png)
- TIFF (.tiff, .tif)
- GIF (.gif)
- BMP (.bmp)
- WEBP (.webp)

## Supported Image Formats

```
PowerPointGenerator/
├── image_uploader.py       # Main application code
├── requirements.txt        # Python dependencies
├── image_index.json       # Auto-generated index file
├── uploaded_images/       # Directory for uploaded images
└── README.md             # This file
```

## Next Steps

This is the foundation for your PowerPoint generator. Future enhancements could include:

1. **GUI Interface** - Add a Tkinter or web-based interface
2. **PowerPoint Generation** - Use python-pptx to create presentations
3. **Image Processing** - Add filters, cropping, resizing
4. **Batch Operations** - Process multiple images at once
5. **Export Options** - Export index to CSV or Excel
6. **Cloud Storage** - Support for cloud-based image storage

## Example Workflow

```bash
# Start the application
python image_uploader.py

# Upload images
> upload /path/to/image1.tiff
✓ Uploaded: image1.tiff (ID: 1)
  Format: TIFF, Size: 1920x1080

> upload /path/to/image2.jpg
✓ Uploaded: image2.jpg (ID: 2)
  Format: JPEG, Size: 1280x720

# List all images
> list
Total images: 2
  [1] image1.tiff - TIFF (1920x1080)
  [2] image2.jpg - JPEG (1280x720)

# Get statistics
> stats
Statistics:
  Total images: 2
  Total size: 4.5 MB
  Formats: {'TIFF': 1, 'JPEG': 1}
```

## License

This project is open source and available for modification.
