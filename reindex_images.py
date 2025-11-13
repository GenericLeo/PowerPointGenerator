"""
Re-index existing images to add identifier and grouping metadata
This script updates all images in the index with the new metadata fields
"""

from image_uploader import ImageUploader, ImageIdentifier
import json

def reindex_images():
    """Re-index all existing images with identifier metadata"""
    
    uploader = ImageUploader()
    
    print("=" * 70)
    print("Re-indexing Images")
    print("=" * 70)
    print()
    
    # Get all images
    images = uploader.index.images
    
    if not images:
        print("No images found in index.")
        return
    
    print(f"Found {len(images)} image(s) to re-index...")
    print()
    
    updated_count = 0
    
    for img in images:
        filename = img['filename']
        
        # Extract identifier and numerical prefix from filename
        numerical_prefix, identifier, full_match = ImageIdentifier.extract_identifier_and_number(filename)
        
        # Update metadata
        if 'metadata' not in img:
            img['metadata'] = {}
        
        old_identifier = img['metadata'].get('identifier')
        old_prefix = img['metadata'].get('numerical_prefix')
        
        img['metadata']['identifier'] = identifier
        img['metadata']['numerical_prefix'] = numerical_prefix
        img['metadata']['identifier_match'] = full_match
        
        # Show what changed
        if old_identifier != identifier or old_prefix != numerical_prefix:
            print(f"Updated: {filename}")
            print(f"  Group: {old_prefix or 'None'} → {numerical_prefix or 'None'}")
            print(f"  Type:  {old_identifier or 'None'} → {identifier or 'None'}")
            print()
            updated_count += 1
    
    # Save the updated index
    uploader.index.save_index()
    
    print("=" * 70)
    print(f"✓ Re-indexing complete!")
    print(f"  Total images: {len(images)}")
    print(f"  Updated: {updated_count}")
    print(f"  Unchanged: {len(images) - updated_count}")
    print("=" * 70)
    print()
    print("Please restart the GUI to see the updated sorting and grouping.")

if __name__ == "__main__":
    reindex_images()
