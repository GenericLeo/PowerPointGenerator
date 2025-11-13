"""
Test script to demonstrate identifier extraction from filenames
"""

from image_uploader import ImageIdentifier

# Test filenames (including actual user filenames)
test_filenames = [
    "NiCoCr_HT1250c_48h_0001_1_PDBSE1(COMP).tif",
    "NiCoCr_HT1250c_48h_0001_2_UD.tif",
    "NiCoCr_HT1250c_48h_0002_1_PDBSE1(COMP).tif",
    "NiCoCr_HT1250c_48h_0002_2_UD.tif",
    "NiCoCr_HT1250c_48h_0003_2_UDmod2.tif",
    "NiCoCr_HT1250c_48h_0003_2_UDmod.tif",
    "NiCoCr_HT1250c_48h_0003_1_PDBSE1(COMP) mod.tif",
    "Co K_alpha_1 Map Data 1.tif",
    "Co K_alpha_1 Map Data 2.tif",
    "O K_alpha_1 Map Data 3.tif",
    "Cr K_alpha_1 Map Data 2.tif",
    "Electron Image 1.tif",
    "Electron Image 3.tif",
    "Electron Image 4.tif",
    "Spectrum 1.tiff",
    "Spectrum 2.tiff",
    "Spectrum 5.tiff",
    "0001 UD.tif",
    "0001 LD.tif",
    "0001 PDBSE.tif",
]

print("=" * 80)
print("Identifier Extraction Test")
print("=" * 80)
print()

for filename in test_filenames:
    prefix, identifier, full_match = ImageIdentifier.extract_identifier_and_number(filename)
    print(f"Filename: {filename:35} → Group: {prefix or 'None':6} | Type: {identifier or 'None':8} | Match: {full_match or 'None'}")

print()
print("=" * 80)
print("Groupable Identifiers:", ", ".join(ImageIdentifier.GROUPABLE_IDENTIFIERS))
print("Non-Groupable Identifiers:", ", ".join(ImageIdentifier.NON_GROUPABLE_IDENTIFIERS))
print("=" * 80)
