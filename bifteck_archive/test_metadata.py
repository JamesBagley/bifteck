from bifteck import XptFile

# Test plate metadata extraction
with XptFile('DinosaurusRex.xpt') as xpt:
    plate_id, barcode = xpt._extract_plate_metadata(3)
    print(f"✓ Plate ID: {plate_id}")
    print(f"✓ Barcode: {barcode}")

# Test full extraction
print("\nFull extraction:")
with XptFile('DinosaurusRex.xpt') as xpt:
    df = xpt.extract_stream(subset_number=3)
    print(f"\nDataFrame columns: {df.columns.tolist()[:6]}...")
    print(f"\nFirst row metadata:")
    print(f"  PlateID: {df['PlateID'].iloc[0]}")
    print(f"  Barcode: {df['Barcode'].iloc[0]}")
