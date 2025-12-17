from bifteck import read_xpt_file

# Test the convenience function
print("Testing read_xpt_file convenience function on DinosaurusRex.xpt")
print("="*80)

df = read_xpt_file('DinosaurusRex.xpt')

print("\n" + "="*80)
print("Results:")
print("="*80)
print(f"Total rows: {len(df)}")
print(f"\nColumns: {df.columns.tolist()}")
print(f"\nUnique subsets: {df['Subset'].unique().tolist()}")
print(f"\nUnique PlateIDs: {df['PlateID'].unique().tolist()}")
print(f"\nUnique Barcodes: {df['Barcode'].unique().tolist()}")
print(f"\nFirst few rows:")
print(df.head(10))
