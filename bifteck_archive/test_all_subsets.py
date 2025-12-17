from bifteck import read_xpt_stream
import olefile

ole = olefile.OleFileIO('DinosaurusRex.xpt')

for s in [1, 2, 3]:
    print(f'\n=== SUBSET {s} ===')
    try:
        df = read_xpt_stream(ole, subset_number=s)
        print(df[['PlateID', 'Barcode', 'Time']].head(2))
    except Exception as e:
        print(f"  [!] Error: {e}")
    print()

ole.close()
