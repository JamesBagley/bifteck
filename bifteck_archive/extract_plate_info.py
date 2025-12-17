import olefile
import struct

def extract_plate_info(file_path, subset_number):
    """Extract plate ID and barcode from HEADER stream"""
    ole = olefile.OleFileIO(file_path)
    
    header_path = ['SUBSETS', str(subset_number), 'HEADER']
    
    if not ole.exists(header_path):
        ole.close()
        return None, None
    
    header = ole.openstream(header_path).read()
    ole.close()
    
    print(f"Header size: {len(header)} bytes")
    print(f"First 100 bytes: {header[:100].hex(' ')}")
    
    # Search for plate ID and barcode
    # Based on pattern: they appear at offsets 62 and 69 in DinosaurusRex
    # Let's find them by looking for the specific pattern
    
    plate_id = None
    barcode = None
    
    # List all length-prefixed strings
    all_strings = []
    for i in range(len(header) - 2):
        length = header[i]
        if 1 <= length <= 50:
            string_start = i + 1
            string_end = string_start + length
            
            if string_end <= len(header):
                try:
                    text = header[string_start:string_end].decode('ascii')
                    if text.isprintable() and len(text) == length:
                        all_strings.append((i, length, text))
                        print(f"Offset {i}: Length={length}, Text='{text}'")
                except:
                    pass
    
    # The last two strings are typically Plate ID and Barcode
    if len(all_strings) >= 2:
        plate_id = all_strings[-2][2]
        barcode = all_strings[-1][2]
    
    return plate_id, barcode

# Test on DinosaurusRex.xpt
print("Testing DinosaurusRex.xpt SUBSET 3:")
print("="*60)
pid, bc = extract_plate_info('DinosaurusRex.xpt', 3)
print(f"\n\nResults:")
print(f"  Plate ID: {pid}")
print(f"  Barcode: {bc}")
