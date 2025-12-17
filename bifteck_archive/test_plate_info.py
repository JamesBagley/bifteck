import olefile

def extract_plate_info(file_path, subset_number):
    """Extract plate ID and barcode from HEADER stream"""
    ole = olefile.OleFileIO(file_path)
    
    header_path = ['SUBSETS', str(subset_number), 'HEADER']
    
    if not ole.exists(header_path):
        ole.close()
        return None, None
    
    header = ole.openstream(header_path).read()
    ole.close()
    
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
                except:
                    pass
    
    # The last two strings are typically Plate ID and Barcode
    plate_id = None
    barcode = None
    if len(all_strings) >= 2:
        plate_id = all_strings[-2][2]
        barcode = all_strings[-1][2]
    
    return plate_id, barcode

# Test on multiple files
print("Testing plate info extraction:\n")

test_files = [
    ('DinosaurusRex.xpt', [1, 2, 3]),
]

# Check if archive files exist
import os
archive_path = 'bifteck_archive'
if os.path.exists(archive_path):
    archive_files = [
        ('SlowExperiment.xpt', [2, 3, 4, 5, 6, 7]),
        ('SweepExperiment.xpt', [2, 3, 4, 5, 6, 7]),
        ('Multiplate test.xpt', [2, 3]),
    ]
    test_files.extend([(os.path.join(archive_path, f), s) for f, s in archive_files])

for file_path, subsets in test_files:
    if not os.path.exists(file_path):
        continue
    
    print(f"\n{file_path}:")
    print("-" * 60)
    
    for subset in subsets:
        plate_id, barcode = extract_plate_info(file_path, subset)
        if plate_id or barcode:
            print(f"  SUBSET {subset}: ID='{plate_id}', Barcode='{barcode}'")
