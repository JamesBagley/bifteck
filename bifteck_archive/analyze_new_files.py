import olefile
import zlib
import struct
from datetime import datetime, timedelta

def analyze_file_structure(filename):
    """Deep analysis of file structure using landmarks"""
    print(f"\n{'='*80}")
    print(f"ANALYZING: {filename}")
    print(f"{'='*80}")
    
    ole = olefile.OleFileIO(filename)
    
    # Get all data streams
    data_streams = []
    for entry in ole.listdir():
        if entry[0] == 'SUBSETS' and len(entry) >= 3 and entry[2] == 'DATA':
            stream_path = ['SUBSETS', entry[1], entry[2]]
            stream_size = ole.get_size(stream_path)
            data_streams.append((entry[1], stream_path, stream_size))
    
    data_streams.sort(key=lambda x: int(x[0]))
    
    print(f"\nFound {len(data_streams)} data streams")
    
    for subset_num, stream_path, size in data_streams:
        print(f"\n--- SUBSET {subset_num} ({size} bytes) ---")
        
        try:
            raw = ole.openstream(stream_path).read()
            zlib_off = raw.find(b'\x78\x9c')
            
            if zlib_off == -1:
                print("  ❌ No zlib header found")
                continue
                
            data = zlib.decompress(raw[zlib_off:])
            print(f"  Decompressed size: {len(data)} bytes")
            
            # Calculate timepoints
            HEADER_SIZE = 628
            MATRIX_STRIDE = 9216
            
            if len(data) < HEADER_SIZE:
                print(f"  ❌ Too small (< {HEADER_SIZE} bytes)")
                continue
            
            num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
            matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
            footer = data[matrix_end_offset:]
            
            print(f"  Timepoints: {num_timepoints}")
            print(f"  Matrix end: {matrix_end_offset}")
            print(f"  Footer size: {len(footer)} bytes")
            
            if num_timepoints == 0:
                print("  ⚠️ No timepoints (empty data)")
                continue
            
            # Search for temperature landmark: pattern before temps
            # Pattern: 18 00 00 00 02 00 [TP_COUNT] 00 00 00 01 00 00 00
            temp_pattern = struct.pack('<IHHHIH', 0x18, 2, num_timepoints & 0xFFFF, 0, 1, 0)
            temp_search = footer.find(temp_pattern)
            
            if temp_search != -1:
                temp_offset = temp_search + len(temp_pattern)
                print(f"  ✅ Temperature landmark found at offset {temp_search}")
                print(f"  Temperature data starts at offset {temp_offset}")
                
                # Read first temperature
                if temp_offset + 8 <= len(footer):
                    first_temp = struct.unpack('<d', footer[temp_offset:temp_offset+8])[0]
                    print(f"     First temp: {first_temp:.2f}°C")
            else:
                print(f"  ❌ Temperature landmark NOT found")
                # Try brute force search
                for offset in range(0, min(600, len(footer) - 8)):
                    val = struct.unpack('<d', footer[offset:offset+8])[0]
                    if 20 <= val <= 30:
                        print(f"     Manual search found temp at offset {offset}: {val:.2f}°C")
                        break
            
            # Search for timestamp landmark: "OLE" string
            ole_marker = b'\x00\x03OLE\x00'
            ole_search = footer.find(ole_marker)
            
            if ole_search != -1:
                ts_offset = ole_search + 12  # 12 bytes after start of marker
                print(f"  ✅ Timestamp landmark (OLE) found at offset {ole_search}")
                print(f"  Timestamp data starts at offset {ts_offset}")
                
                # Read first timestamp
                if ts_offset + 8 <= len(footer):
                    first_ts = struct.unpack('<d', footer[ts_offset:ts_offset+8])[0]
                    if 1 < first_ts < 100000:
                        base_date = datetime(1899, 12, 30)
                        dt = base_date + timedelta(days=first_ts)
                        print(f"     First timestamp: {dt}")
                    else:
                        print(f"     First timestamp: {first_ts} (invalid)")
            else:
                print(f"  ❌ Timestamp landmark (OLE) NOT found")
                # Try brute force search
                for offset in range(0, min(1200, len(footer) - 8)):
                    val = struct.unpack('<d', footer[offset:offset+8])[0]
                    if 46000 < val < 46010:
                        base_date = datetime(1899, 12, 30)
                        dt = base_date + timedelta(days=val)
                        print(f"     Manual search found timestamp at offset {offset}: {dt}")
                        break
            
            # Show first 100 bytes of footer
            print(f"\n  First 100 bytes of footer:")
            for i in range(0, min(100, len(footer)), 16):
                hex_str = ' '.join(f'{b:02x}' for b in footer[i:i+16])
                print(f"    {i:04d}: {hex_str}")
                
        except Exception as e:
            print(f"  ❌ Error: {e}")
    
    ole.close()

# Analyze all new files
for filename in ['Experiment1.xpt', '384_OD_test.xpt', 'Multiplate test.xpt', 'Multiplate test 2.xpt']:
    analyze_file_structure(filename)
