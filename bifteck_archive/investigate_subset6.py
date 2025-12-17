import olefile
import zlib
import struct
from datetime import datetime, timedelta

def deep_analysis_subset6(file_path):
    """Deep analysis of SUBSETS/6 to find where the data is"""
    print(f"\n{'='*70}")
    print(f"Deep Analysis: {file_path} - SUBSETS/6")
    print(f"{'='*70}\n")
    
    ole = olefile.OleFileIO(file_path)
    stream_path = ['SUBSETS', '6', 'DATA']
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    HEADER_SIZE = 628
    MATRIX_STRIDE = 9216
    num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
    footer = data[matrix_end_offset:]
    
    print(f"Total data: {len(data)} bytes")
    print(f"Timepoints: {num_timepoints}")
    print(f"Footer size: {len(footer)} bytes")
    
    base_date = datetime(1899, 12, 30)
    
    # Search ENTIRE footer for Excel dates
    print(f"\n--- Searching for Excel dates (Dec 2025 = ~46003) ---")
    found_dates = []
    for offset in range(0, len(footer) - 8):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        if 46000 < val < 46010:  # Dec 2025 range
            dt = base_date + timedelta(days=val)
            found_dates.append((offset, val, dt))
    
    print(f"Found {len(found_dates)} potential dates:")
    for offset, val, dt in found_dates[:20]:  # Show first 20
        print(f"  Offset {offset:4d}: {dt}")
    
    # Search for temperature-like values
    print(f"\n--- Searching for temperatures (20-30°C range) ---")
    found_temps = []
    for offset in range(0, len(footer) - 8):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        if 20 <= val <= 30:
            found_temps.append((offset, val))
    
    print(f"Found {len(found_temps)} potential temperatures:")
    for offset, val in found_temps[:20]:  # Show first 20
        print(f"  Offset {offset:4d}: {val:.1f}°C")
    
    # Check if there's a pattern in the date offsets
    if len(found_dates) >= num_timepoints:
        print(f"\n--- Analyzing date offset pattern ---")
        print(f"Need {num_timepoints} timestamps")
        
        # Try to find a stride pattern
        if len(found_dates) >= 2:
            for i in range(min(5, len(found_dates) - 1)):
                stride = found_dates[i+1][0] - found_dates[i][0]
                print(f"  Offset {found_dates[i][0]} -> {found_dates[i+1][0]}: stride = {stride}")

deep_analysis_subset6('SweepExperiment.xpt')
