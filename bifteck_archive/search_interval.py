import olefile
import zlib
import struct
from datetime import datetime, timedelta

def find_interval_in_footer(file_path, expected_interval_seconds):
    """Search footer for interval information"""
    print(f"\n{'='*60}")
    print(f"File: {file_path}")
    print(f"Expected interval: {expected_interval_seconds}s ({expected_interval_seconds/3600:.2f}h)")
    print(f"{'='*60}")
    
    ole = olefile.OleFileIO(file_path)
    stream_path = ['SUBSETS', '2', 'DATA']
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    HEADER_SIZE = 628
    MATRIX_STRIDE = 9216
    data_body_len = len(data) - HEADER_SIZE
    num_timepoints = data_body_len // MATRIX_STRIDE
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
    
    footer = data[matrix_end_offset:]
    
    # Search for our expected interval
    # As seconds
    # As fractional days
    interval_days = expected_interval_seconds / (24 * 3600)
    
    print(f"\nSearching for:")
    print(f"  - {expected_interval_seconds} (seconds)")
    print(f"  - {interval_days:.10f} (fractional days)")
    print(f"  - {interval_days * 24:.10f} (fractional hours)")
    
    found = []
    for offset in range(0, len(footer) - 8, 8):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        
        # Check if it matches (with tolerance)
        if abs(val - expected_interval_seconds) < 1:
            found.append((offset, val, "seconds"))
        elif abs(val - interval_days) < 0.0000001:
            found.append((offset, val, "fractional days"))
        elif abs(val - interval_days * 24) < 0.001:
            found.append((offset, val, "fractional hours"))
    
    if found:
        print(f"\nFound {len(found)} potential matches:")
        for offset, val, interpretation in found:
            print(f"  Offset {offset:4d}: {val:20.10f} ({interpretation})")
    else:
        print("\n  No matches found")
    
    # Also just dump all non-zero, reasonable-scale values
    print(f"\n--- All non-zero footer values (first 800 bytes) ---")
    for offset in range(0, min(800, len(footer) - 8), 8):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        if val != 0 and abs(val) < 1000000:
            print(f"  {offset:4d}: {val:20.10f}")

# SlowExperiment has 1 hour + 1 second intervals based on txt file
find_interval_in_footer('SlowExperiment.xpt', 3601)

# SweepExperiment has exactly 1 hour intervals
find_interval_in_footer('SweepExperiment.xpt', 3600)
