import olefile
import zlib
import struct
from datetime import datetime, timedelta

def deep_search_second_timestamp(filename, subset_num):
    """Search entire footer for second timestamp"""
    ole = olefile.OleFileIO(filename)
    stream_path = ['SUBSETS', str(subset_num), 'DATA']
    
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    HEADER_SIZE = 628
    MATRIX_STRIDE = 9216
    num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
    footer = data[matrix_end_offset:]
    
    # Find first timestamp
    ole_marker = b'\x00\x03OLE\x00'
    ole_search = footer.find(ole_marker)
    first_ts_offset = ole_search + 12
    first_ts = struct.unpack('<d', footer[first_ts_offset:first_ts_offset+8])[0]
    base_date = datetime(1899, 12, 30)
    first_dt = base_date + timedelta(days=first_ts)
    
    print(f"{filename} - SUBSET {subset_num}")
    print(f"First timestamp at offset {first_ts_offset}: {first_dt}")
    print(f"Footer size: {len(footer)} bytes")
    print(f"\nSearching ENTIRE footer for timestamps near first one...")
    
    found_count = 0
    for offset in range(0, len(footer) - 8):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        # Look for timestamps within a few hours of the first one
        if 40000 < val < 50000:
            dt = base_date + timedelta(days=val)
            time_diff = abs((dt - first_dt).total_seconds())
            if time_diff < 7200:  # Within 2 hours
                print(f"  Offset {offset:4d} (Δ{offset-first_ts_offset:4d}): {dt} (Δ{time_diff:.0f}s from first)")
                found_count += 1
    
    print(f"\nTotal timestamps found: {found_count}")
    ole.close()

deep_search_second_timestamp('Multiplate test.xpt', 2)
