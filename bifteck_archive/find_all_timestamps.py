import olefile
import zlib
import struct
from datetime import datetime, timedelta

def find_all_timestamps(file_path):
    """Find ALL potential timestamp values in the footer"""
    print(f"\n{'='*70}")
    print(f"Analyzing: {file_path}")
    print(f"{'='*70}")
    
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
    
    # Get the known start time
    OFFSET_START_TIME = 228
    start_time_float = struct.unpack('<d', footer[OFFSET_START_TIME:OFFSET_START_TIME+8])[0]
    base_date = datetime(1899, 12, 30)
    start_dt = base_date + timedelta(days=start_time_float)
    
    print(f"Timepoints: {num_timepoints}")
    print(f"Start time: {start_dt} ({start_time_float:.10f})")
    
    # Find ALL Excel dates in footer
    print(f"\n--- All potential timestamps in footer (searching byte-by-byte) ---")
    
    timestamps = []
    for offset in range(0, len(footer) - 8):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        
        # Check if this looks like an Excel date near our start time
        if abs(val - start_time_float) < 1.0:  # Within 1 day
            dt = base_date + timedelta(days=val)
            timestamps.append((offset, val, dt))
    
    print(f"Found {len(timestamps)} potential timestamps:")
    for offset, val, dt in timestamps:
        delta_from_start = (dt - start_dt).total_seconds()
        print(f"  Offset {offset:4d}: {val:.10f} = {dt} (Δ {delta_from_start:.1f}s)")
    
    # Look for clusters at specific offsets
    if len(timestamps) >= num_timepoints:
        print(f"\n--- Analyzing offset patterns ---")
        print(f"Need {num_timepoints} timestamps")
        
        # Get unique offsets (mod 8 to align)
        offset_mods = {}
        for offset, val, dt in timestamps:
            mod = offset % 8
            if mod not in offset_mods:
                offset_mods[mod] = []
            offset_mods[mod].append((offset, val, dt))
        
        print(f"\nGrouped by offset % 8:")
        for mod in sorted(offset_mods.keys()):
            entries = offset_mods[mod]
            print(f"  Mod {mod}: {len(entries)} timestamps")
            if len(entries) >= num_timepoints:
                print(f"    ✓ Has enough timestamps!")
                # Calculate stride
                if len(entries) >= 2:
                    stride = entries[1][0] - entries[0][0]
                    print(f"    Stride appears to be: {stride} bytes")

find_all_timestamps('SlowExperiment.xpt')
find_all_timestamps('SweepExperiment.xpt')
