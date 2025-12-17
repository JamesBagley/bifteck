import olefile
import zlib
import struct
from datetime import datetime, timedelta

def analyze_timing(file_path):
    """Extract timing information from XPT file"""
    print(f"\n=== Analyzing timing in: {file_path} ===")
    ole = olefile.OleFileIO(file_path)
    
    # Load DATA stream
    stream_path = ['SUBSETS', '2', 'DATA']
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    # Calculate structure
    HEADER_SIZE = 628
    MATRIX_STRIDE = 9216
    data_body_len = len(data) - HEADER_SIZE
    num_timepoints = data_body_len // MATRIX_STRIDE
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
    
    # Extract start time
    OFFSET_START_TIME = 228
    start_time_offset = matrix_end_offset + OFFSET_START_TIME
    start_time_float = struct.unpack('<d', data[start_time_offset:start_time_offset+8])[0]
    base_date = datetime(1899, 12, 30)
    start_dt = base_date + timedelta(days=start_time_float)
    
    print(f"Start time: {start_dt}")
    
    # Look for additional timestamps in footer to calculate interval
    print("\n--- Searching for timestamp patterns in footer ---")
    footer = data[matrix_end_offset:]
    
    # Search for Excel date values near the start time
    potential_times = []
    for offset in range(0, min(1000, len(footer)), 8):
        try:
            val = struct.unpack('<d', footer[offset:offset+8])[0]
            # Excel dates for dates near Dec 2025 should be around 45000-46000
            if 45000 < val < 47000:
                dt = base_date + timedelta(days=val)
                # Only include if it's within reasonable range
                if abs((dt - start_dt).total_seconds()) < 24*3600*30:  # Within 30 days
                    potential_times.append((offset, val, dt))
        except:
            pass
    
    if potential_times:
        print(f"\nFound {len(potential_times)} potential timestamps:")
        for i, (offset, val, dt) in enumerate(potential_times[:15]):
            print(f"  Offset {offset:4d}: {dt} (value: {val:.8f})")
        
        # Calculate intervals between consecutive timestamps
        if len(potential_times) >= 2:
            print("\n--- Calculated intervals ---")
            for i in range(min(5, len(potential_times)-1)):
                delta_days = potential_times[i+1][1] - potential_times[i][1]
                delta_seconds = delta_days * 24 * 3600
                print(f"  Time {i} -> {i+1}: {delta_seconds:.1f} seconds ({delta_seconds/3600:.2f} hours)")
    
    # Also check the HEADER stream
    print("\n--- Checking HEADER stream ---")
    try:
        header_stream = ['SUBSETS', '2', 'HEADER']
        header_raw = ole.openstream(header_stream).read()
        print(f"Header size: {len(header_raw)} bytes")
        
        # Look for timing info in header
        for offset in range(0, min(500, len(header_raw)), 8):
            try:
                val = struct.unpack('<d', header_raw[offset:offset+8])[0]
                # Look for seconds (3600) or fractional days (1/24)
                if 3500 < val < 3700:
                    print(f"  Offset {offset}: {val} seconds ({val/3600:.3f} hours)")
                elif 0.04 < val < 0.05:
                    print(f"  Offset {offset}: {val} days ({val*24:.3f} hours, {val*24*3600:.1f} seconds)")
                elif 0 < val < 0.01 and val > 0.001:  # Fractional hours
                    print(f"  Offset {offset}: {val} (potential {val*24*3600:.1f} seconds)")
            except:
                pass
    except Exception as e:
        print(f"Could not read header: {e}")

analyze_timing('SlowExperiment.xpt')
analyze_timing('SweepExperiment.xpt')
