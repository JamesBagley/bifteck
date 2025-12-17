import olefile
import zlib
import struct
from datetime import datetime, timedelta

def detailed_footer_analysis(file_path):
    """Deep dive into footer structure"""
    print(f"\n{'='*60}")
    print(f"Analyzing: {file_path}")
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
    print(f"Footer size: {len(footer)} bytes")
    print(f"Timepoints: {num_timepoints}")
    
    # Extract start time
    OFFSET_START_TIME = 228
    start_time_float = struct.unpack('<d', footer[OFFSET_START_TIME:OFFSET_START_TIME+8])[0]
    base_date = datetime(1899, 12, 30)
    start_dt = base_date + timedelta(days=start_time_float)
    print(f"Start time: {start_dt}")
    
    # Look at values around the start time offset
    print(f"\n--- Values near start time (offset {OFFSET_START_TIME}) ---")
    for offset in range(OFFSET_START_TIME-40, OFFSET_START_TIME+80, 8):
        if offset < 0 or offset + 8 > len(footer):
            continue
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        
        # Try to interpret as different things
        interpretations = []
        
        # Excel date?
        if 40000 < val < 50000:
            dt = base_date + timedelta(days=val)
            interpretations.append(f"Date: {dt}")
        
        # Seconds?
        if 0 < val < 86400:
            interpretations.append(f"Seconds: {val:.1f} ({val/3600:.2f}h)")
        
        # Fractional days?
        if 0 < val < 1:
            secs = val * 24 * 3600
            interpretations.append(f"Days: {val:.6f} ({secs:.1f}s, {secs/3600:.2f}h)")
        
        marker = "  <<<" if offset == OFFSET_START_TIME else ""
        if interpretations:
            print(f"  {offset:4d}: {val:20.10f} | {' | '.join(interpretations)}{marker}")
        elif -100 < val < 100 and val != 0:
            print(f"  {offset:4d}: {val:20.10f}{marker}")

# Analyze both files
detailed_footer_analysis('SlowExperiment.xpt')
detailed_footer_analysis('SweepExperiment.xpt')
