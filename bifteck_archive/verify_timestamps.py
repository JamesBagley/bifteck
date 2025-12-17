import olefile
import zlib
import struct
from datetime import datetime, timedelta

def extract_timestamps_properly(file_path):
    """Extract timestamps using the discovered pattern"""
    print(f"\n{'='*70}")
    print(f"File: {file_path}")
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
    base_date = datetime(1899, 12, 30)
    
    # Extract timestamps at offset 1012 with stride 86
    TIMESTAMP_OFFSET = 1012
    TIMESTAMP_STRIDE = 86
    
    print(f"Extracting {num_timepoints} timestamps:")
    print(f"  Starting offset: {TIMESTAMP_OFFSET}")
    print(f"  Stride: {TIMESTAMP_STRIDE} bytes\n")
    
    timestamps = []
    for i in range(num_timepoints):
        offset = TIMESTAMP_OFFSET + (i * TIMESTAMP_STRIDE)
        if offset + 8 > len(footer):
            print(f"  Warning: Offset {offset} out of bounds!")
            break
        
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        dt = base_date + timedelta(days=val)
        timestamps.append(dt)
        print(f"  [{i:2d}] Offset {offset:4d}: {dt}")
    
    # Calculate intervals
    if len(timestamps) >= 2:
        print(f"\n Intervals between measurements:")
        for i in range(len(timestamps) - 1):
            delta = (timestamps[i+1] - timestamps[i]).total_seconds()
            print(f"  {i}->{i+1}: {delta:.1f} seconds ({delta/3600:.4f} hours)")
    
    return timestamps

ts1 = extract_timestamps_properly('SlowExperiment.xpt')
ts2 = extract_timestamps_properly('SweepExperiment.xpt')
