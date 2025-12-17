import olefile
import zlib
import struct
from datetime import datetime, timedelta

def find_timestamp_array(file_path):
    """Search for an array of timestamps in the footer"""
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
    print(f"Footer size: {len(footer)} bytes")
    print(f"Start time: {start_dt}")
    
    # Search for arrays of Excel dates
    print(f"\n--- Searching for timestamp arrays ---")
    
    # Try different strides (8, 16, 24, 32 bytes)
    for stride in [8, 16, 24, 32, 40]:
        print(f"\nTrying stride {stride} bytes:")
        
        for start_offset in range(0, min(500, len(footer) - num_timepoints * stride), 8):
            timestamps = []
            valid = True
            
            for i in range(num_timepoints):
                offset = start_offset + (i * stride)
                if offset + 8 > len(footer):
                    valid = False
                    break
                
                val = struct.unpack('<d', footer[offset:offset+8])[0]
                
                # Check if this looks like an Excel date near our start time
                if not (abs(val - start_time_float) < 1.0):  # Within 1 day
                    valid = False
                    break
                
                timestamps.append(val)
            
            if valid and len(timestamps) == num_timepoints:
                # Convert to actual times
                times = [base_date + timedelta(days=t) for t in timestamps]
                
                # Calculate intervals
                intervals = []
                for i in range(len(times) - 1):
                    delta = (times[i+1] - times[i]).total_seconds()
                    intervals.append(delta)
                
                print(f"  ✓ FOUND at offset {start_offset}, stride {stride}!")
                print(f"    Timestamps:")
                for i, (ts, dt) in enumerate(zip(timestamps, times)):
                    if i < 5 or i >= num_timepoints - 2:  # Show first 5 and last 2
                        print(f"      [{i:2d}] {ts:.10f} = {dt}")
                    elif i == 5:
                        print(f"      ...")
                
                if intervals:
                    print(f"\n    Intervals (seconds):")
                    for i, interval in enumerate(intervals):
                        if i < 5 or i >= len(intervals) - 2:
                            print(f"      {i}->{i+1}: {interval:.1f}s ({interval/3600:.3f}h)")
                        elif i == 5:
                            print(f"      ...")
                    
                    avg_interval = sum(intervals) / len(intervals)
                    min_interval = min(intervals)
                    max_interval = max(intervals)
                    print(f"\n    Average: {avg_interval:.1f}s ({avg_interval/3600:.3f}h)")
                    print(f"    Range: {min_interval:.1f}s to {max_interval:.1f}s")
                
                return start_offset, stride

find_timestamp_array('SlowExperiment.xpt')
find_timestamp_array('SweepExperiment.xpt')
