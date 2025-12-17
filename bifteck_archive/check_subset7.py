import olefile
import zlib
import struct
from datetime import datetime, timedelta

def analyze_subset_7(file_path):
    """Analyze SUBSETS/7 structure"""
    print(f"\n{'='*70}")
    print(f"File: {file_path}")
    print(f"{'='*70}")
    
    ole = olefile.OleFileIO(file_path)
    stream_path = ['SUBSETS', '7', 'DATA']
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
    
    print(f"Total size: {len(data)} bytes")
    print(f"Timepoints: {num_timepoints}")
    print(f"Matrix end: {matrix_end_offset}")
    print(f"Footer size: {len(footer)} bytes")
    
    # Find start time
    OFFSET_START_TIME = 228
    start_time_float = struct.unpack('<d', footer[OFFSET_START_TIME:OFFSET_START_TIME+8])[0]
    start_dt = base_date + timedelta(days=start_time_float)
    print(f"Start time: {start_dt}")
    
    # Check temperatures
    OFFSET_TEMP_ARRAY = 364
    TEMP_STRIDE = 24
    print(f"\nTemperatures:")
    for i in range(num_timepoints):
        offset = OFFSET_TEMP_ARRAY + (i * TEMP_STRIDE)
        if offset + 8 > len(footer):
            print(f"  [{i}] Out of bounds!")
            break
        temp = struct.unpack('<d', footer[offset:offset+8])[0]
        print(f"  [{i}] {temp:.1f}°C")
    
    # Try to find timestamps starting around offset 1012
    print(f"\nSearching for timestamp array around offset 1012:")
    for start_offset in range(1000, 1030, 2):
        timestamps = []
        for stride in [86, 84, 88, 90]:
            timestamps = []
            valid = True
            for i in range(num_timepoints):
                offset = start_offset + (i * stride)
                if offset + 8 > len(footer):
                    valid = False
                    break
                val = struct.unpack('<d', footer[offset:offset+8])[0]
                if not (abs(val - start_time_float) < 1.0):
                    valid = False
                    break
                timestamps.append(val)
            
            if valid and len(timestamps) == num_timepoints:
                print(f"\n  ✓ Found at offset {start_offset}, stride {stride}!")
                times = [base_date + timedelta(days=t) for t in timestamps]
                for i, dt in enumerate(times):
                    print(f"    [{i:2d}] {dt}")
                return start_offset, stride
    
    print("  Not found with simple pattern")

analyze_subset_7('SlowExperiment.xpt')
analyze_subset_7('SweepExperiment.xpt')
