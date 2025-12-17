import olefile
import zlib
import struct
from datetime import datetime, timedelta

def check_all_subsets(file_path):
    """Check all SUBSETS to find which has valid data"""
    print(f"\n{'='*70}")
    print(f"File: {file_path}")
    print(f"{'='*70}\n")
    
    ole = olefile.OleFileIO(file_path)
    base_date = datetime(1899, 12, 30)
    
    for subset_num in range(2, 8):
        stream_path = ['SUBSETS', str(subset_num), 'DATA']
        try:
            raw = ole.openstream(stream_path).read()
            zlib_off = raw.find(b'\x78\x9c')
            data = zlib.decompress(raw[zlib_off:])
            
            HEADER_SIZE = 628
            MATRIX_STRIDE = 9216
            num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
            matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
            footer = data[matrix_end_offset:]
            
            # Check start time
            OFFSET_START_TIME = 228
            start_time_float = struct.unpack('<d', footer[OFFSET_START_TIME:OFFSET_START_TIME+8])[0]
            start_dt = base_date + timedelta(days=start_time_float)
            
            # Check first temperature
            OFFSET_TEMP_ARRAY = 364
            temp = struct.unpack('<d', footer[OFFSET_TEMP_ARRAY:OFFSET_TEMP_ARRAY+8])[0]
            
            # Check for timestamp at 1012
            ts_offset = 1012 if len(footer) > 1012 else 1006
            if ts_offset + 8 <= len(footer):
                ts_val = struct.unpack('<d', footer[ts_offset:ts_offset+8])[0]
                ts_dt = base_date + timedelta(days=ts_val)
            else:
                ts_dt = None
            
            print(f"SUBSETS/{subset_num}:")
            print(f"  Timepoints: {num_timepoints}")
            print(f"  Start time: {start_dt}")
            print(f"  First temp: {temp:.1f}°C")
            if ts_dt:
                print(f"  First timestamp (offset {ts_offset}): {ts_dt}")
            print()
            
        except Exception as e:
            print(f"SUBSETS/{subset_num}: Error - {e}\n")

check_all_subsets('SlowExperiment.xpt')
check_all_subsets('SweepExperiment.xpt')
