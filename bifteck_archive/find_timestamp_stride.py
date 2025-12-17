import olefile
import zlib
import struct
from datetime import datetime, timedelta

def analyze_timestamp_stride(filename, subset_num):
    """Analyze timestamp array structure"""
    ole = olefile.OleFileIO(filename)
    stream_path = ['SUBSETS', str(subset_num), 'DATA']
    
    if not ole.exists(stream_path):
        return
    
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    HEADER_SIZE = 628
    MATRIX_STRIDE = 9216
    num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
    footer = data[matrix_end_offset:]
    
    # Find OLE marker
    ole_marker = b'\x00\x03OLE\x00'
    ole_search = footer.find(ole_marker)
    
    if ole_search == -1:
        print(f"{filename} SUBSET {subset_num}: No OLE marker found")
        return
    
    ts_offset = ole_search + 12
    base_date = datetime(1899, 12, 30)
    
    print(f"\n{filename} - SUBSET {subset_num} ({num_timepoints} timepoints)")
    print(f"OLE marker at: {ole_search}, TS start at: {ts_offset}")
    
    # Read all timestamps with different stride guesses
    for stride_guess in [24, 48, 86, 96, 120]:
        print(f"\n  Trying stride = {stride_guess}:")
        all_valid = True
        for i in range(num_timepoints):
            offset = ts_offset + (i * stride_guess)
            if offset + 8 <= len(footer):
                val = struct.unpack('<d', footer[offset:offset+8])[0]
                if 40000 < val < 50000:
                    dt = base_date + timedelta(days=val)
                    print(f"    [{i}] offset {offset}: {dt}")
                else:
                    print(f"    [{i}] offset {offset}: INVALID ({val})")
                    all_valid = False
            else:
                print(f"    [{i}] offset {offset}: OUT OF BOUNDS")
                all_valid = False
        
        if all_valid:
            print(f"  ✅ Stride {stride_guess} looks correct!")
    
    ole.close()

# Analyze files with different timepoint counts
analyze_timestamp_stride('Multiplate test.xpt', 2)
analyze_timestamp_stride('Multiplate test.xpt', 3)
analyze_timestamp_stride('SlowExperiment.xpt', 2)  # 11 timepoints
analyze_timestamp_stride('SlowExperiment.xpt', 4)  # 10 timepoints
