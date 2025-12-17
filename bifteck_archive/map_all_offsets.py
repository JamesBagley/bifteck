import olefile
import zlib
import struct
from datetime import datetime, timedelta

def find_offsets_for_subset(file_path, subset_num):
    """Find temperature and timestamp offsets for a specific subset"""
    ole = olefile.OleFileIO(file_path)
    stream_path = ['SUBSETS', str(subset_num), 'DATA']
    
    if not ole.exists(stream_path):
        return None
    
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    HEADER_SIZE = 628
    MATRIX_STRIDE = 9216
    num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
    footer = data[matrix_end_offset:]
    base_date = datetime(1899, 12, 30)
    
    # Find temperature offset (search for 20-30°C values)
    temp_candidates = []
    for offset in range(0, min(600, len(footer) - 8)):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        if 20 <= val <= 30:
            temp_candidates.append(offset)
    
    temp_offset = temp_candidates[0] if temp_candidates else None
    
    # Find timestamp offset (search for Dec 2025 dates)
    ts_candidates = []
    for offset in range(0, min(1200, len(footer) - 8)):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        if 46000 < val < 46010:
            ts_candidates.append(offset)
    
    # Find the one that gives us num_timepoints valid timestamps with stride 86
    ts_offset = None
    for candidate in ts_candidates:
        valid_count = 0
        for i in range(num_timepoints):
            offset = candidate + (i * 86)
            if offset + 8 > len(footer):
                break
            val = struct.unpack('<d', footer[offset:offset+8])[0]
            if 46000 < val < 46010:
                valid_count += 1
        if valid_count >= num_timepoints - 1:  # Allow one invalid at end
            ts_offset = candidate
            break
    
    return {
        'subset': subset_num,
        'timepoints': num_timepoints,
        'temp_offset': temp_offset,
        'ts_offset': ts_offset,
        'has_data': temp_offset is not None and ts_offset is not None
    }

# Check all subsets in both files
for file in ['SlowExperiment.xpt', 'SweepExperiment.xpt']:
    print(f"\n{'='*70}")
    print(f"{file}")
    print(f"{'='*70}\n")
    print(f"{'Subset':<10} {'TPs':<6} {'Temp Offset':<15} {'TS Offset':<15} {'Status'}")
    print("-" * 70)
    
    for subset in range(2, 8):
        info = find_offsets_for_subset(file, subset)
        if info:
            status = "✓ VALID" if info['has_data'] else "✗ EMPTY"
            temp_str = str(info['temp_offset']) if info['temp_offset'] else "N/A"
            ts_str = str(info['ts_offset']) if info['ts_offset'] else "N/A"
            print(f"{subset:<10} {info['timepoints']:<6} {temp_str:<15} {ts_str:<15} {status}")
