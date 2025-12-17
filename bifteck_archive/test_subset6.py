import olefile
import zlib
import struct
from datetime import datetime, timedelta

# Test extraction with discovered offsets for SUBSET 6
def test_subset6_extraction(file_path):
    ole = olefile.OleFileIO(file_path)
    stream_path = ['SUBSETS', '6', 'DATA']
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    HEADER_SIZE = 628
    MATRIX_STRIDE = 9216
    num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
    footer = data[matrix_end_offset:]
    base_date = datetime(1899, 12, 30)
    
    print(f"Testing SUBSET 6 extraction from {file_path}\n")
    
    # Extract temperatures at offset 348, stride 24
    print("Temperatures:")
    temps = []
    for i in range(num_timepoints):
        offset = 348 + (i * 24)
        temp = struct.unpack('<d', footer[offset:offset+8])[0]
        temps.append(temp)
        print(f"  [{i}] {temp:.1f}°C")
    
    # Extract timestamps at offset 956, stride 86
    print("\nTimestamps:")
    timestamps = []
    for i in range(num_timepoints):
        offset = 956 + (i * 86)
        ts = struct.unpack('<d', footer[offset:offset+8])[0]
        dt = base_date + timedelta(days=ts)
        timestamps.append(dt)
        print(f"  [{i}] {dt}")
    
    # Calculate intervals
    if len(timestamps) >= 2:
        print("\nIntervals:")
        for i in range(len(timestamps)-1):
            interval = (timestamps[i+1] - timestamps[i]).total_seconds()
            print(f"  {i}->{i+1}: {interval:.1f}s ({interval/3600:.3f}h)")

test_subset6_extraction('SweepExperiment.xpt')
