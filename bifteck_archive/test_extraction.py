import olefile
import zlib
import struct

def analyze_xpt_structure(file_path, stream_path='Data/OD600(600)'):
    """Analyze the XPT file structure to understand timing and temperature data"""
    print(f"\n=== Analyzing: {file_path} ===")
    ole = olefile.OleFileIO(file_path)
    
    # List available streams
    print("Available streams:")
    for stream in ole.listdir():
        print(f"  {'/'.join(stream)}")
    
    # Analyze first DATA stream (SUBSETS/2/DATA)
    stream_path = ['SUBSETS', '2', 'DATA']
    print(f"\nUsing stream: {'/'.join(stream_path)}")
    
    # Load and decompress data
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    print(f"Total decompressed size: {len(data)} bytes")
    
    # Try to find timing information in header
    HEADER_SIZE = 628
    header = data[:HEADER_SIZE]
    
    # Look for patterns that might indicate timing
    print("\n--- Looking for timing information ---")
    for offset in range(0, min(300, len(header)), 8):
        try:
            val = struct.unpack('<d', header[offset:offset+8])[0]
            # Excel dates are typically between 40000-50000 for 2010-2030
            if 40000 < val < 50000:
                print(f"Offset {offset}: {val} (potential Excel date)")
        except:
            pass
    
    # Calculate body structure
    MATRIX_STRIDE = 9216
    data_body_len = len(data) - HEADER_SIZE
    num_timepoints = data_body_len // MATRIX_STRIDE
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
    
    print(f"\nCalculated timepoints: {num_timepoints}")
    print(f"Matrix ends at: {matrix_end_offset}")
    print(f"Footer size: {len(data) - matrix_end_offset}")
    
    # Analyze footer section for temperatures and timing
    print("\n--- Analyzing footer section ---")
    footer = data[matrix_end_offset:]
    
    # Look for temperature-like values (20-30 typical range)
    print("\nSearching for temperature patterns (expecting ~20-30°C):")
    TEMP_STRIDE = 24
    OFFSET_TEMP_ARRAY = 364
    
    if len(footer) > OFFSET_TEMP_ARRAY:
        current_offset = OFFSET_TEMP_ARRAY
        temps_found = []
        for i in range(min(15, num_timepoints)):
            if current_offset + 8 > len(footer):
                break
            val = struct.unpack('<d', footer[current_offset:current_offset+8])[0]
            temps_found.append(val)
            current_offset += TEMP_STRIDE
        
        print(f"Temperature values at offset {OFFSET_TEMP_ARRAY}, stride {TEMP_STRIDE}:")
        print(temps_found[:10])
    
    # Look for time intervals
    print("\n--- Searching for time interval information ---")
    for offset in range(0, min(500, len(footer)), 8):
        try:
            val = struct.unpack('<d', footer[offset:offset+8])[0]
            # Look for values that could be seconds (3600 for 1h) or fractional days
            if 3500 < val < 3700:  # Around 1 hour in seconds
                print(f"Footer offset {offset}: {val} seconds ({val/3600:.2f} hours)")
            elif 0.04 < val < 0.05:  # Around 1 hour in days (1/24)
                print(f"Footer offset {offset}: {val} days ({val*24:.2f} hours)")
        except:
            pass

# Analyze both files
analyze_xpt_structure('SlowExperiment.xpt')
analyze_xpt_structure('SweepExperiment.xpt')
