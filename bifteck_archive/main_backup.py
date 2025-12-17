import olefile
import zlib
import struct
import numpy as np
import pandas as pd
import os
import re
from datetime import datetime, timedelta

# --- STRUCTURAL CONSTANTS (DERIVED FROM MAP) ---
HEADER_SIZE = 628

# Matrix Constants
MATRIX_STRIDE = 9216  # 384 wells * 3 doubles * 8 bytes

# Footer Constants (Relative to End of Matrix)
OFFSET_START_TIME = 228
OFFSET_TEMP_ARRAY = 364
TEMP_STRIDE = 24  # Double + 16 bytes padding/flags

# Timestamp array in footer
OFFSET_TIMESTAMP_ARRAY = 1012
TIMESTAMP_STRIDE = 86  # Bytes between consecutive timestamps


def parse_interval_from_txt(txt_file_path):
    """
    Parse the time interval from a BioTek TXT export file.
    
    Args:
        txt_file_path: Path to a .txt export file
    
    Returns:
        int: Interval in seconds, or None if not found
    """
    try:
        with open(txt_file_path, 'r') as f:
            # Read first few lines to find the time row
            for i, line in enumerate(f):
                if i > 50:  # Only check first 50 lines
                    break
                if line.startswith('Time,'):
                    # Parse time values: Time,0:00:00,1:00:01,2:00:01,...
                    parts = line.strip().split(',')
                    if len(parts) >= 3:
                        # Parse first two time values
                        time1_str = parts[1]  # e.g., "0:00:00"
                        time2_str = parts[2]  # e.g., "1:00:01"
                        
                        # Convert H:MM:SS to seconds
                        def parse_time(t):
                            components = t.split(':')
                            if len(components) == 3:
                                h, m, s = map(int, components)
                                return h * 3600 + m * 60 + s
                            return None
                        
                        t1 = parse_time(time1_str)
                        t2 = parse_time(time2_str)
                        
                        if t1 is not None and t2 is not None:
                            interval = t2 - t1
                            if interval > 0:
                                return interval
    except:
        pass
    return None


def extract_biotek_final(file_path, stream_path=None, subset_number=None):
    """
    Extract data from BioTek XPT files.
    
    Args:
        file_path: Path to the .xpt file
        stream_path: OLE stream path (list like ['SUBSETS', '2', 'DATA'])
        subset_number: Which SUBSET to extract (2-7), default: last non-empty one
    
    Returns:
        pandas DataFrame with extracted data
    """
    print(f"--- Extracting: {file_path} ---")
    ole = olefile.OleFileIO(file_path)

    # 1. Auto-detect stream if not provided
    if stream_path is None:
        if subset_number is None:
            # Find last non-empty DATA stream (typically SUBSETS/7)
            for num in range(7, 1, -1):
                test_stream = ['SUBSETS', str(num), 'DATA']
                if ole.exists(test_stream):
                    stream_path = test_stream
                    break
        else:
            stream_path = ['SUBSETS', str(subset_number), 'DATA']
        
        if stream_path is None:
            raise ValueError("No DATA stream found in XPT file")
        print(f"  [+] Using stream: {'/'.join(stream_path)}")
    
    # 2. Load and Decompress Data
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    if zlib_off == -1:
        raise ValueError("No zlib compressed data found")
    data = zlib.decompress(raw[zlib_off:])

    # 3. Calculate Geometry
    total_len = len(data)
    data_body_len = total_len - HEADER_SIZE

    # Calculate N Timepoints (Integer Division)
    num_timepoints = data_body_len // MATRIX_STRIDE
    if num_timepoints == 0:
        raise ValueError(f"No complete timepoints found (data too small)")

    # Calculate exact boundary where Matrix ends and Footer begins
    matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)

    print(f"  [+] Timepoints: {num_timepoints}")
    print(f"  [+] Matrix end: {matrix_end_offset} / {total_len} bytes")

    # 4. Extract Data Matrix
    matrix_bytes = data[HEADER_SIZE : matrix_end_offset]

    # Load as array of doubles
    total_doubles = len(matrix_bytes) // 8
    matrix_floats = struct.unpack(f'<{total_doubles}d', matrix_bytes)

    # Reshape: (N, 384, 3) where each well has (Value, Flag, Status)
    # Total: 1152 doubles per timepoint = 384 wells * 3 values
    np_matrix = np.array(matrix_floats).reshape(num_timepoints, 1152)

    # Extract only the measurement values (every 3rd element starting at 0)
        # Try to find corresponding TXT file
        base_name = os.path.splitext(file_path)[0]
        dir_path = os.path.dirname(file_path)
        
        # Look for any TXT files with similar names
        txt_files = []
        if os.path.exists(dir_path):
            for filename in os.listdir(dir_path):
                if filename.lower().endswith('.txt') and base_name.lower() in filename.lower():
                    txt_files.append(os.path.join(dir_path, filename))
        
        # Try to parse interval from first matching TXT file
        for txt_file in txt_files:
            parsed_interval = parse_interval_from_txt(txt_file)
            if parsed_interval:
                interval_seconds = parsed_interval
                print(f"  [+] Detected interval from {os.path.basename(txt_file)}: {interval_seconds}s ({interval_seconds/3600:.2f}h)")
                break
        
        # Fallback to default
        if interval_seconds is None:
            interval_seconds = 3600
            print(f"  [!] Using default interval: {interval_seconds}s ({interval_seconds/3600:.2f}h)")
            print(f"      (Specify interval_seconds parameter or provide TXT export if different)")
    # Label Wells A1..P24 (16 rows × 24 columns = 384 wells)
    well_labels = [f"{r}{c}" for r in "ABCDEFGHIJKLMNOP" for c in range(1, 25)]
    df.columns = well_labels

    # 5. Extract Temperatures
    temps = []
    current_offset = matrix_end_offset + OFFSET_TEMP_ARRAY
    footer_end = len(data)
    
    for i in range(num_timepoints):
        if current_offset + 8 > footer_end:
            print(f"  [!] Warning: Temperature data truncated at timepoint {i}")
            # Pad with NaN for missing values
            temps.extend([np.nan] * (num_timepoints - i))
            break

        val = struct.unpack('<d', data[current_offset : current_offset + 8])[0]
        
        # Validate temperature is reasonable (10-40°C typical range)
        if not (10 <= val <= 40):
            print(f"  [!] Warning: Unusual temperature {val:.1f}°C at timepoint {i}")
        
        temps.append(val)
        current_offset += TEMP_STRIDE

    df.insert(0, "Temperature", temps)
    print(f"  [+] Temperatures: {len(temps)} readings")

    # 6. Extract Start Time and Calculate Intervals
    start_time_offset = matrix_end_offset + OFFSET_START_TIME
    if start_time_offset + 8 > len(data):
        raise ValueError("Cannot read start time from footer")
    
    start_time_float = struct.unpack('<d', data[start_time_offset : start_time_offset+8])[0]

    # Excel Serial Date Conversion (days since Dec 30, 1899)
    base_date = datetime(1899, 12, 30)
    start_dt = base_date + timedelta(days=start_time_float)

    prDetermine interval
    if interval_seconds is None:
        # Default to 1 hour if not specified
        interval_seconds = 3600
        print(f"  [!] Using default interval: {interval_seconds}s ({interval_seconds/3600:.2f}h)")
        print(f"      (Specify interval_seconds parameter if different)")
    else:
        print(f"  [+] Using interval: {interval_seconds}s ({interval_seconds/3600:.2f}h
        interval_seconds = 3600
        print(f"  [!] Using default interval: {interval_seconds}s (1 hour)")

    # Generate Time Vectors
    timestamps = []
    elapsed_labels = []

    for i in range(num_timepoints):
        # Absolute timestamp
        curr_dt = start_dt + timedelta(seconds=i * interval_seconds)
        timestamps.append(curr_dt)

        # Elapsed time formatted as H:MM:SS
        total_seconds = i * interval_seconds
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        elapsed_labels.append(f"{hours}:{minutes:02d}:{seconds:02d}")

    df.insert(0, "Time_Abs", timestamps)
    df.insert(0, "Time", elapsed_labels)

    # 7. Clean Data
    # Remove columns that are all zeros (empty wells)
    df = df.loc[:, (df != 0).any(axis=0)]
    
    print(f"  [+] Final shape: {df.shape[0]} timepoints × {df.shape[1]} columns")
    print(f"  [+] Extraction complete!\n")
    
    return df
