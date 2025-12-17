import olefile
import zlib
import struct
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import sys

# --- STRUCTURAL CONSTANTS (DERIVED FROM MAP) ---
HEADER_SIZE = 628

# Matrix Constants
MATRIX_STRIDE = 9216  # 384 wells * 3 doubles * 8 bytes

# Footer Constants (Relative to End of Matrix)
OFFSET_START_TIME = 228
TEMP_STRIDE = 24  # Double + 16 bytes padding/flags
TIMESTAMP_STRIDE = 86  # Bytes between consecutive timestamps

# Offsets vary based on number of timepoints
# These will be determined dynamically


class XptFile:
    """
    Wrapper class for BioTek XPT files that provides extraction capabilities
    while maintaining access to the file path for error reporting.
    
    Usage:
        with XptFile('experiment.xpt') as xpt:
            df = xpt.extract_stream(subset_number=2)
            # or
            all_data = xpt.extract_all_streams()
    """
    
    def __init__(self, file_path):
        """
        Initialize XptFile wrapper.
        
        Args:
            file_path: Path to the XPT file
        """
        self.file_path = file_path
        self.ole = None
    
    def __enter__(self):
        """Open the OLE file and return self for context manager."""
        self.ole = olefile.OleFileIO(self.file_path)
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Close the OLE file when exiting context manager."""
        if self.ole:
            self.ole.close()
        return False
    
    def _extract_plate_metadata(self, subset_number):
        """
        Extract plate ID and barcode from HEADER stream.
        
        Args:
            subset_number: Which SUBSET to extract from
        
        Returns:
            tuple: (plate_id, barcode) or (None, None) if not found
        """
        header_path = ['SUBSETS', str(subset_number), 'HEADER']
        
        if not self.ole.exists(header_path):
            return None, None
        
        header = self.ole.openstream(header_path).read()
        
        # Find all length-prefixed ASCII strings in header
        all_strings = []
        for i in range(len(header) - 2):
            length = header[i]
            if 1 <= length <= 50:
                string_start = i + 1
                string_end = string_start + length
                
                if string_end <= len(header):
                    try:
                        text = header[string_start:string_end].decode('ascii')
                        if text.isprintable() and len(text) == length:
                            all_strings.append(text)
                    except:
                        pass
        
        # Last two strings are typically Plate ID and Barcode
        plate_id = all_strings[-2] if len(all_strings) >= 2 else None
        barcode = all_strings[-1] if len(all_strings) >= 1 else None
        
        return plate_id, barcode

    def extract_stream(self, stream_path=None, subset_number=None):
        """
        Extract data from BioTek XPT files.
        
        Args:
            stream_path: OLE stream path (list like ['SUBSETS', '2', 'DATA'])
            subset_number: Which SUBSET to extract (2-7), default: 2
        
        Returns:
            pandas DataFrame with extracted data
        """

        if subset_number is None and stream_path is None:
            raise ValueError("Either stream_path or subset_number must be provided")

        # 1. Auto-detect stream if not provided
        if stream_path is None:
            stream_path = ['SUBSETS', str(subset_number), 'DATA']
            if not self.ole.exists(stream_path):
                raise ValueError(f"Stream SUBSETS/{subset_number}/DATA not found in XPT file")
            print(f"  [+] Using stream: {'/'.join(stream_path)}", file=sys.stderr)
        
        # 1a. Extract plate metadata from HEADER
        plate_id, barcode = self._extract_plate_metadata(subset_number)
        if plate_id:
            print(f"  [+] Plate ID: {plate_id}", file=sys.stderr)
        if barcode:
            print(f"  [+] Barcode: {barcode}", file=sys.stderr)
        
        # 2. Load and Decompress Data
        raw = self.ole.openstream(stream_path).read()
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

        print(f"  [+] Timepoints: {num_timepoints}", file=sys.stderr)
        print(f"  [+] Matrix end: {matrix_end_offset} / {total_len} bytes", file=sys.stderr)

        # 4. Extract footer for landmark detection
        footer = data[matrix_end_offset:]
        
        # 4a. Find temperature offset by searching for first reasonable temp (20-30°C)
        # This universal approach works across all file types with different timepoint counts
        OFFSET_TEMP_ARRAY = None
        for offset in range(0, min(600, len(footer) - 8)):
            val = struct.unpack('<d', footer[offset:offset+8])[0]
            if 20 <= val <= 30:
                OFFSET_TEMP_ARRAY = offset
                break
        
        if OFFSET_TEMP_ARRAY is None:
            raise Warning(f"Could not find temperature data in {self.file_path}, {stream_path}, using fallback offset")
            OFFSET_TEMP_ARRAY = 364
        
        # 4b. Find timestamp offset using "OLE" landmark
        # Universal marker: \x00\x03OLE\x00 appears 12 bytes before all timestamp arrays
        OFFSET_TIMESTAMP_ARRAY = None
        ole_marker = b'\x00\x03OLE\x00'
        ole_search = footer.find(ole_marker)
        if ole_search != -1:
            OFFSET_TIMESTAMP_ARRAY = ole_search + 12  # 12 bytes after start of marker
        else:
            raise Warning(f"timestamp marker not found in {self.file_path}, {stream_path}, using fallback offset")
            OFFSET_TIMESTAMP_ARRAY = 974

        # 5. Extract Data Matrix
        matrix_bytes = data[HEADER_SIZE : matrix_end_offset]

        # Load as array of doubles
        total_doubles = len(matrix_bytes) // 8
        matrix_floats = struct.unpack(f'<{total_doubles}d', matrix_bytes)

        # Reshape: (N, 384, 3) where each well has (Value, Flag, Status)
        # Total: 1152 doubles per timepoint = 384 wells * 3 values
        np_matrix = np.array(matrix_floats).reshape(num_timepoints, 1152)

        # Extract only the measurement values (every 3rd element starting at 0)
        measurements = np_matrix[:, 0::3]

        df = pd.DataFrame(measurements)

        # Label Wells A1..P24 (16 rows × 24 columns = 384 wells)
        well_labels = [f"{r}{c}" for r in "ABCDEFGHIJKLMNOP" for c in range(1, 25)]
        df.columns = well_labels

        # 6. Extract Temperatures
        temps = []
        current_offset = OFFSET_TEMP_ARRAY
        
        for i in range(num_timepoints):
            if current_offset + 8 > len(footer):
                raise Warning(f"Temperature data out of bounds at timepoint {i} in {self.file_path}, {stream_path}, using NaN for remaining values")
                # Pad with NaN for missing values
                temps.extend([np.nan] * (num_timepoints - i))
                break

            val = struct.unpack('<d', footer[current_offset : current_offset + 8])[0]
            
            # Validate temperature is reasonable (10-40°C typical range)
            if not (10 <= val <= 40):
                raise Warning(f"Unusual temperature {val:.1f}°C at timepoint {i} in {self.file_path}, {stream_path}")        
            temps.append(val)
            current_offset += TEMP_STRIDE

        df.insert(0, "Temperature", temps)
        print(f"  [+] Temperatures: {len(temps)} readings", file=sys.stderr)

        # 7. Extract Timestamps from Footer
        # The actual timestamps are stored in the footer at a fixed offset
        base_date = datetime(1899, 12, 30)
        
        timestamps = []
        
        for i in range(num_timepoints):
            ts_offset = OFFSET_TIMESTAMP_ARRAY + (i * TIMESTAMP_STRIDE)
            
            if ts_offset + 8 > len(footer):
                raise Warning(f"Timestamp {i} out of bounds at {self.file_path}, {stream_path}, using fallback of 1 second intervals")
                # Fallback: use 1 second intervals from last known time
                if timestamps:
                    last_time = timestamps[-1]
                    timestamps.append(last_time + timedelta(seconds=1))
                else:
                    # Try to get start time from standard offset
                    start_time_offset = OFFSET_START_TIME
                    if start_time_offset + 8 <= len(footer):
                        start_time_float = struct.unpack('<d', footer[start_time_offset:start_time_offset+8])[0]
                        timestamps.append(base_date + timedelta(days=start_time_float))
                    else:
                        raise ValueError("Cannot read timestamps from footer")
                continue
            
            ts_float = struct.unpack('<d', footer[ts_offset:ts_offset+8])[0]
            
            # Validate it's a reasonable Excel date (between year 1900-2100)
            if not (1 < ts_float < 100000):
                print(f"  [!] Warning: Invalid timestamp at index {i} (value: {ts_float}), using fallback", file=sys.stderr)
                if timestamps:
                    last_time = timestamps[-1]
                    timestamps.append(last_time + timedelta(seconds=3600))
                else:
                    raise ValueError(f"First timestamp is invalid: {ts_float}, {self.file_path}, {stream_path}")
                continue
            
            timestamp = base_date + timedelta(days=ts_float)
            timestamps.append(timestamp)
        
        if timestamps:
            print(f"  [+] Start time: {timestamps[0]}", file=sys.stderr)
            print(f"  [+] End time: {timestamps[-1]}", file=sys.stderr)
            
            # Calculate and show interval statistics
            if len(timestamps) >= 2:
                intervals = [(timestamps[i+1] - timestamps[i]).total_seconds() 
                            for i in range(len(timestamps)-1)]
                avg_interval = sum(intervals) / len(intervals)
                print(f"  [+] Average interval: {avg_interval:.1f}s ({avg_interval/3600:.3f}h)", file=sys.stderr)
        
        # Generate elapsed time labels
        elapsed_labels = []
        for i, ts in enumerate(timestamps):
            if i == 0:
                elapsed_labels.append("0:00:00")
            else:
                total_seconds = int((ts - timestamps[0]).total_seconds())
                hours = total_seconds // 3600
                minutes = (total_seconds % 3600) // 60
                seconds = total_seconds % 60
                elapsed_labels.append(f"{hours}:{minutes:02d}:{seconds:02d}")

        df.insert(0, "Time_Abs", timestamps)
        df.insert(0, "Time", elapsed_labels)
        
        # Add plate metadata
        df.insert(0, "Barcode", barcode if barcode else "")
        df.insert(0, "PlateID", plate_id if plate_id else "")

        # 8. Clean Data
        # Remove columns that are all zeros (empty wells)
        df = df.loc[:, (df != 0).any(axis=0)]
        
        print(f"  [+] Final shape: {df.shape[0]} timepoints × {df.shape[1]} columns", file=sys.stderr)
        print(f"  [+] Extraction complete!\n", file=sys.stderr)
        
        return df

def read_xpt_file(file_path):
    """
    Convenience function to extract all data streams from an XPT file.
    
    Args:
        file_path: Path to the XPT file
    
    Returns:
        pandas DataFrame with all extracted data in long format
    """
    with XptFile(file_path) as xpt:
        streams = xpt.ole.listdir()
        data_streams = [s for s in streams if len(s) == 3 and s[0] == 'SUBSETS' and s[2] == 'DATA']
        frames = []
        for stream in data_streams:
            subset_number = stream[1]
            print(f"Processing {file_path} subset {subset_number}...", file=sys.stderr)
            try:
                df = xpt.extract_stream(subset_number=int(subset_number)).assign(File=file_path, Subset=subset_number)
                frames.append(df)
            except Exception as e:
                print(f"  [!] Warning: Error processing {file_path} subset {subset_number}, skipping: {e}", file=sys.stderr)
                continue
        
        if not frames:
            raise ValueError(f"No valid data extracted from {file_path}")
        
        return pd.concat(frames, ignore_index=True).melt(
            id_vars=["File", "Subset", "PlateID", "Barcode", "Time", "Time_Abs", "Temperature"], 
            var_name="Well", 
            value_name="Measurement"
        )


def main():
    """
    Command-line interface for extracting BioTek XPT files.
    Outputs CSV format suitable for importing into R.
    
    Usage:
        python bifteck.py input.xpt
        python bifteck.py input.xpt > output.csv
    """
    import argparse
    import sys
    
    parser = argparse.ArgumentParser(
        description='Extract data from BioTek XPT files to CSV format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python bifteck.py experiment.xpt
  python bifteck.py experiment.xpt > output.csv
  python bifteck.py experiment.xpt --output results.csv
        """
    )
    parser.add_argument('file_path', help='Path to the XPT file to extract')
    parser.add_argument('-o', '--output', help='Output CSV file (default: stdout)', default=None)
    
    args = parser.parse_args()
    
    try:
        # Extract data
        df = read_xpt_file(args.file_path)
        
        # Output as CSV
        if args.output:
            df.to_csv(args.output, index=False)
            print(f"\nData exported to {args.output}", file=sys.stderr)
        else:
            # Print to stdout (diagnostic messages go to stderr)
            print(df.to_csv(index=False))
    
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()