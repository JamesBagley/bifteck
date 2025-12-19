import olefile
import zlib
import struct
import polars as pl
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
        elif subset_number is None:
            # Extract subset number from stream path
            if len(stream_path) >= 2 and stream_path[0] == 'SUBSETS':
                subset_number = int(stream_path[1])
        
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
        
        # 4a. Find temperature array by searching for the pattern: nan, 0.0, then valid temps
        # Temperature arrays can be in FOOTER (small files) or MATRIX (large files)
        # Universal approach: search for the nan/0.0/temp pattern
        import math
        OFFSET_TEMP_ARRAY = None
        temp_array_absolute_offset = None
        temp_array_in_footer = False
        
        # Search from near end of matrix backward into footer
        search_start = max(HEADER_SIZE, matrix_end_offset - 10000)
        search_end = min(len(data) - 56, matrix_end_offset + 1000)
        
        for offset in range(search_start, search_end):
            if offset + 56 <= len(data):
                val0 = struct.unpack('<d', data[offset:offset+8])[0]
                val1 = struct.unpack('<d', data[offset+24:offset+32])[0]
                val2 = struct.unpack('<d', data[offset+48:offset+56])[0]
                
                # Check for: nan, 0.0, temp pattern
                if math.isnan(val0) and abs(val1) < 0.01 and 15 <= val2 <= 50:
                    temp_array_absolute_offset = offset
                    # Check if in footer or matrix region
                    if offset >= matrix_end_offset:
                        OFFSET_TEMP_ARRAY = offset - matrix_end_offset
                        temp_array_in_footer = True
                    break
        
        if temp_array_absolute_offset is None:
            print(f"  [!] Warning: No temperature data found")
            OFFSET_TEMP_ARRAY = None
        
        # 4b. Find timestamp offset using "OLE" landmark
        # Universal marker: \x00\x03OLE\x00 appears 12 bytes before all timestamp arrays
        # For large files (>~80 timepoints), timestamps are embedded in matrix data
        # Need to find OLE marker that can accommodate all timestamps
        OFFSET_TIMESTAMP_ARRAY = None
        timestamp_array_absolute_offset = None
        ole_marker = b'\x00\x03OLE\x00'
        
        # Search entire decompressed data for OLE markers
        pos = 0
        while True:
            pos = data.find(ole_marker, pos)
            if pos == -1:
                break
            
            ts_offset = pos + 12
            # Check if this OLE marker can fit all timestamps
            remaining_bytes = len(data) - ts_offset
            max_timestamps = remaining_bytes // TIMESTAMP_STRIDE
            
            if max_timestamps >= num_timepoints:
                # Verify first few values are valid Excel dates
                valid_count = 0
                for i in range(min(10, num_timepoints)):
                    offset = ts_offset + (i * TIMESTAMP_STRIDE)
                    if offset + 8 <= len(data):
                        val = struct.unpack('<d', data[offset:offset+8])[0]
                        # Valid Excel date range: ~0 to 80000 (years 1900-2119)
                        if 0 < val < 80000:
                            valid_count += 1
                
                if valid_count >= min(8, num_timepoints):
                    timestamp_array_absolute_offset = ts_offset
                    # Convert to footer offset if in footer region
                    if ts_offset >= matrix_end_offset:
                        OFFSET_TIMESTAMP_ARRAY = ts_offset - matrix_end_offset
                    break
            
            pos += 1
        
        if timestamp_array_absolute_offset is None:
            print(f"  [!] Warning: No timestamp array found")
            OFFSET_TIMESTAMP_ARRAY = None

        # 5. Extract Data Matrix
        matrix_bytes = data[HEADER_SIZE : matrix_end_offset]

        # Load as array of doubles
        total_doubles = len(matrix_bytes) // 8
        matrix_floats = struct.unpack(f'<{total_doubles}d', matrix_bytes)

        # Extract only the measurement values (every 3rd element starting at 0)
        # 1152 doubles per timepoint = 384 wells * 3 values (Value, Flag, Status)
        well_labels = [f"{r}{c}" for r in "ABCDEFGHIJKLMNOP" for c in range(1, 25)]
        
        # First, validate which timepoints contain real data vs garbage
        # Plate reader data should be in reasonable range (typically 0-10 for absorbance/fluorescence)
        # Garbage data will have extreme values or random near-zero floating point noise
        valid_timepoints = []
        for t in range(num_timepoints):
            # Sample first 20 values from this timepoint
            sample_values = [abs(matrix_floats[t * 1152 + i * 3]) for i in range(min(20, 384))]
            max_val = max(sample_values)
            avg_val = sum(sample_values) / len(sample_values)
            
            # Valid data: max < 1e10 (catches garbage like 3.794275e+81)
            # Also check average isn't suspiciously tiny (all near-zero)
            if max_val < 1e10:
                valid_timepoints.append(t)
            else:
                print(f"  [!] Warning: Timepoint {t+1} contains garbage data (max: {max_val:.2e}, avg: {avg_val:.2e}), skipping", file=sys.stderr)
        
        # Update num_timepoints to reflect actual valid data
        actual_num_timepoints = len(valid_timepoints)
        if actual_num_timepoints < num_timepoints:
            print(f"  [+] Valid timepoints: {actual_num_timepoints} (removed {num_timepoints - actual_num_timepoints} garbage blocks)", file=sys.stderr)
            num_timepoints = actual_num_timepoints
        
        measurements = {}
        for well_idx, well in enumerate(well_labels):
            measurements[well] = [matrix_floats[t * 1152 + well_idx * 3] for t in valid_timepoints]
        
        df = pl.DataFrame(measurements)

        # 6. Extract Temperatures
        # Array structure: [NaN, 0.0, temp0, temp1, temp2, ...]
        # FOOTER location: Has temps for all timepoints
        # MATRIX location: May be missing first timepoint temp (data recording issue)
        temps = []
        
        if temp_array_absolute_offset is not None:
            # Skip first 2 values (nan, 0.0) and extract actual temperatures
            # Start at index 2 (first real temperature)
            for i in range(num_timepoints):
                # Index 2 = timepoint 0, Index 3 = timepoint 1, etc.
                array_index = i + 2
                offset = temp_array_absolute_offset + (array_index * TEMP_STRIDE)
                
                if offset + 8 > len(data):
                    print(f"  [!] Warning: Temperature data truncated at timepoint {i}", file=sys.stderr)
                    temps.extend([None] * (num_timepoints - i))
                    break

                val = struct.unpack('<d', data[offset:offset+8])[0]
                
                # Check if this is a valid temperature
                if 10 <= val <= 40:
                    temps.append(val)
                else:
                    # Not a valid temp - could be end of array or missing data
                    if temp_array_in_footer:
                        # Footer location should have all temps, this is unexpected
                        print(f"  [!] Warning: Invalid temperature {val:.1f}°C at timepoint {i}", file=sys.stderr)
                        temps.append(None)
                    else:
                        # Matrix location may have incomplete data (known issue with large files)
                        temps.append(None)
        else:
            # No temperature data found
            temps = [None] * num_timepoints

        df = df.with_columns(pl.Series("Temperature", temps))
        print(f"  [+] Temperatures: {len([t for t in temps if t is not None])} readings", file=sys.stderr)

        # 7. Extract Timestamps
        # Timestamps are stored as Excel date format (days since 1899-12-30)
        # For large files (>~80 timepoints), timestamps are embedded in matrix data
        base_date = datetime(1899, 12, 30)
        
        timestamps = []
        
        if timestamp_array_absolute_offset is not None:
            for i in range(num_timepoints):
                ts_offset = timestamp_array_absolute_offset + (i * TIMESTAMP_STRIDE)
                
                if ts_offset + 8 > len(data):
                    print(f"  [!] Warning: Timestamp {i} out of bounds", file=sys.stderr)
                    timestamps.append(None)
                    continue
                
                ts_float = struct.unpack('<d', data[ts_offset:ts_offset+8])[0]
                
                # Validate it's a reasonable Excel date (between year 1900-2100)
                if not (1 < ts_float < 80000):
                    print(f"  [!] Warning: Invalid timestamp at index {i} (value: {ts_float})", file=sys.stderr)
                    timestamps.append(None)
                    continue
                
                timestamp = base_date + timedelta(days=ts_float)
                timestamps.append(timestamp)
        else:
            # No timestamp array found
            print(f"  [!] Warning: No timestamp array found", file=sys.stderr)
            timestamps = [None] * num_timepoints
        
        # Generate timestamp strings and elapsed time labels
        valid_timestamps = [ts for ts in timestamps if ts is not None]
        if valid_timestamps:
            print(f"  [+] Start time: {valid_timestamps[0]}", file=sys.stderr)
            print(f"  [+] End time: {valid_timestamps[-1]}", file=sys.stderr)
            
            # Calculate and show interval statistics
            if len(valid_timestamps) >= 2:
                # Calculate intervals only between consecutive valid timestamps
                intervals = []
                for i in range(len(timestamps)-1):
                    if timestamps[i] is not None and timestamps[i+1] is not None:
                        intervals.append((timestamps[i+1] - timestamps[i]).total_seconds())
                if intervals:
                    avg_interval = sum(intervals) / len(intervals)
                    print(f"  [+] Average interval: {avg_interval:.1f}s ({avg_interval/3600:.3f}h)", file=sys.stderr)
        
        # Generate elapsed time labels and ISO strings
        time_abs_strs = []
        elapsed_labels = []
        first_valid_ts = next((ts for ts in timestamps if ts is not None), None)
        
        for i, ts in enumerate(timestamps):
            if ts is None:
                time_abs_strs.append(None)
                elapsed_labels.append(None)
            else:
                time_abs_strs.append(ts.isoformat())
                if first_valid_ts is not None:
                    total_seconds = int((ts - first_valid_ts).total_seconds())
                    hours = total_seconds // 3600
                    minutes = (total_seconds % 3600) // 60
                    seconds = total_seconds % 60
                    elapsed_labels.append(f"{hours}:{minutes:02d}:{seconds:02d}")
                else:
                    elapsed_labels.append(None)

        df = df.with_columns([
            pl.Series("Read", list(range(1, num_timepoints + 1))),
            pl.Series("Time_Abs", time_abs_strs),
            pl.Series("Time", elapsed_labels),
            pl.Series("PlateID", [plate_id if plate_id else ""] * num_timepoints),
            pl.Series("Barcode", [barcode if barcode else ""] * num_timepoints)
        ])

        # 8. Clean Data
        # Remove well columns that are all zeros (empty wells)
        cols_to_keep = ["Read", "PlateID", "Barcode", "Time", "Time_Abs", "Temperature"]
        for col in well_labels:
            if df[col].sum() != 0:
                cols_to_keep.append(col)
        df = df.select(cols_to_keep)
        
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
                df = xpt.extract_stream(subset_number=int(subset_number)).with_columns([
                    pl.lit(file_path).alias("File"),
                    pl.lit(subset_number).alias("Subset")
                ])
                frames.append(df)
            except Exception as e:
                print(f"  [!] Warning: Error processing {file_path} subset {subset_number}, skipping: {e}", file=sys.stderr)
                continue
        
        if not frames:
            raise ValueError(f"No valid data extracted from {file_path}")
        
        return pl.concat(frames).unpivot(
            index=["File", "Subset", "PlateID", "Barcode", "Time", "Time_Abs", "Temperature", "Read"],
            variable_name="Well",
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
            df.write_csv(args.output)
            print(f"\nData exported to {args.output}", file=sys.stderr)
        else:
            # Print to stdout (diagnostic messages go to stderr)
            print(df.write_csv())
    
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()