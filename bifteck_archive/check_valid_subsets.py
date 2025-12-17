from main import extract_biotek_final

def check_which_subsets_have_data(file_path):
    """Check all subsets to see which have valid data"""
    import olefile
    import zlib
    import struct
    
    print(f"\n{'='*70}")
    print(f"Checking: {file_path}")
    print(f"{'='*70}\n")
    
    ole = olefile.OleFileIO(file_path)
    
    for subset in range(2, 8):
        stream_path = ['SUBSETS', str(subset), 'DATA']
        if not ole.exists(stream_path):
            print(f"SUBSET {subset}: Stream does not exist")
            continue
        
        try:
            raw = ole.openstream(stream_path).read()
            zlib_off = raw.find(b'\x78\x9c')
            data = zlib.decompress(raw[zlib_off:])
            
            HEADER_SIZE = 628
            MATRIX_STRIDE = 9216
            num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
            matrix_end_offset = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
            footer = data[matrix_end_offset:]
            
            # Check first temperature
            OFFSET_TEMP = 364
            temp = struct.unpack('<d', footer[OFFSET_TEMP:OFFSET_TEMP+8])[0]
            
            # Check first timestamp
            OFFSET_TS = 1012
            ts_val = struct.unpack('<d', footer[OFFSET_TS:OFFSET_TS+8])[0]
            
            has_valid_data = (10 <= temp <= 40) and (1 < ts_val < 100000)
            
            status = "✓ HAS DATA" if has_valid_data else "✗ EMPTY/INVALID"
            print(f"SUBSET {subset}: {status} - {num_timepoints} timepoints, temp={temp:.1f}°C, ts={ts_val:.2e}")
            
        except Exception as e:
            print(f"SUBSET {subset}: ERROR - {e}")

check_which_subsets_have_data('SlowExperiment.xpt')
check_which_subsets_have_data('SweepExperiment.xpt')
