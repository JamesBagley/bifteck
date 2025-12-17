import olefile
import zlib
import struct

def analyze_temp_patterns(file_path, subset_numbers):
    """Analyze temperature patterns across different timepoint counts"""
    ole = olefile.OleFileIO(file_path)
    
    results = []
    for subset_num in subset_numbers:
        stream_path = ['SUBSETS', str(subset_num), 'DATA']
        if not ole.exists(stream_path):
            continue
            
        raw = ole.openstream(stream_path).read()
        zlib_off = raw.find(b'\x78\x9c')
        if zlib_off == -1:
            continue
        data = zlib.decompress(raw[zlib_off:])
        
        HEADER_SIZE = 628
        MATRIX_STRIDE = 9216
        num_timepoints = (len(data) - HEADER_SIZE) // MATRIX_STRIDE
        
        if num_timepoints == 0:
            continue
            
        matrix_end = HEADER_SIZE + (num_timepoints * MATRIX_STRIDE)
        footer = data[matrix_end:]
        
        print(f"\n{file_path} - SUBSET {subset_num} ({num_timepoints} timepoints)")
        print("="*60)
        
        # Search for the specific pattern
        temp_pattern = struct.pack('<IHHHIH', 0x18, 2, num_timepoints & 0xFFFF, 0, 1, 0)
        pattern_search = footer.find(temp_pattern)
        if pattern_search != -1:
            offset = pattern_search + len(temp_pattern)
            print(f"✓ Pattern found at offset {pattern_search}, temp at {offset}")
            # Extract temps to verify
            for i in range(min(3, num_timepoints)):
                temp_off = offset + (i * 24)
                if temp_off + 8 <= len(footer):
                    temp = struct.unpack('<d', footer[temp_off:temp_off+8])[0]
                    print(f"  Temp {i}: {temp:.2f}°C")
            results.append({'file': file_path, 'subset': subset_num, 'timepoints': num_timepoints, 'method': 'pattern', 'offset': offset})
        else:
            print(f"✗ Pattern NOT found")
            # Brute force search
            for offset in range(0, min(600, len(footer) - 8)):
                val = struct.unpack('<d', footer[offset:offset+8])[0]
                if 20 <= val <= 30:
                    print(f"  Brute force found temp at offset {offset}: {val:.2f}°C")
                    # Check if this is the start of an array
                    all_valid = True
                    for i in range(min(3, num_timepoints)):
                        temp_off = offset + (i * 24)
                        if temp_off + 8 <= len(footer):
                            temp = struct.unpack('<d', footer[temp_off:temp_off+8])[0]
                            print(f"    Temp {i}: {temp:.2f}°C (stride 24)")
                            if not (10 <= temp <= 40):
                                all_valid = False
                        else:
                            all_valid = False
                    if all_valid:
                        results.append({'file': file_path, 'subset': subset_num, 'timepoints': num_timepoints, 'method': 'brute_force', 'offset': offset})
                    break
    
    ole.close()
    return results

# Test all files
all_results = []
all_results.extend(analyze_temp_patterns('SlowExperiment.xpt', [2, 3]))
all_results.extend(analyze_temp_patterns('SweepExperiment.xpt', [4, 5, 6, 7]))
all_results.extend(analyze_temp_patterns('384_OD_test.xpt', [1, 2, 3]))
all_results.extend(analyze_temp_patterns('Multiplate test.xpt', [2, 3]))
all_results.extend(analyze_temp_patterns('Multiplate test 2.xpt', [2, 3]))

print("\n" + "="*80)
print("SUMMARY")
print("="*80)
for r in all_results:
    print(f"{r['file']:30s} SUBSET {r['subset']:2d} ({r['timepoints']:2d} TP): {r['method']:15s} @ offset {r['offset']}")
