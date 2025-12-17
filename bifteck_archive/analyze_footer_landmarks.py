import olefile
import zlib
import struct
from datetime import datetime, timedelta

def analyze_footer_structure(file_path, subset_num):
    """Analyze footer byte structure to find landmarks"""
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
    
    # Find where temperatures actually are
    temp_offset = None
    base_date = datetime(1899, 12, 30)
    for offset in range(0, min(600, len(footer) - 8)):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        if 20 <= val <= 30:
            temp_offset = offset
            break
    
    # Find where timestamps actually are
    ts_offset = None
    for offset in range(0, min(1200, len(footer) - 8)):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        if 46000 < val < 46010:
            # Verify it's the start of an array by checking next value
            next_offset = offset + 86
            if next_offset + 8 <= len(footer):
                next_val = struct.unpack('<d', footer[next_offset:next_offset+8])[0]
                if 46000 < next_val < 46010:
                    ts_offset = offset
                    break
    
    return {
        'file': file_path,
        'subset': subset_num,
        'timepoints': num_timepoints,
        'footer_size': len(footer),
        'temp_offset': temp_offset,
        'ts_offset': ts_offset,
        'footer_bytes': footer
    }

def find_byte_landmarks(footer, target_offset, context_size=100):
    """Look for byte patterns near a target offset"""
    start = max(0, target_offset - context_size)
    end = min(len(footer), target_offset + context_size)
    region = footer[start:end]
    
    # Look for distinctive byte patterns
    patterns = []
    
    # Look for repeated bytes
    for i in range(len(region) - 4):
        pattern = region[i:i+4]
        if pattern == b'\xff\xff\xff\xff':
            patterns.append(('0xFFFFFFFF', start + i))
        elif pattern == b'\x00\x00\x00\x00':
            patterns.append(('0x00000000', start + i))
        elif pattern[0] == pattern[1] == pattern[2] == pattern[3] and pattern[0] != 0:
            patterns.append((f'Repeated 0x{pattern[0]:02x}', start + i))
    
    # Look for specific values as integers
    for i in range(0, len(region) - 4, 4):
        val = struct.unpack('<I', region[i:i+4])[0]
        if val == target_offset:
            patterns.append((f'Offset value {target_offset}', start + i))
    
    return patterns

# Analyze all subsets with data
print("="*80)
print("FOOTER STRUCTURE ANALYSIS - Looking for Landmarks")
print("="*80)

results = []
for file in ['SlowExperiment.xpt', 'SweepExperiment.xpt']:
    for subset in range(2, 8):
        result = analyze_footer_structure(file, subset)
        if result and result['temp_offset'] is not None:
            results.append(result)

# Group by timepoint count
by_timepoints = {}
for r in results:
    tp = r['timepoints']
    if tp not in by_timepoints:
        by_timepoints[tp] = []
    by_timepoints[tp].append(r)

# Analyze each group
for tp_count in sorted(by_timepoints.keys()):
    group = by_timepoints[tp_count]
    print(f"\n{'='*80}")
    print(f"TIMEPOINTS: {tp_count} (n={len(group)} samples)")
    print(f"{'='*80}")
    
    # Get consistent offsets
    temp_offsets = [r['temp_offset'] for r in group]
    ts_offsets = [r['ts_offset'] for r in group]
    
    print(f"\nConsistent offsets:")
    print(f"  Temperature: {set(temp_offsets)}")
    print(f"  Timestamp:   {set(ts_offsets)}")
    
    # Analyze footer bytes around temperature offset
    sample = group[0]
    footer = sample['footer_bytes']
    temp_off = sample['temp_offset']
    ts_off = sample['ts_offset']
    
    print(f"\n--- Bytes around Temperature offset ({temp_off}) ---")
    start = max(0, temp_off - 60)
    end = min(len(footer), temp_off + 20)
    
    for i in range(start, end, 16):
        hex_str = ' '.join(f'{b:02x}' for b in footer[i:i+16])
        marker = " <-- TEMP START" if i <= temp_off < i+16 else ""
        print(f"  {i:04d}: {hex_str}{marker}")
    
    print(f"\n--- Bytes around Timestamp offset ({ts_off}) ---")
    if ts_off:
        start = max(0, ts_off - 60)
        end = min(len(footer), ts_off + 20)
        
        for i in range(start, end, 16):
            hex_str = ' '.join(f'{b:02x}' for b in footer[i:i+16])
            marker = " <-- TS START" if i <= ts_off < i+16 else ""
            print(f"  {i:04d}: {hex_str}{marker}")
    
    # Look for patterns that precede the data
    print(f"\n--- Looking for distinctive patterns before temperature ---")
    if temp_off >= 20:
        preceding = footer[temp_off-20:temp_off]
        print(f"  20 bytes before temp: {preceding.hex()}")
        
        # Check if there's a count or size marker
        for offset in [temp_off-4, temp_off-8, temp_off-12]:
            if offset >= 0:
                as_int32 = struct.unpack('<I', footer[offset:offset+4])[0]
                as_int16 = struct.unpack('<H', footer[offset:offset+2])[0]
                print(f"  At offset {offset} (Δ{offset-temp_off}): int32={as_int32}, int16={as_int16}")
    
    print(f"\n--- Looking for distinctive patterns before timestamp ---")
    if ts_off and ts_off >= 20:
        preceding = footer[ts_off-20:ts_off]
        print(f"  20 bytes before ts: {preceding.hex()}")
        
        # Check if there's a count or size marker
        for offset in [ts_off-4, ts_off-8, ts_off-12]:
            if offset >= 0:
                as_int32 = struct.unpack('<I', footer[offset:offset+4])[0]
                as_int16 = struct.unpack('<H', footer[offset:offset+2])[0]
                print(f"  At offset {offset} (Δ{offset-ts_off}): int32={as_int32}, int16={as_int16}")

# Look for relationships between offsets
print(f"\n{'='*80}")
print("ANALYZING OFFSET RELATIONSHIPS")
print(f"{'='*80}\n")

print("Temperature to Timestamp distance:")
for tp_count in sorted(by_timepoints.keys()):
    group = by_timepoints[tp_count]
    sample = group[0]
    distance = sample['ts_offset'] - sample['temp_offset'] if sample['ts_offset'] else None
    print(f"  {tp_count} timepoints: {distance} bytes")

print("\nFooter size vs timepoints:")
for tp_count in sorted(by_timepoints.keys()):
    group = by_timepoints[tp_count]
    sample = group[0]
    print(f"  {tp_count} timepoints: {sample['footer_size']} bytes")
