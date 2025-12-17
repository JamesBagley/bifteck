import olefile
import struct

def analyze_contents_stream(file_path):
    """Examine the Contents stream in detail"""
    print(f"\n{'='*70}")
    print(f"Analyzing Contents stream: {file_path}")
    print(f"{'='*70}")
    
    ole = olefile.OleFileIO(file_path)
    
    # Read Contents stream
    raw = ole.openstream('Contents').read()
    print(f"Size: {len(raw)} bytes")
    
    # Try to find strings
    print(f"\nSearching for text strings (min length 4):")
    text = ''
    offset = 0
    for i, b in enumerate(raw):
        if 32 <= b < 127:
            if not text:
                offset = i
            text += chr(b)
        elif text:
            if len(text) >= 4:
                print(f"  [{offset:5d}] '{text}'")
            text = ''
    
    # Look for numbers that might be intervals
    print(f"\nSearching for potential interval values as doubles:")
    for i in range(0, len(raw) - 8, 1):  # Search byte by byte
        try:
            val = struct.unpack('<d', raw[i:i+8])[0]
            
            # 1 hour = 3600 seconds
            if 3590 < val < 3610:
                print(f"  Offset {i:5d}: {val:.2f} (close to 3600 = 1 hour)")
            
            # 1 hour as fractional day
            elif 0.041 < val < 0.043:
                secs = val * 24 * 3600
                print(f"  Offset {i:5d}: {val:.10f} days ({secs:.1f} seconds)")
        except:
            pass
    
    # Look for 32-bit integers that might be seconds
    print(f"\nSearching for potential interval values as 32-bit integers:")
    for i in range(0, len(raw) - 4, 1):
        try:
            val = struct.unpack('<I', raw[i:i+4])[0]
            if 3590 < val < 3610:
                print(f"  Offset {i:5d}: {val} (close to 3600 = 1 hour)")
        except:
            pass
    
    # Show first part as hex
    print(f"\nFirst 256 bytes as hex:")
    for i in range(0, min(256, len(raw)), 16):
        hex_str = ' '.join(f'{b:02x}' for b in raw[i:i+16])
        ascii_str = ''.join(chr(b) if 32 <= b < 127 else '.' for b in raw[i:i+16])
        print(f"  {i:04x}: {hex_str:<48} {ascii_str}")

analyze_contents_stream('SlowExperiment.xpt')
analyze_contents_stream('SweepExperiment.xpt')
