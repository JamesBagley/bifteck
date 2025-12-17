import olefile
import struct

def analyze_header_streams(file_path):
    """Examine all HEADER streams in detail"""
    print(f"\n{'='*70}")
    print(f"Analyzing: {file_path}")
    print(f"{'='*70}")
    
    ole = olefile.OleFileIO(file_path)
    
    # Find all HEADER streams
    header_streams = [s for s in ole.listdir() if len(s) == 3 and s[2] == 'HEADER']
    
    for stream_path in header_streams:
        print(f"\n--- Stream: {'/'.join(stream_path)} ---")
        
        raw = ole.openstream(stream_path).read()
        print(f"Size: {len(raw)} bytes")
        
        # Try to parse as doubles
        if len(raw) >= 8:
            num_doubles = len(raw) // 8
            print(f"\nAs doubles ({num_doubles} values):")
            for i in range(num_doubles):
                offset = i * 8
                val = struct.unpack('<d', raw[offset:offset+8])[0]
                
                # Show non-zero and interesting values
                if val != 0 or i < 3:  # Always show first 3
                    interpretations = []
                    
                    # Check if it could be seconds
                    if 0 < val < 86400:
                        interpretations.append(f"seconds: {val:.1f} ({val/3600:.2f}h)")
                    
                    # Check if it could be fractional days
                    if 0 < val < 1:
                        secs = val * 24 * 3600
                        interpretations.append(f"days: {val:.10f} ({secs:.1f}s)")
                    
                    # Check if Excel date
                    if 40000 < val < 50000:
                        interpretations.append(f"Excel date")
                    
                    if interpretations:
                        print(f"  [{i:2d}] offset {offset:3d}: {val:25.15f} | {' | '.join(interpretations)}")
                    elif abs(val) < 1000:
                        print(f"  [{i:2d}] offset {offset:3d}: {val:25.15f}")
        
        # Also show as raw hex for first 128 bytes
        print(f"\nRaw hex (first 128 bytes):")
        for i in range(0, min(128, len(raw)), 16):
            hex_str = ' '.join(f'{b:02x}' for b in raw[i:i+16])
            ascii_str = ''.join(chr(b) if 32 <= b < 127 else '.' for b in raw[i:i+16])
            print(f"  {i:04x}: {hex_str:<48} {ascii_str}")
        
        # Try to find strings
        print(f"\nSearching for text strings:")
        text = ''
        for b in raw:
            if 32 <= b < 127:
                text += chr(b)
            elif text:
                if len(text) > 3:
                    print(f"  '{text}'")
                text = ''

analyze_header_streams('SlowExperiment.xpt')
analyze_header_streams('SweepExperiment.xpt')
