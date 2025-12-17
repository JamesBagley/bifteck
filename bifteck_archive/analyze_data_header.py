import olefile
import zlib
import struct

def analyze_data_header(file_path):
    """Examine the 628-byte header of the DATA stream"""
    print(f"\n{'='*70}")
    print(f"Analyzing DATA header: {file_path}")
    print(f"{'='*70}")
    
    ole = olefile.OleFileIO(file_path)
    stream_path = ['SUBSETS', '2', 'DATA']
    raw = ole.openstream(stream_path).read()
    zlib_off = raw.find(b'\x78\x9c')
    data = zlib.decompress(raw[zlib_off:])
    
    # First 628 bytes is the header
    header = data[:628]
    
    print(f"Header size: {len(header)} bytes")
    
    # Search for interval as double (3600 seconds or 0.0416667 days)
    print(f"\nSearching for 3600 (1 hour in seconds) as double:")
    for i in range(0, len(header) - 8):
        val = struct.unpack('<d', header[i:i+8])[0]
        if 3590 < val < 3610:
            print(f"  Offset {i:3d}: {val:.4f} seconds")
    
    print(f"\nSearching for ~0.0417 (1 hour in days) as double:")
    for i in range(0, len(header) - 8):
        val = struct.unpack('<d', header[i:i+8])[0]
        if 0.041 < val < 0.043:
            secs = val * 24 * 3600
            print(f"  Offset {i:3d}: {val:.10f} days = {secs:.1f} seconds")
    
    # Search as 32-bit integer
    print(f"\nSearching for 3600 as 32-bit integer:")
    for i in range(0, len(header) - 4):
        val = struct.unpack('<I', header[i:i+4])[0]
        if val == 3600:
            print(f"  Offset {i:3d}: {val}")
        elif 3590 < val < 3610:
            print(f"  Offset {i:3d}: {val} (close)")
    
    # Search as 16-bit integer (maybe it's stored in minutes?)
    print(f"\nSearching for 60 (1 hour in minutes) as 16-bit integer:")
    for i in range(0, len(header) - 2):
        val = struct.unpack('<H', header[i:i+2])[0]
        if val == 60:
            print(f"  Offset {i:3d}: {val}")
    
    # Show all non-zero doubles in header
    print(f"\nAll non-zero doubles in header:")
    for i in range(0, len(header) - 8, 8):
        val = struct.unpack('<d', header[i:i+8])[0]
        if val != 0 and abs(val) < 1e10:
            # Try to interpret
            interp = []
            if 40000 < val < 50000:
                interp.append("Excel date")
            if 0 < val < 100:
                interp.append(f"small value")
            if 100 < val < 10000:
                interp.append(f"maybe seconds?")
            
            interp_str = f" ({', '.join(interp)})" if interp else ""
            print(f"  [{i:3d}] {val:20.10f}{interp_str}")
    
    # Show raw hex
    print(f"\nRaw hex (all 628 bytes):")
    for i in range(0, len(header), 16):
        hex_str = ' '.join(f'{b:02x}' for b in header[i:i+16])
        ascii_str = ''.join(chr(b) if 32 <= b < 127 else '.' for b in header[i:i+16])
        print(f"  {i:04x}: {hex_str:<48} {ascii_str}")

analyze_data_header('SlowExperiment.xpt')
print("\n" + "="*70)
analyze_data_header('SweepExperiment.xpt')
