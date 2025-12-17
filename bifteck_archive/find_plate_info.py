import olefile

ole = olefile.OleFileIO('DinosaurusRex.xpt')

# List all streams
print("All streams in file:")
for stream in ole.listdir():
    print(f"  {'/'.join(stream)}")

print("\n" + "="*60)
print("Searching for 'YeeHaw' and 'Dinosaurus Rex' in all streams...")
print("="*60)

for stream_path in ole.listdir():
    try:
        raw = ole.openstream(stream_path).read()
        path_str = '/'.join(stream_path)
        
        # Search for the strings (both as ASCII and UTF-16)
        yeehaw_ascii = b'YeeHaw'
        yeehaw_utf16 = 'YeeHaw'.encode('utf-16-le')
        dino_ascii = b'Dinosaurus Rex'
        dino_utf16 = 'Dinosaurus Rex'.encode('utf-16-le')
        
        found = False
        
        if yeehaw_ascii in raw:
            pos = raw.find(yeehaw_ascii)
            print(f"\n{path_str}: Found 'YeeHaw' (ASCII) at offset {pos}")
            print(f"  Context: {raw[max(0,pos-20):pos+30]}")
            found = True
            
        if yeehaw_utf16 in raw:
            pos = raw.find(yeehaw_utf16)
            print(f"\n{path_str}: Found 'YeeHaw' (UTF-16) at offset {pos}")
            print(f"  Context bytes: {raw[max(0,pos-20):pos+40].hex(' ')}")
            found = True
            
        if dino_ascii in raw:
            pos = raw.find(dino_ascii)
            print(f"\n{path_str}: Found 'Dinosaurus Rex' (ASCII) at offset {pos}")
            print(f"  Context: {raw[max(0,pos-20):pos+40]}")
            found = True
            
        if dino_utf16 in raw:
            pos = raw.find(dino_utf16)
            print(f"\n{path_str}: Found 'Dinosaurus Rex' (UTF-16) at offset {pos}")
            print(f"  Context bytes: {raw[max(0,pos-20):pos+40].hex(' ')}")
            found = True
            
    except Exception as e:
        print(f"{path_str}: Error - {e}")

ole.close()
