import olefile
from main import extract_biotek_final

def test_file(filename, subset_num=None):
    """Test a file and report what happens"""
    print(f"\n{'='*80}")
    print(f"Testing: {filename}")
    print(f"{'='*80}")
    
    # First, check what streams exist
    try:
        ole = olefile.OleFileIO(filename)
        print("\nAvailable streams:")
        for entry in ole.listdir():
            if entry[0] == 'SUBSETS' and len(entry) >= 3:
                stream_path = ['SUBSETS', entry[1], entry[2]]
                stream_size = ole.get_size(stream_path)
                print(f"  {'/'.join(entry)}: {stream_size} bytes")
        ole.close()
    except Exception as e:
        print(f"Error listing streams: {e}")
        return
    
    # Try to extract with default settings
    if subset_num is None:
        print("\nTrying extraction with auto-detect (subset_number=2)...")
        try:
            df = extract_biotek_final(filename, subset_number=2)
            print(f"✅ SUCCESS! Shape: {df.shape}")
            print(f"\nFirst few rows:")
            print(df[['Time', 'Time_Abs', 'Temperature']].head(10))
        except Exception as e:
            print(f"❌ FAILED: {e}")
    else:
        print(f"\nTrying extraction with subset_number={subset_num}...")
        try:
            df = extract_biotek_final(filename, subset_number=subset_num)
            print(f"✅ SUCCESS! Shape: {df.shape}")
            print(f"\nFirst few rows:")
            print(df[['Time', 'Time_Abs', 'Temperature']].head(10))
        except Exception as e:
            print(f"❌ FAILED: {e}")

# Test each new file
print("TESTING NEW XPT FILES")
print("="*80)

# Test 1: Empty file
test_file('Experiment1.xpt')

# Test 2: Single-point reads
test_file('384_OD_test.xpt')

# Test 3: Multiplate with empty first stream
test_file('Multiplate test.xpt')

# Test 4: Multiplate with deleted first stream
test_file('Multiplate test 2.xpt')

# Try other subsets for multiplate files
print(f"\n\n{'='*80}")
print("TESTING OTHER SUBSETS FOR MULTIPLATE FILES")
print("="*80)

for subset in [3, 4, 5, 6, 7]:
    test_file('Multiplate test.xpt', subset_num=subset)
