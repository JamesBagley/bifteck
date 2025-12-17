# ARCHIVED CODE - Alternative approaches that were tested but not used in final implementation

"""
This file contains code snippets that were explored during development but not included
in the final main.py implementation. Kept for reference and future debugging.
"""

# =============================================================================
# TIMESTAMP DETECTION - Brute Force Search (Replaced by OLE Landmark)
# =============================================================================
# This approach searches for timestamp values by looking for Excel dates in a 
# reasonable range. Replaced by the OLE marker search which is more reliable.

def find_timestamp_bruteforce(footer, num_timepoints):
    """
    Brute force search for first timestamp (Excel date ~46000 for 2025)
    
    NOTE: This is replaced by OLE landmark search in production code.
    The OLE marker (\x00\x03OLE\x00) appears before all timestamp arrays
    and is more reliable than searching for date values.
    """
    OFFSET_TIMESTAMP_ARRAY = None
    
    for offset in range(0, min(1200, len(footer) - 8)):
        val = struct.unpack('<d', footer[offset:offset+8])[0]
        if 40000 < val < 50000:  # Reasonable date range (2009-2037)
            # Verify it's the start of an array by checking next value
            next_offset = offset + 86
            if next_offset + 8 <= len(footer):
                next_val = struct.unpack('<d', footer[next_offset:next_offset+8])[0]
                if 40000 < next_val < 50000:
                    OFFSET_TIMESTAMP_ARRAY = offset
                    break
    
    return OFFSET_TIMESTAMP_ARRAY


# =============================================================================
# TEMPERATURE DETECTION - Pattern Matching (Not Universal)
# =============================================================================
# This approach tried to find a specific byte pattern before temperature data.
# It only works for some files with 10-11 timepoints and often gives wrong offsets.

def find_temperature_pattern(footer, num_timepoints):
    """
    Search for temperature using specific pattern: 18 00 00 00 02 00 [TP_COUNT] 00 00 00 01 00 00 00
    
    NOTE: This pattern is NOT universal:
    - Only found in 2 out of 13 valid test files
    - When found, often points to wrong offset (reads 0.00°C)
    - Does not work for files with 1, 2, 4 timepoints
    
    Replaced by brute force 20-30°C search which works for all valid files.
    """
    import struct
    
    OFFSET_TEMP_ARRAY = None
    
    if num_timepoints in [10, 11]:
        temp_pattern = struct.pack('<IHHHIH', 0x18, 2, num_timepoints & 0xFFFF, 0, 1, 0)
        temp_search = footer.find(temp_pattern)
        if temp_search != -1:
            OFFSET_TEMP_ARRAY = temp_search + len(temp_pattern)
    
    return OFFSET_TEMP_ARRAY


# =============================================================================
# HARDCODED OFFSETS - Original Approach (Replaced by Landmarks)
# =============================================================================
# The original code used hardcoded offsets based on timepoint count.
# This was replaced by landmark-based detection for better flexibility.

def get_hardcoded_offsets(num_timepoints):
    """
    Original hardcoded offset approach
    
    NOTE: These offsets worked for specific test files but are not universal.
    Landmark-based detection is more robust across different file types.
    """
    if num_timepoints == 11:
        OFFSET_TEMP_ARRAY = 364
        OFFSET_TIMESTAMP_ARRAY = 974
    elif num_timepoints == 10:
        OFFSET_TEMP_ARRAY = 348
        OFFSET_TIMESTAMP_ARRAY = 918
    else:
        # Fallback
        OFFSET_TEMP_ARRAY = 364
        OFFSET_TIMESTAMP_ARRAY = 1012
    
    return OFFSET_TEMP_ARRAY, OFFSET_TIMESTAMP_ARRAY


# =============================================================================
# NOTES ON FILE STRUCTURE VARIATIONS
# =============================================================================
"""
Through testing multiple XPT files, we discovered:

1. TEMPERATURE DATA:
   - Always uses stride of 24 bytes between consecutive readings
   - First temperature is always a double (8 bytes) at some offset
   - Offset varies by file type and timepoint count
   - Reliable method: Search for first value in 20-30°C range
   - Unreliable method: Pattern matching (only works for some files)

2. TIMESTAMP DATA:
   - Always uses stride of 86 bytes between consecutive timestamps
   - Always preceded by OLE marker: \x00\x03OLE\x00
   - Timestamp starts 12 bytes after the OLE marker
   - Excel date format (days since 1899-12-30)
   - This landmark approach is universal across all tested files

3. FOOTER STRUCTURE:
   - Size varies based on number of timepoints and metadata
   - Contains: temperatures, timestamps, audit trail, metadata
   - No fixed structure, must use landmarks or search

4. EDGE CASES:
   - Empty files (0 timepoints): Should raise error
   - Single timepoint files: May not have timestamp arrays
   - Corrupted files: May have invalid or missing data
"""
