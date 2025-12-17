from main import extract_biotek_final

# Test the improved extraction function
print("="*60)
print("TESTING IMPROVED EXTRACTION")
print("="*60)

# Test on SlowExperiment
df1 = extract_biotek_final('SlowExperiment.xpt')
print("\n--- SlowExperiment Results ---")
print(f"Shape: {df1.shape}")
print(f"\nFirst few rows:")
print(df1.head())
print(f"\nTemperature range: {df1['Temperature'].min():.1f}°C to {df1['Temperature'].max():.1f}°C")
print(f"Time range: {df1['Time'].iloc[0]} to {df1['Time'].iloc[-1]}")

# Test on SweepExperiment
print("\n" + "="*60)
df2 = extract_biotek_final('SweepExperiment.xpt')
print("\n--- SweepExperiment Results ---")
print(f"Shape: {df2.shape}")
print(f"\nFirst few rows:")
print(df2.head())
print(f"\nTemperature range: {df2['Temperature'].min():.1f}°C to {df2['Temperature'].max():.1f}°C")
print(f"Time range: {df2['Time'].iloc[0]} to {df2['Time'].iloc[-1]}")

# Save to CSV for verification
df1.to_csv('SlowExperiment_extracted.csv', index=False)
df2.to_csv('SweepExperiment_extracted.csv', index=False)
print("\n" + "="*60)
print("✓ Files saved:")
print("  - SlowExperiment_extracted.csv")
print("  - SweepExperiment_extracted.csv")
