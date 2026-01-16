# Bifteck reads Biotek Gen5 xpt files

Use in python:
```
!uv pip install https://github.com/JamesBagley/bifteck.git

from bifteck import read_xpt_file

my_plates = read_xpt_file("my_multiplate_experiment.xpt")

my_plates.head()
```

Use in R:
```
library(readr)

# Direct pipe (prints progress to stderr)
df <- read_csv(pipe("uv run bifteck experiment.xpt"))

# Or save to file first
system("uv run bifteck experiment.xpt -o output.csv")

df <- read_csv("output.csv")
```

# What's needed in the working directory for R users:

Just the pyproject.toml file (which contains the dependencies and package configuration)

The bifteck directory with the code

uv will handle installing dependencies automatically when running the command

# Disclaimer
Intended use is for discontinuous kinetic experiments, I don't know exactly what will happen with other file types.

# How it Works
File Format
BioTek XPT files are OLE (Object Linking and Embedding) files, also known as Compound File Binary Format. They are analagous to zip files but have a different terminology, they contain "streams" (analogous to files) which are organized in a directory structure.

Data Organization
Inside an XPT file, data is organized as:

Each SUBSET represents one plate in a multi-plate experiment. The DATA stream contains:

A 628-byte header
Matrix blocks (9216 bytes each) containing all 384 well measurements for each timepoint
A footer with temperature and timestamp arrays
The HEADER stream contains plate identification (Plate ID and Barcode).

Implementation
This library uses:

olefile to navigate the OLE structure and extract streams
zlib to decompress the DATA streams
struct to parse binary data (doubles, timestamps)
polars for fast DataFrame operations and CSV export
The extracted data is returned in long format (one row per well per timepoint) for easy analysis.

Happy coding!