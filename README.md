#Bifteck reads Biotek Gen5 xpr files

Use in python:

!uv pip install https://github.com/JamesBagley/bifteck.git

from bifteck.bifteck import read_xpt_file

my_plates = read_xpt_file("my_multiplate_experiment.xpt") 
my_plates.head()


Use in R:

library(readr)

# Direct pipe (prints progress to stderr)
df <- read_csv(pipe("uv run bifteck.py experiment.xpt"))

# Or save to file first
system("uv run bifteck.py experiment.xpt -o output.csv")
df <- read_csv("output.csv")


Intended use is for discontinuous kinetic experiments, I don't know exactly what will happen with other file types.

Happy coding!