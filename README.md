(c) Dr Ryan Calmus, University of Iowa, 2023

## Author: Ryan Calmus, PhD
## Name: rawDataToExcel

# Description:
Takes a struct, rawData, organized hierarchically in the form of a manuscript, and generates a set of
Excel files encompassing the data from every panel of every figure of every section of the paper. Each
section will result in a separate Excel file, and each figure will result in its own Excel worksheet
(tab). Each figure may contain multiple tables, e.g. for multiple panels or to describe multiple
features of a single plot. Each table will be written to a separate table within the given worksheet,
titled according to the name of the field. Individual row and column headings will be drawn from the
MATLAB table containing the data to export. Additional labels spanning the entire set of columns and
rows can optionally be specified. rawData should be organized as follows (where square brackets denote
optional fields):
   rawData.paperSection.figureName.tables(1:n).table = <MATLAB table of data>
   [rawData.paperSection.figureName.tables(1:n).colLabelHorizontal = <char label for columns of table>]
   [rawData.paperSection.figureName.tables(1:n).rowLabelVertical = <char label for rows of table>]

# Installation:
Copy the entire file hierarchy to the installation path of choice and add it to your MATLAB path. Call:

> rawDataToExcel(dataStruct)

to generate the shareable Excel files.
