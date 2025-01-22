# NEMWeb

## Easy access to Australia's electricity generation data from nemweb.com.au

## Purpose
A Python program providing easy access to the large amount of archived
data of Australia's main electricity grid provided by the market
operator, AEMO.

## Usage
There are three key files
* NEMWeb.py
* DUID Categories.csv
* requirements.txt

Follow the usual Python procedures to use pip to install all the
dependencies in requirements.txt. The first time NEMWeb.py is run (on
Windows) it will read in the most recent 12 calendar months of
electricity generation data from nemweb.com.au, saving binary PKL files
to %APPDATA%\\NEMWeb then saving the following two files to
%USERPROFILE%\\Documents;
1. NEM Year by Category.xlsx
    * A row for each DUID (a power source/sink on the grid) and each
	rooftop solar region: NSW1, QLD1, SA1, TAS1 and VIC1. Columns are
	In, Out and Out/In. It's 50-60 KB in size.
2. NEM In Out List.xlsx
    * a row for each 5-minute interval with two sets of category
	  columns; a negative set for power taken from the grid and a
	  positive set for power put onto the grid. It's over 11MB in size.

Progress is displayed and three interactive graphs are shown; the raw
Dispatch_SCADA graph, a repair Dispatch_SCADA graph and the final
generation graph with rooftop solar included. The graphs show overall GW
generation in 5-minute intervals. NEMWeb.py runs much faster once the
binary PKL files exist.

## Introduction (From Google AI)
The National Electricity Market (NEM) is Australia's wholesale
electricity market and physical power system, which generates, buys,
sells, and transports electricity. The NEM is one of the world's largest
interconnected power systems, supplying around 80% of Australia's
electricity consumption.

The NEM covers five regions on the east coast of Australia: Queensland,
New South Wales, Victoria, Tasmania, and South Australia.

The Australian Energy Market Operator (AEMO) is responsible for
scheduling the lowest priced generation available to meet demand, and
for restoring energy systems to a secure operating state in the event of
an emergency.

## Background
AEMO provides the last 13 months of NEM generation in the Dispatch_SCADA
and ROOFTOP_PV/ACTUAL subdirectories of
www.nemweb.com.au/REPORTS/ARCHIVE.
* SCADA - Supervisory Control And Data Acquisition

### Dispatch_SCADA
Dispatch_SCADA data consists of over 20 million records per year. Each
active generator or storage device sends a record to AEMO every 5
minutes of its power output or input to the electricity grid. Power
values are in MW with input values being negative. Each generator or
storage device is identified using a short Dispatchable Unit Identifier
(DUID). For exanple, ER02 is Eraring unit 2, a 720 MW generator. AEMO
packages Dispatch_SCADA data into daily ZIP files.

One good source of information about the NEM;'s generators and storage
devices is the spreadsheet, nem-registration-and-exemption-list.xlsx,
at aemo.com.au/-/media/files/electricity/nem/participant_information.

### ROOFTOP_PV
Rooftop_PV data (PV - PhotoVoltaic) is the rooftop solar generation of
the five Australian states provided by home or business owners to their
local electricity distribution network. It does not include rooftop
generation used by the home or business owner. The data is provided in
several forms as half hourly records of MW power output and packaged by
AEMO into weekly ZIP files. NEMWeb uses the total power output for each
state, identified with region identifiers as NSW1, QLD1, SA1, TAS1 and
VIC1. 365 day's worth of this data contains 87,600 records.

## Execution
* Load the most recent 12 months of
  * 'Dispatch_SCADA' 5-minute interval electricity generator/storage
	data into a set of daily dataframe and
  * 'ROOFTOP_PV/ACTUAL' 30-minute interval rooftop solar data into a set
	of weekly dataframes.
* Assemble a single 12 month dataframe, each row a 5-minute interval,
  and load it with the generator/storage Dispatch_SCADA data. Five empty
  columns are included for loading rooftop solar data.
  * Note: This dataframe typically has around 50 million cells (100 MB).
* Display the first graph of the raw Dispatch_SCADA data.
* Find and interpolate any missing generator/storage data (a rare
  occurrence).
* Display the second graph of repaired Dispatch_SCADA data.
* Place rooftop solar data in the five empty columns by interpolating
  the 30-minute interval data.
* Display the third graph of the raw Dispatch_SCADA data.
* Generate a dataframe of the average power in to and out of each DUID
  and region ID:
  * Add an Out/In column.
  * Save the dataframe as an Excel spreadsheet.
* Generate a dataframe totalling each category's input and output power:
  * Read in a csv data file where the first column is the DUID or
    region ID and the second column is its category.
  * Build a dataframe totalling each category's input and output power, each
    row being a 5-minute interval.
  * Create an "Unknown" category if any DUID has not been categorized.
  * Save the dataframe as an Excel spreadsheet.

If run using Idle (Python's Integrated Development and Learning
Environment) the roughly 100MB of the 12 month dataframe can be saved to
an excel spreadsheet using the command nemweb.save_5_minute_dataframe().
  * The file save can take a long time and Excel typically takes
	minutes to load it.

### Repairing Missing Data
Sometimes there are short intervals of missing data. For example, nine
intervals from 1:30 PM to 2:10 PM on 5th September 2024 and four
intervals from 1:30 PM to 1:45 PM on 19th November 2024 are missing.
These durations are sufficiently short that linear estimates between the
good neighbouring data can repair this missing data.

### File Repositories
NEMWeb uses three file repositories:

* the remote source AEMO's archives at nemweb.com.au,
* an optional local mirror copy of the AEMO's remote archive and
* a local data directory for storing NEMWeb's compact pickle data.

NEMWeb takes a while to download and read in remote source files. The
user may optionally download a copy of these ARCHIVE files using, for
example wget.exe, to provide a local mirror copy of the ARCHIVE files to
speed up the time it takes NEMWeb's to read in the source data. The data
from each daily Dispatch_SCADA or weekly ROOFTOP_PV ZIP file is then
saved as Python pickle data in the local data directory using the
original ZIP filename but with a .pkl extension.
