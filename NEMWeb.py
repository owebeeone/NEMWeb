"""Easy access to Australia's electricity data from nemweb.com.au.

Quick Summary
=============

There are three key files - NEMWeb.py, 'DUID Categories.csv' and
requirements.txt. Follow the usual Python procedures to use pip to
install all the dependencies in requirements.txt.

The first time NEMWeb.py is run (on Windows) it will read in the most
recent 12 calendar months of electricity generation data from
nemweb.com.au, saving binary PKL files to %APPDATA%\\NEMWeb and saving
'NEM Year by Category.xlsx' and 'NEM In Out List.xlsx' to
%USERPROFILE%\\Documents. Progress is displayed and three graphs are
shown.

NEMWeb.py runs much faster once the binary PKL files exist.

Background
==========
The Australian National Electricity Market (NEM) covers Australia's
eastern states plus South Australia generating an average of about 25GW
power. The Australian Energy Market Operator (AEMO) operates the NEM
grid. See nemweb.com.au for more information.

Source Data
===========
nemweb.com.au/REPORTS/ARCHIVE/Dispatch_SCADA

    This web directory contains daily ZIP files of 13 recent months of
    data of 'dispatchable units' attached to the NEM grid. A
    dispatchable unit is a generator or storage plant and is identified
    using a Dispatchable Unit Identifier (DUID). Each daily ZIP file
    contains 288 5-minute interval ZIP files. Each 5-minute interval ZIP
    file contains a CSV file with a list of records containing, notably,
    the MW power level for each active DUID.

    Note: 288 = 60min/5min * 24hr

nemweb.com.au/REPORTS/ARCHIVE/ROOFTOP_PV/ACTUAL

    This web directory contains Thursday-to-Wednesday weekly ZIP files
    of 13 recent months of data of rooftop solar generation. Each weekly
    ZIP file contains 336 30-minute interval ZIP files. Each 30-minute
    interval ZIP file contains a CSV file with a list of records
    containing, notably, the MW power level for various regions. Only
    the total state tallies specified by the five region identifiers
    NSW1, QLD1, SA1, TAS1 and VIC1 are used. The 30-minute interval data
    is linearly interpolated to 5-minute interval data for inclusion in
    a table with the 5-minute interval Dispatch_SCADA data.

    Note: 336 = 30min/60min * 24hr * 7days

NEM Registration and Exemption List.xlsx

    This XLSX file contains extensive information on the NEM grid that
    is updated periodically. To locate the latest file go to
        aemo.com.au/energy-systems/electricity/
            national-electricity-market-nem/participate-in-the-market/
            registration
    then locate the 'NEM Registration and Exemption List' download link.
    The 'PU and Scheduled Loads' tab in the XLSX file contains relevant
    information on dispatch units.

    The data in this file is used to help identify the category to be
    assigned to each DUID in the NEMWeb.py project's file called 'DUID
    Categories.csv'.

Program Output
==============
The default output is two files;
    - 'NEM Year by Category.xlsx'

        There is one spreadsheet with either 105120 or 105408 rows for
        all the 5-minute intervals over the past 12 calendar months. A
        file called 'DUID Categories.csv' is used by NEMWeb.py to assign
        a category to each DUID plus the five rooftop solar regions. The
        XLSX file has a negative column and a positive columns for each
        category found in 'DUID Categories.csv', e.g. -battery,
        +battery. Negative columns specify the average MW that category
        took from the grid over the year and vice versa for positive
        columns.

        Note: 105120 = 60min/5min * 24hr * 365days
              105408 = 60min/5min * 24hr * 366days - leap year

        corresponding to 365 days or 366 days for a leap year. There are
        two sets of columns, the first set

    - 'NEM In Out List.xlsx'

        There is one spreadsheet with a row for each active DUID plus
        the five rooftop solar regions. Three data columns are titled
        In, Out and Out/In. The In and Out columns are the average MW
        power taken from or put onto the grid over the 12 months. The
        Out/In column helps identify battery and pumped hydro storage
        DUIDs.

It is possible to save the complete 5-minute interval data for all the
DUIDs and rooftop solar regions but, with 470 DUIDs, the dataframe
contains almost 50M cells which takes a long time to save.

DUID Categories.csv
===================

'DUID Categories.csv' is used to categorise the DUIDs and rooftop solar
regions. Column headings are DUID, Cat, 'Batt Load', 'NEM Tech Type' and
'NEM Fuel Type'. NEMWeb.py only uses the first three columns. The other
two columns are copied from the 'PU and Scheduled Loads' tab in 'NEM
Registration and Exemption List.xlsx' to aid in identifying a DUID's
category.

Each DUID corresponds to a piece of machinery that transfers power to or
from the grid. Any category name can be used as NEMWeb.py builds the
list of categories when loading 'DUID Categories.csv'. Most DUIDs report
power added to the grid with a positive value and power drawn from the
grid with a negative value. Battery loads draw power from the grid but
report positive values so the 'Batt load' column identifies these DUIDs.

If a new DUID appears in the nemweb.com.au data the program creates an
'other' category and reports this DUID in the program output. This new
DUID should to be added to 'DUID Categories.csv' to avoid the 'other'
category.


NEMWeb.py Architecture
======================
Classes
-------
    NEMWebCommon
        Base class with nemweb.com.au source data functionality.
    DispatchSCADA(NEMWebCommon)
        Reads daily ZIP files in the Dispatch_SCADA web directory.
    RooftopPV(NEMWebCommon)
        Reads weekly ZIP files in the ROOFTOP_PV/ACTUAL web directory.
    NEMWeb
        Main class.

Methods
-------
main(mirror_dir=None,
     data_dir='%APPDATA%\\nemweb\\',
     output_dir='%USERPROFILE%\\Documents\\')

    data_dir    A local directory for storing daily and weekly pickle
                files each created from a source ZIP file.

    mirror_dir  An optional local directory with previously downloaded
                copies of daily and weekly ZIP files from Dispatch_SCADA
                and ROOFTOP_PV/ACTUAL web directories.

    Performs the following actions.
        - Obtains the lists of daily and weekly files from
          nemweb.com.au, the local mirror, if present, and the local
          data directory, if the data exists.
        - Determines the most recent available 12 calendar months.
        - Loads the daily and weekly data. After a ZIP file is scanned
          the data is saved as a pickle file for faster loading next
          time.
        - Create df_five_minute and initialise it with the daily data.
        - Show an interactive graph of the GW total of raw daily data.
        - Repair any missing data by linear interpolatation.
        - Show an interactive graph of the GW total of repaired daily
          data.
        - Add the weekly data by linear interpolation.
        - Show an interactive graph of the GW total of the combined
          data.
        - Tally the categories and save the two XLSX files.
        - Return the NEMWeb data structure containing all the data for
          potential further or interactive use.

Notable Properties
------------------
    nemweb
        .dispatch_scada
            .daily_data[]   365 or 366 days of daily data.
        .rooftop_pv
            .weekly_data[]  52 to 54 weeks of weekly data.
        .df_five_minute     A dataframe with 12 calendar months of
                            5-minute intervals for all active DUIDs and
                            the five rooftop solar regions. About 50M
                            cells!
        .df_in_out          A dataframe tallying negative and positive
                            values in each column of df_five_minute.
                            Saved to 'NEM In Out List.xlsx'.
        .df_categories      A dataframe tallying negative and positive
                            values in each row of df_five_minute
                            according to the categories specified in
                            'DUID Categories.csv'. Saved to
                            'NEM Year by Category.xlsx'.

Missing Data
============
Very occasionally a short sequence of data is missing. Missing
Dispatch_SCADA sequences are estimated by linearly interpolating between
the neightbouring values. Missing rooftop solar data is detected but not
repaired.

"""

__version__ = "0.4"  # 22 Jan 2025
__author__ = "John Hilton"

import os
from re import search  # regular expressions
from io import BytesIO, TextIOWrapper
import time
from zipfile import ZipFile
from csv import reader
from datetime import datetime, timedelta
from numpy import zeros, concatenate, ndarray
from pickle import dump, load as pickle_load
from requests import get as requests_get
from bs4 import BeautifulSoup
from glob import glob
from pathlib import Path
from os.path import exists as path_exists, basename as path_basename, dirname as path_dirname
from os import getenv as os_getenv
from urllib.parse import urlparse
from pandas import date_range, DataFrame, concat as pandas_concat
import matplotlib.pyplot as plt
from platformdirs import user_data_dir, user_documents_dir
from dataclasses import dataclass, field


@dataclass
class PlotManager:
    """Manage a collection of plots."""

    plots: list["PlotBase"] = field(init=False, default_factory=list)
    is_interactive: bool = field(init=False, default=False)
    is_closed: bool = field(init=False, default=False)

    def pause_on_close(self, timeout_seconds: float = 5) -> None:
        """Pause on close.
        Wait timeout_seconds seconds before closing the plots unless interacts with the plots.
        """
        start_time = time.time()
        is_timed_out = False
        while not self.is_closed and (not is_timed_out or self.is_interactive):
            manager = plt.get_current_fig_manager()
            if manager is not None:
                canvas = manager.canvas
                if canvas.figure.stale:
                    # Update the screen as the canvas wasn't fully drawn yet.
                    canvas.draw_idle()
                canvas.start_event_loop(0.1)
            is_timed_out = time.time() - start_time >= timeout_seconds
        return


@dataclass
class PlotBase:
    """Base class for a plot."""

    title: str
    plot_manager: PlotManager
    cid_close: object = field(init=False)
    cid_click: object = field(init=False)

    def onclick(self, event):
        """Set the plot manager's is_interactive flag to True when the plot is clicked.
        This is used to stop the plots from closing.
        """
        self.plot_manager.is_interactive = True

    def onclose(self, event):
        """Set the plot manager's is_closed flag to True when the plot is closed and should
        cause the rest of the plots to close.
        """
        self.plot_manager.is_closed = True

    def init_plot(self) -> tuple[plt.Figure, plt.Axes]:
        """Initialise the plot."""
        # Plot the results.
        fig, ax = plt.subplots()

        self.cid_close = fig.canvas.mpl_connect("close_event", lambda e: self.onclose(e))
        self.cid_click = fig.canvas.mpl_connect("button_press_event", lambda e: self.onclick(e))

        # Add a title to the plot.
        ax.set_title(self.title)
        # ax.set_aspect('equal')
        # Use the "tight" layout.
        plt.tight_layout(pad=0, h_pad=0, w_pad=0, rect=[0, 0, 1, 1])
        ax.set_in_layout(False)
        plt.ion()
        # Render the image to the size of the screen.
        fig.set_size_inches(10.5, 10.5)

        return fig, ax


class Blank:
    """An empty utility class providing the (key,value) features of
    dict().
    """

    pass


@dataclass
class NEMWebCommon:
    """Base class providing common functionality for Dispatch_SCADA and
    Rooftop_PV data.
    """

    mirror_dir: str
    data_dir: str
    filelist: list[tuple[bool, str]] = field(init=False)

    WEB_ARCHIVE_URL = "https://www.nemweb.com.au/REPORTS/ARCHIVE/"

    def __post_init__(self):
        """Builds three sorted lists of daily or weekly data files then
        merges them into one list prioritizing PKL files over local
        mirror files over nemweb.com.au ZIP files.
        """
        # Build
        #   .remote_zipfile_list,
        #   .local_zipfile_list and
        #   .local_pklfile_list.
        self.build_file_lists()

        # Select the daily files from the three repositories.
        self.filelist = self.select_files()

    def build_file_lists(self):
        """Builds three sorted lists of daily or weekly data files;
        local PKL files, local mirror ZIP files and nemweb.com.au ZIP
        files.
        """

        cls = self.__class__

        # Build a sorted list of the remote ZIP files.
        webdir = self.WEB_ARCHIVE_URL + cls.WEB_SUBDIR
        rzlist = self.build_url_list(webdir)
        rzlist.sort()  # The list appears to always be sorted anyway.

        # Sift out Rooftop_PV SATELLITE files by assuming the first
        # filename length is valid.
        valid_len = len(rzlist[0])
        while len(rzlist[-1]) != valid_len:
            rzlist.pop()
        self.remote_zipfile_list = rzlist

        if self.mirror_dir is not None:
            # Build a sorted list of the local daily ZIP files.
            path = "".join((self.mirror_dir, cls.WEB_SUBDIR))
            pattern = "".join((path, cls.FILENAME_PREFIX, "*.zip"))
            lzlist = glob(pattern)
            lzlist.sort()
            self.local_zipfile_list = lzlist
        else:
            self.local_zipfile_list = []

        # Build a sorted list of the local daily pickle files.
        path = "".join((self.data_dir, cls.WEB_SUBDIR))
        pattern = "".join((path, cls.FILENAME_PREFIX, "*.pkl"))
        lplist = glob(pattern)
        lplist.sort()
        self.local_pklfile_list = lplist

    @staticmethod
    def build_url_list(webdir):
        """Return a list of urls to the files in webdir.
        Arguments:
            webdir:  The web directory containing ZIP data files.
        Returns:
            A list of urls of each ZIP file in webdir.
        """

        page = requests_get(webdir).text  # read the page from the web

        soup = BeautifulSoup(page, "html.parser")  # parse the html

        # Scan the list to create a list of ZIP file URLs.
        o = urlparse(webdir)
        return [
            o.scheme + "://" + o.hostname + node.get("href")
            for node in soup.find_all("a")
            if node.get("href").endswith("zip")
        ]

    def select_files(self) -> list[tuple[bool, str]]:
        """Merge the three data file lists into one prioritizing local
        PKL data files over local ZIP files then nemweb.com.au ZIP
        files.
        """

        rzlist = self.remote_zipfile_list
        lzlist = self.local_zipfile_list
        lplist = self.local_pklfile_list

        # Process the three lists prioritizing pkl over local over
        # remote.
        lp = 0
        lz = 0
        rz = 0
        lpflag = lp < len(lplist)
        lzflag = lz < len(lzlist)
        rzflag = rz < len(rzlist)

        # Determine the first and last days.
        first_ymd = "9"
        last_ymd = "0"
        if lpflag:
            first_ymd = lplist[0][-12:-4]
            last_ymd = lplist[-1][-12:-4]
            lp_ymd = first_ymd
        if lzflag:
            first = lzlist[0][-12:-4]
            last = lzlist[-1][-12:-4]
            if first_ymd > first:
                first_ymd = first
            if last_ymd < last:
                last_ymd = last
            lz_ymd = first
        if rzflag:
            first = rzlist[0][-12:-4]
            last = rzlist[-1][-12:-4]
            if first_ymd > first:
                first_ymd = first
            if last_ymd < last:
                last_ymd = last
            rz_ymd = first

        # Determine the total number of days.
        first = datetime(int(first_ymd[0:4]), int(first_ymd[4:6]), int(first_ymd[6:8]))
        last = datetime(int(last_ymd[0:4]), int(last_ymd[4:6]), int(last_ymd[6:8]))
        total_days = (last - first).days + 1

        print(f"  Combined local and remote date range for {self.WEB_SUBDIR[:-1]}:")

        print(
            f"   From {first.strftime('%d %b %Y')}"
            f" through {last.strftime('%d %b %Y')}"
            f" - {total_days} days"
        )

        # Determine the last day.
        ymd = "0"
        if lpflag and ymd > lp_ymd:
            ymd = lp_ymd
            choice = "lp"
        if lzflag and ymd > lz_ymd:
            ymd = lz_ymd
            choice = "lz"
        if rzflag and ymd > rz_ymd:
            ymd = rz_ymd
            choice = "rz"

        filelist = []
        while True:
            # Select the interval's file.
            choice = None
            ymd = "9"
            if lpflag and ymd > lp_ymd:
                ymd = lp_ymd
                choice = "lp"
            if lzflag and ymd > lz_ymd:
                ymd = lz_ymd
                choice = "lz"
            if rzflag and ymd > rz_ymd:
                ymd = rz_ymd
                choice = "rz"

            # Finished?
            if choice is None:
                break

            # Select the interval's file.
            match choice:
                case "rz":
                    filelist.append((True, rzlist[rz]))
                case "lz":
                    filelist.append((True, lzlist[lz]))
                case "lp":
                    filelist.append((False, lplist[lp]))

            # Update the indices.
            if rzflag and rz_ymd == ymd:
                rz += 1
                rzflag = rz < len(rzlist)
                if rzflag:
                    rz_ymd = rzlist[rz][-12:-4]
            if lzflag and lz_ymd == ymd:
                lz += 1
                lzflag = lz < len(lzlist)
                if lzflag:
                    lz_ymd = lzlist[lz][-12:-4]
            if lpflag and lp_ymd == ymd:
                lp += 1
                lpflag = lp < len(lplist)
                if lpflag:
                    lp_ymd = lplist[lp][-12:-4]

        return filelist

    def build_twelve_month_list(self):
        """Builds a list of daily or weekly ZIP files from
        self.start_date to self.end_date.
        """

        cls = self.__class__
        MINUTES_PER_DAY = 24 * 60
        days_per_file = int(cls.INTERVALS_PER_FILE * cls.MINUTES_PER_INTERVAL / MINUTES_PER_DAY)
        number_of_files = int((self.end_date - self.start_date).days / days_per_file)
        list_ = []

        start_yyyymmdd = self.start_date.strftime("%Y%m%d")
        end_yyyymmdd = self.end_date.strftime("%Y%m%d")
        flag = True
        for o in self.filelist:
            is_zip, filepath = o
            if flag:
                if filepath[-12:-4] < start_yyyymmdd:
                    continue
                flag = False
            # if filepath[-12:-4] >= end_yyyymmdd:
            #    break
            list_.append((is_zip, filepath))
            if len(list_) >= number_of_files:
                break
        return list_

    def load_nem_data(self):
        """Load daily or weekly NEM data into a list of dataframes. If a
        previously created daily or weekly PKL file exists it is loaded
        otherwise the ZIP file is read in and a corresponding PKL file
        saved."""

        data_dir = Path("".join((self.data_dir, self.WEB_SUBDIR)))
        data_dir.mkdir(parents=True, exist_ok=True)

        # Each day or week's data will be added to a list.
        nem_MW_df_list = []

        for is_zip, filepath in self.twelve_month_list:
            self.load = Blank()
            load = self.load
            if is_zip:
                self.load_nem_data_file(filepath)
                # Save the day's data to a pickle file.
                fullpath = Path(data_dir) / Path(load.filename).with_suffix(".pkl")
                print(f"Saving {fullpath}")
                with open(fullpath, "wb") as ofile:
                    # dump(load.duid_list,ofile)
                    dump(load.nem_MW_df, ofile)
            else:
                self.load_pkl_file(filepath)

            # Append the dataframe.
            nem_MW_df_list.append(load.nem_MW_df)

        return nem_MW_df_list

    def load_pkl_file(self, fullpath):
        """Load a daily or weekly dataframe from a local PKL file into
        self.load.
        """

        load = self.load
        load.filename = path_basename(fullpath)
        # print(f'Loading {load.filename}')
        with open(fullpath, "rb") as ifile:
            # load.duid_list = pickle_load(ifile)
            load.nem_MW_df = pickle_load(ifile)

    def parse_csvfile(self, csv_reader):
        print("Override parse_csvfile()!")
        raise Exception

    def initialise_data_MW(self):
        print("Override initialise_data_MW()!")
        raise Exception

    def finalise_data_MW(self):
        print("Override finalise_data_MW()!")
        raise Exception

    def load_nem_data_file(self, source_file, validate=False):
        """Load the MW power data from either a Dispatch_SCADA or a
        Rooftop_PV data file into a dataframe. Setting validate to True
        will enable some validation of the input file.
        """

        load = self.load
        load.source_file = source_file
        load.validate = validate

        # Create and initialise the data_MW array and duid_list.
        self.initialise_data_MW()
        load.duid_list = []
        load.interval_stamp = None

        # Track the inner ZIP and CSV files for warning/error
        # notification.
        load.track = Blank()
        load.track.last_inner_zip = None

        # Extract the filename and parse the day's date.
        load.filename = path_basename(source_file)
        o = Blank()
        o.year = int(load.filename[-12:-8])
        o.month = int(load.filename[-8:-6])
        o.day = int(load.filename[-6:-4])
        load.date = o

        # Start time is end of the first interval.
        load.start_stamp = datetime(o.year, o.month, o.day, hour=0, minute=self.BEGIN_MINUTE)

        if load.validate:
            # yyyy/mm/dd hh:mm:ss
            strings = (
                load.filename[-12:-8],
                "/",
                self.filename[-8:-6],
                "/",
                self.filename[-6:-4],
                "xx:xx:00",
            )
            load.interval_stamp = "".join(strings)

        if source_file.startswith("http"):
            # Cloud file.
            print(f"Downloading {load.filename}")
            r = requests_get(source_file, stream=True)
            zf = ZipFile(BytesIO(r.content))
        elif path_exists(source_file):
            # Local file.
            print("Reading " + load.filename)
            zf = ZipFile(source_file)
        else:
            print(f"Unknown url or path: {source_file}")

        # Confirm 288 (=24*12) or 336 (=7*24*2) files in the ZIP file.
        if len(zf.namelist()) != self.INTERVALS_PER_FILE:
            num = len(zf.namelist())
            print(f"Warning:  {load.filename} has {num} files - expected {self.INTERVALS_PER_FILE}")

        # Extract the ZIP files from either the Dispatch_SCADA day ZIP
        # file or the Rooftop_PV weekly ZIP file.
        for filename in zf.namelist():
            if search(r"\.zip$", filename):
                load.inner_zip_filename = filename
                content = BytesIO(zf.read(filename))
                self.read_interval_zip(load, ZipFile(content))
            else:
                print(f"Unexpected file: {filename}")

        self.finalise_data_MW()

        index = date_range(
            start=load.start_stamp,
            periods=self.INTERVALS_PER_FILE,
            freq=timedelta(minutes=self.MINUTES_PER_INTERVAL),
        )
        load.nem_MW_df = DataFrame(load.data_MW, index=index, columns=load.duid_list)
        print("Done")

    def read_interval_zip(self, load, zf):
        """Read in a five or thirty minute interval ZIP file."""

        for filename in zf.namelist():  # should just be one
            if search(r"\.[Cc][Ss][Vv]$", filename):
                load.csv_filename = filename
                tail = filename[len(self.FILENAME_PREFIX) :]
                year = int(tail[:4])
                month = int(tail[4:6])
                day = int(tail[6:8])
                hour = int(tail[8:10])
                minute = int(tail[10:12])
                csv_stamp = datetime(year, month, day, hour, minute)
                seconds = (csv_stamp - load.start_stamp).total_seconds()
                load.interval = int(seconds / (60 * self.MINUTES_PER_INTERVAL) + 0.5)
                if load.validate:
                    # yyyy/mm/dd hh:mm:ss
                    load.interval_stamp[12:17] = f"{hour:02d}:{minute:02d}"
                content = BytesIO(zf.read(filename))
                wrapper = TextIOWrapper(content, encoding="utf-8")
                self.parse_csvfile(load, reader(wrapper))
            else:
                print(f"Unexpected file: {filename}")


class DispatchSCADA(NEMWebCommon):
    """A class to handle nemweb.com.au Dispatch_SCADA daily data."""

    WEB_SUBDIR = "Dispatch_SCADA/"
    FILENAME_PREFIX = "PUBLIC_DISPATCHSCADA_"
    MINUTES_PER_INTERVAL = 5
    BEGIN_MINUTE = 5  # First SETTLEMENTDATE is five minutes past
    # midnight - 00:05:00. The end of the interval.
    INTERVALS_PER_FILE = 24 * int(60 / 5)  # 288 in one day

    def __init__(self, mirror_dir, data_dir):
        """Call the base class initialiser."""

        super().__init__(mirror_dir, data_dir)

    def initialise_data_MW(self):
        """Dispatch_SCADA specific initialisation."""

        self.load.data_MW = [None for i in range(self.INTERVALS_PER_FILE)]

    def finalise_data_MW(self):
        """Dispatch_SCADA specific finalisation."""

        # Replace the array of arrays with a numpy ndarray.
        cls = self.__class__
        new_data_MW = zeros((cls.INTERVALS_PER_FILE, len(self.load.duid_list)))
        for row, row_data in enumerate(self.load.data_MW):
            if row_data is not None:  # This happens when data is
                # missing.
                for col, power in enumerate(row_data):
                    new_data_MW[row, col] = power
        self.load.data_MW = new_data_MW

    def parse_csvfile(self, load, csvreader):
        """Reads a NEM SCADA CSV file containing one day's worth of
        generator data.
        """

        # Filling in this interval.
        # Create an array of zeros to match the current duid_list.
        interval_MW = [0.0 for i in range(len(load.duid_list))]
        for lineno, row in enumerate(csvreader):
            match row[0]:
                case "C":
                    # Ignore the first and last records.
                    # C,NEMP.WORLD,DISPATCHSCADA,AEMO,PUBLIC,2022/07/25,
                    #   00:00:11,0000000367567721,DISPATCHSCADA,
                    #   0000000367567715
                    pass
                case "I":
                    # Process the headings.
                    # I,DISPATCH,UNIT_SCADA,1,SETTLEMENTDATE,DUID,
                    #   SCADAVALUE
                    if not hasattr(load, "timedate_col"):
                        load.timedate_col = row.index("SETTLEMENTDATE")
                        load.duid_col = row.index("DUID")
                        load.MW_col = row.index("SCADAVALUE")
                case "D":
                    # D,DISPATCH,UNIT_SCADA,1,"2022/07/25 00:05:00",
                    #   BARCSF1,0.10
                    # Scan the date and time if required.
                    power = row[load.MW_col]
                    if load.validate or power != "0":
                        dt = datetime.strptime(row[4], "%Y/%m/%d %H:%M:%S")

                    # Validation check should never fail.
                    if load.validate:
                        if (
                            row[1] != "DISPATCH"
                            or row[2] != "UNIT_SCADA"
                            or row[3] != "1"
                            or dt != load.interval_stamp
                        ):
                            print("Invalid CSV line " + csvreader.line_num)
                            continue

                    # Skip 0MW entries
                    if power == "0":  # skip 0.0MW entries
                        continue

                    # Obtain the duid index.
                    duid = row[load.duid_col]
                    if duid in load.duid_list:
                        duid_index = load.duid_list.index(duid)
                    else:  # new duid
                        duid_index = len(load.duid_list)
                        load.duid_list.append(duid)
                        interval_MW.append(0.0)
                    ##                        # Ensure data_MW is big enough
                    ##                        if load.data_MW.shape[1] < duid_index:
                    ##                            interval_MW = None
                    ##                            # Grow by 10 more DUIDs.
                    ##                            load.data_MW.resize(288,duid_index+10)
                    ##                            interval_MW = self.data_MW[load.interval]

                    # Read the MW data into self.data_MW.
                    try:
                        interval_MW[duid_index] = float(power)
                    except ValueError:
                        print(
                            f"Failed to read line {lineno + 1} of\n"
                            + f"{load.csv_filename}\n"
                            + f"in {load.inner_zip_filename}\n"
                            + f"in {load.filename}"
                        )
                case _:
                    # Unexpected - print it
                    print(f"Unexpected row at line {lineno + 1}\n", row)
        load.data_MW[load.interval] = interval_MW


class RooftopPV(NEMWebCommon):
    """A class to handle nemweb.com.au Rooftop_PV weekly data."""

    WEB_SUBDIR = "ROOFTOP_PV/ACTUAL/"
    FILENAME_PREFIX = "PUBLIC_ROOFTOP_PV_ACTUAL_MEASUREMENT_"
    MINUTES_PER_INTERVAL = 30
    BEGIN_MINUTE = 0  # Midnight 00:00:00 - See discussion below.
    INTERVALS_PER_FILE = 7 * 24 * int(60 / 30)  # 336 in one week
    REGION_IDS = ["NSW1", "QLD1", "SA1", "TAS1", "VIC1"]

    """Discussion on BEGIN_MINUTE...

    An example weekly filename is
        <prefix>_20240118.zip
    where <prefix> is "PUBLIC_ROOFTOP_PV_ACTUAL_MEASUREMENT"
    The first 30 minute interval ZIP file in the weekly ZIP file is
        <prefix>_20240118000000_0000000408687300.zip
    The INTERVAL_DATETIME values in this file are
        17/01/2024  11:30:00 PM
        
    AEMO states that both Dispatch_SCADA and Rooftop_PV interval data is
    timestamped at the end of the interval. From the example above the
    interval's timestamp, 20240118000000, is 30 minutes after the
    INTERVAL_DATETIME column's timestamp, they do not match. Therefore,
    INTERVAL_DATETIME is presumed to indicate the start of the interval.
    This program will use the filename timestamp of 00:00:00 to match
    the end-of-interval nature of the 5 minute Dispatch_SCADA data.

    Also, a special case needs to be handled where the end of the last
    week matches the last day of a 12 month period as 30 minutes of data
    is missing. It is safe to assume zero power as this is nighttime.
    """

    def __init__(self, mirror_dir, data_dir):
        """Call the base class initialiser."""

        super().__init__(mirror_dir, data_dir)

    def initialise_data_MW(self):
        """Dispatch_SCADA specific initialisation."""

        # As the number of IDS are known create a numpy ndarray.
        cls = self.__class__
        shape = (cls.INTERVALS_PER_FILE, len(cls.REGION_IDS))
        self.load.data_MW = zeros(shape)

    def finalise_data_MW(self):
        """Dispatch_SCADA specific finalisation."""

        # Nothing to do.
        cls = self.__class__
        self.load.duid_list = cls.REGION_IDS

    def parse_csvfile(self, load, csv_reader):
        """Reads a NEM rooftop solar CSV file containing one week's
        worth of rooftop solar data.
        """

        cls = self.__class__

        # Filling in this interval.
        # interval_MW = [0.0 for i in range(len(cls.REGION_IDS))]
        # load.duid_list = cls.REGION_IDS
        # print(f'self interval: {self.interval}') #debug
        # interval_MW = load.data_MW[load.interval]
        # Loop over each record.
        for lineno, row in enumerate(csv_reader):
            match row[0]:
                case "C":
                    # Ignore the first and last records.
                    # C,NEMP.WORLD,ROOFTOP_PV_ACTUAL_MEASUREMENT,AEMO,
                    #   PUBLIC,2024/09/05,15:00:01,0000000431913499,
                    #   DEMAND,0000000431913499
                    # C,"END OF REPORT",13
                    pass
                case "I":
                    # Headings record.
                    # Initialise column indices.
                    # I,ROOFTOP,ACTUAL,2,INTERVAL_DATETIME,REGIONID,
                    #   POWER,QI,TYPE,LASTCHANGED
                    if not hasattr(load, "interval_col"):
                        load.interval_col = row.index("INTERVAL_DATETIME")
                        load.regionid_col = row.index("REGIONID")
                        load.power_col = row.index("POWER")
                case "D":
                    # Data record.
                    # D,ROOFTOP,ACTUAL,2,"2024/09/05 14:30:00",NSW1,...
                    #   3088.502,1,MEASUREMENT,"2024/09/05 14:49:17"

                    # Obtain the region index - ignore non-identified
                    # regions.
                    o = row[load.regionid_col]
                    if o not in cls.REGION_IDS:
                        continue
                    state_index = cls.REGION_IDS.index(o)
                    # timestamp = row[load.interval_col]

                    # Read the MW data into data_MW.
                    power = row[load.power_col]
                    if power == "0":
                        continue
                    if power == "":
                        if load.track.last_inner_zip is None:
                            print(f"Missing power(s) in\n  {load.filename}")
                        if load.track.last_inner_zip != load.inner_zip_filename:
                            print(f"  {load.inner_zip_filename}")
                            load.track.last_inner_zip = load.inner_zip_filename
                            load.track.last_csv = None
                        if load.track.last_csv != load.csv_filename:
                            print(f"    {load.csv_filename}")
                            load.track.lastcsv = load.csv_filename
                        region_id = self.REGION_IDS[state_index]
                        print(f"      Record no. {lineno + 1} - {region_id}")
                    else:
                        try:
                            # Read the MW data.
                            value = float(power)
                        except ValueError:
                            print(
                                f"Failed to read line {lineno + 1} of"
                                f"\n{load.csv_filename}"
                                f"\nin {load.inner_zip_filename}"
                                f"\nin {load.filename}"
                            )
                            value = float(0)
                        load.data_MW[load.interval, state_index] = value
                case _:
                    print(f"Unexpected record at line {lineno + 1}\n", row)

    def build_df_thirty_minute(self, start_date):
        """Concatenate the more than 12 months of weekly rooftop solar
        data to 12 months including the first midnight for later
        interpolation.
        """

        df = pandas_concat(self.nem_MW_df_list)
        end_date = datetime(start_date.year + 1, start_date.month, 1)
        start_row = df.index.get_loc(start_date)
        end_row = df.index.get_loc(end_date)
        df = df.iloc[start_row : end_row + 1]
        return df


class NEMWeb:
    """Provide easy access to nemweb.com.au archive data.
    o First three lists are generated for Dispatch_SCADA and ROOFTOP_PV
      archives:
        (i) the archive files available on www.nemweb.com.au,
        (ii) local ZIP file copies, if available, and
        (iii) the list of pickle files from previously loaded ZIP files.
    o Data for the previous full 12 months is loaded for Dispatch_SCADA
      and ROOFTOP_PV. The dataframes from any loaded ZIP files are saved
      as pickle data files.
    o A 12 month dataframe is built from the Dispatch_SCADA data.
    o Any of the rare segments of missing Dispatch_SCADA data is
      repaired by interpolation.
    o The ROOFTOP_PV data is added to the 12 month dataframe
      interpolating the 30-minute intervals to 5-minute intervals.
      NOTE: Missing rooftop solar data is detected but yet to be
      repaired.
    o A 'DUID categories.csv' file is loaded and used to produce a
      12 month dataframe having category column headings.
    o The categories dataframe is saved as an Excel spreadsheet.
    """

    def __init__(self, mirror_dir, data_dir, output_dir):
        """(i) Initialise, (ii) obtain the filenames of the
        Dispatch_SCADA daily data and the ROOFTOP_PV/ACTUAL weekly data,
        (iii) determine the most recent available 12 calenday month data
        and (iv) read in that data.
        """

        # Save the current time for reporting elapsed seconds.
        self.begin_time = datetime.now().replace(microsecond=0)

        # Save the local directory names.
        self.mirror_dir = mirror_dir
        self.data_dir = data_dir
        self.output_dir = output_dir

        print(f"Obtaining filenames from\n    {NEMWebCommon.WEB_ARCHIVE_URL[8:-1]}")
        if mirror_dir is not None:
            print(f"\n    {mirror_dir}")
        print(f"\n    {self.data_dir[:-1]}")

        self.dispatch_scada = DispatchSCADA(self.mirror_dir, self.data_dir)
        self.rooftop_pv = RooftopPV(self.mirror_dir, self.data_dir)
        self.show_elapsed()

        print("Determining the most recent 12 months available.")
        self.calculate_most_recent_full_twelve_months()
        self.show_elapsed()

        for o in (self.dispatch_scada, self.rooftop_pv):
            o.twelve_month_list = o.build_twelve_month_list()
            print(f"Loading NEM data: {o.WEB_SUBDIR[:-1]}")
            o.nem_MW_df_list = o.load_nem_data()
            self.show_elapsed()

    def build_df_five_minute(self):
        """Build a dataframe with 5-minute interval rows and columns for
        all the DUIDs and the five rooftop solar regions.
        """

        ds = self.dispatch_scada

        # Create the row timestamps.
        start_timestamp = ds.start_date.replace(minute=ds.BEGIN_MINUTE)
        periods = (ds.end_date - ds.start_date).days * ds.INTERVALS_PER_FILE
        freq = timedelta(minutes=ds.MINUTES_PER_INTERVAL)
        all_timestamps = date_range(start=start_timestamp, periods=periods, freq=freq)

        # Create the column headings - a sorted list of all DUIDs plus
        # the region ids in the data.
        all_sorted_ids = (
            list(set([duid for item in ds.nem_MW_df_list for duid in item.axes[1]]))
            + self.rooftop_pv.REGION_IDS
        )
        all_sorted_ids.sort()

        # Create the numpy ndarray.
        all_data_MW = zeros((len(all_timestamps), len(all_sorted_ids)))

        # Populate the ndarray with Dispatch_SCADA data.
        num_data_points = 0
        for i, df in enumerate(ds.nem_MW_df_list):
            if (i + 1) % 60 == 0:
                print(f" {i + 1}", end="")
            first_row = (df.axes[0][0] - start_timestamp).days * ds.INTERVALS_PER_FILE
            for duid in df.axes[1]:
                col = all_sorted_ids.index(duid)
                for i, power in enumerate(df[duid]):
                    if power != 0:
                        all_data_MW[first_row + i, col] = power
                        num_data_points += 1
        print("")

        self.dispatch_scada.num_data_points = num_data_points
        return DataFrame(all_data_MW, all_timestamps, all_sorted_ids)

    def add_rooftop_solar_data(self):
        """Add the rooftop solar data to the main dataframe."""
        rtdf = self.rooftop_pv.df_thirty_minute
        nwdf = self.df_five_minute
        num_data_points = 0
        # Step over each column.
        for region_id in rtdf.axes[1]:
            print(f" {region_id}", end="")
            # Lookup the rooftop solar column.
            rt_iter = iter(rtdf[region_id].iloc)

            # Lookup the main dataframe column.
            dst_col = nwdf.axes[1].to_list().index(region_id)

            start_power = next(rt_iter)
            if start_power != 0:
                num_data_points += 1
            # Step over each 30-minute interval.
            for row6 in range(0, len(nwdf.axes[0]) - 1, 6):
                end_power = next(rt_iter)
                if end_power != 0:
                    num_data_points += 1
                    if start_power != 0:
                        y = start_power
                        dy = (end_power - start_power) / 6
                        # Interpolate 5-minute intervals.
                        for dst_row in range(row6, row6 + 6):
                            nwdf.iloc[dst_row, dst_col] = y
                            y += dy
                start_power = end_power
        print("")
        self.rooftop_pv.num_data_points = num_data_points

    @staticmethod
    def dateof(s):
        return datetime(int(s[:4]), int(s[4:6]), int(s[6:8]))

    def show_elapsed(self):
        td = datetime.now().replace(microsecond=0) - self.begin_time
        print(f" Done - elapsed h:mm:ss: {str(td)}\n")

    def calculate_most_recent_full_twelve_months(self):
        """Given the daily and weekly file lists, determine the first of
        the month start and end dates. Weekly files contain Thu-Wed data
        and have Thursday filename dates. Determine the start and end
        Thursdays so the weekly data covers the first of the month daily
        range.
        """

        day_list = self.dispatch_scada.filelist
        week_list = self.rooftop_pv.filelist
        dateof = self.dateof

        # Select the earliest end date from the end of the lists.
        day_filename = day_list[-1][1]
        week_filename = week_list[-1][1]
        earliest_end_date = dateof(day_filename[-12:-4]) + timedelta(days=1)
        weekly = dateof(week_filename[-12:-4]) + timedelta(days=7)
        if earliest_end_date > weekly:
            earliest_end_date = weekly

        # Go back to the first day of the month.
        end_daily_day = earliest_end_date.replace(day=1)

        # Set the start date one year earlier and make sure data exists.
        start_daily_day = datetime(end_daily_day.year - 1, end_daily_day.month, 1)

        # Weekly rooftop data begins on a Thursday so set the start and
        # end Thursdays. Bracket the year with the prior and following
        # Thursdays.
        i = (start_daily_day.weekday() - 3) % 7
        start_weekly_day = start_daily_day - timedelta(days=i)
        i = (3 - end_daily_day.weekday()) % 7
        end_weekly_day = end_daily_day + timedelta(days=i)

        # Check the daily and weekly starting data is available.
        day_filename = day_list[0][1]
        week_filename = week_list[0][1]
        if start_daily_day < dateof(day_filename[-12:-4]) or start_weekly_day < dateof(
            week_filename[-12:-4]
        ):
            print("**** Data for a full twelve months is not available. ****")
        self.dispatch_scada.start_date = start_daily_day
        self.dispatch_scada.end_date = end_daily_day
        self.rooftop_pv.start_date = start_weekly_day
        self.rooftop_pv.end_date = end_weekly_day

        adbY = "%a %d %b %Y"
        begin = start_daily_day.strftime(adbY)
        end = (end_daily_day - timedelta(days=1)).strftime(adbY)
        print(f" Most recent 12 months: {begin} - {end}")
        begin = start_weekly_day.strftime(adbY)
        end = (end_weekly_day - timedelta(days=1)).strftime(adbY)
        print(f"  Thursday Weeks: {begin} - {end}")

    def build_cat_and_in_out_dfs(self):
        """Build df_categories and df_in_out from df_five_minute and
        'DUID Categories.csv'.
        """

        df_five_minute = self.df_five_minute
        duid_catidx = self.duid_catidx
        num_cat = len(self.cat_list) + 1  # Allow for 'Unknown'
        unknown_duids = []
        row_count = len(df_five_minute.axes[0])
        cat_MW = zeros([row_count, num_cat * 2])
        in_out_MW = zeros([len(df_five_minute.axes[1]), 3])

        # Total the ins and outs for each DUID.
        for src_col, duid in enumerate(df_five_minute.axes[1]):
            if (src_col + 1) % 50 == 0:
                print(f" {src_col + 1}", end="")
            try:
                cat_col, batt_load = duid_catidx[duid]
            except KeyError:
                cat_col, batt_load = num_cat - 1, False
                unknown_duids.append(duid)
            try:
                for row, srcval in enumerate(df_five_minute[duid]):
                    if srcval != 0:
                        if batt_load:  # Battery loads are reported as
                            # +ve.
                            srcval = -srcval
                        if srcval < 0:
                            cat_MW[row, cat_col] += srcval
                            in_out_MW[src_col, 0] += srcval
                        else:
                            cat_MW[row, num_cat + cat_col] += srcval
                            in_out_MW[src_col, 1] += srcval
            except Exception:
                print("\nException!", end="")
        print("")

        # Calculate the in_out average MW by dividing the totals by the
        # count.
        in_out_MW /= row_count

        # Create a copy of cat_list.
        cat_list = [cat for cat in self.cat_list]
        if len(unknown_duids) > 0:
            print(f' DUIDs categorised as "Unknown": {unknown_duids}')
            cat_list.append("Unknown")
        else:
            # Remove the empty +/- 'Unknown' columns
            cat_MW = concatenate(
                (cat_MW[:, : num_cat - 1], cat_MW[:, num_cat : 2 * num_cat - 1]), axis=1
            )
            num_cat -= 1

        # Calculate the turnaround ratio
        for i in range(in_out_MW.shape[0]):
            if -in_out_MW[i, 0] > 0.1:
                in_out_MW[i, 2] = in_out_MW[i, 1] / -in_out_MW[i, 0]

        columns = []
        for pm in ("-", "+"):
            for cat in cat_list:
                columns.append(pm + cat)
        o = (
            DataFrame(cat_MW, df_five_minute.axes[0], columns),
            DataFrame(in_out_MW, df_five_minute.axes[1], ["In", "Out", "Out/In"]),
        )
        return o

    def save_in_out_and_cat(self):
        """Save df_categories to 'NEM In Out List.xlsx'."""

        fullpath = "".join((self.output_dir, "NEM In Out List.xlsx"))
        print(f"Saving {fullpath}")
        self.df_in_out.to_excel(fullpath)
        self.show_elapsed()
        fullpath = "".join((self.output_dir, "NEM Year by Category.xlsx"))
        print(f"Saving {fullpath}")
        self.df_categories.to_excel(fullpath)
        self.show_elapsed()

    def save_5_minute_dataframe(self):
        """Save df_five_minute to 'NEM 12 Month Generation.xlsx'."""

        # Save the current time for reporting elapsed seconds.
        begin_time = datetime.now().replace(microsecond=0)

        fullpath = "".join((self.output_dir, "NEM 12 Month Generation.xlsx"))
        print(f"Saving {fullpath}")
        self.df_five_minute.to_excel(fullpath)
        td = datetime.now().replace(microsecond=0) - begin_time
        print(f" Done - elapsed h:mm:ss: {str(td)}\n")


def find_indices_below_threshold(nums, threshold):
    """Find all indices below the specified threshold value."""

    return [i for i, num in enumerate(nums) if num < threshold]


def repair_missing_periods(df_MW, row_sum_5min_GW):
    """Repair any missing data periods where total generation is less
    than 2.0 GW by inserting linearly interpolated values derived from
    valid adjacent rows.
    """

    two_GW = 2.0  # GW
    start = None
    for i, power_GW in enumerate(row_sum_5min_GW):
        if start is None:
            # Watch for the start of an invalid period.
            if power_GW < two_GW:
                # Found the start.
                start = i
        elif power_GW >= two_GW:
            # Found the end.
            end = i
            if start > 0:
                # The repair period is bounded by start-1 and end.
                repair(df_MW, start, end, start - 1, end)
            else:
                # The repair period is at the start of df_MW.
                repair(df_MW, start, end, end, end)

            # Finished fixing this period - continue watching for the
            # next.
            start = None

    if start is not None:
        # The repair period is at the end of df_MW.
        end = len(row_sum_5min_GW)
        repair(df_MW, start, end, start - 1, start - 1)


def repair(df_MW, start, end, prior, post):
    """Repair missing data by interpolating from adjacent data."""

    num_intervals = end - start
    start_timestamp = df_MW.axes[0][start]
    end_timestamp = df_MW.axes[0][end - 1]
    print(f" Repairing {num_intervals} intervals from {start_timestamp} to {end_timestamp}")

    accums_GW = df_MW.iloc[prior, :]  # Initialise a row of accumulators.
    if prior == post:
        # Special case when the repair period is at the beginning or end
        # df_MW.
        for i in range(start, end):
            df_MW.iloc[i, :] = accums_GW
    else:
        # Calculate the linearly interpolating iteration deltas.
        deltas_GW = (df_MW.iloc[post, :] - accums_GW) / num_intervals
        # Repair the data.
        for i in range(start, end):
            df_MW.iloc[i, :] = accums_GW
            accums_GW += deltas_GW


@dataclass
class PlotGW5min(PlotBase):
    """Plot five minute interval GW data. This will only show the plot once plot_data() is called."""

    row_sum_5min_GW: ndarray

    def plot_data(self):
        fig, ax = self.init_plot()
        ax.plot(self.row_sum_5min_GW)
        ax.set_ylim((0.0))
        ax.set_xlabel("Date/Time")
        ax.set_ylabel("GW")
        ax.set_title(self.title)
        fig.show()


@dataclass
class PlotManagerGW5min(PlotManager):
    """PlotManagerGW5min is a class that manages a collection of plots."""

    def plotGW5min(self, row_sum_5min_GW: ndarray, title: str) -> None:
        """Create PlotGW5min to later show a plot of five minute interval GW data.
        All the plots are stored in LIST_OF_PLOTS and will be shown when do_plots() is called.
        """
        self.plots.append(
            PlotGW5min(plot_manager=self, row_sum_5min_GW=row_sum_5min_GW, title=title)
        )

    def do_plots(self) -> None:
        """Show all the plots stored in LIST_OF_PLOTS."""
        for plot in self.plots:
            plot.plot_data()


def read_duid_categories_csv(file_path):
    """Reads the DUID categories into
    cat_list - a list of categories found in the CSV file
    duid_catidx - a dictionary of DUID names to indices into
    cat_list
    """

    cat_list = []  # Create the list of categories.
    duid_catidx = dict()
    with open(file_path, mode="r", newline="") as file:
        csv_reader = reader(file)
        next(csv_reader)  # Skip the header row: "DUID" "Cat" ...
        for row in csv_reader:
            # Read in the duid and its category
            o = row
            duid, category, batt_load = o[0], o[1], o[2]
            try:
                # look up the category index
                cat_index = cat_list.index(category)
            except ValueError:  # new category
                cat_index = len(cat_list)
                cat_list.append(category)
            duid_catidx[duid] = (cat_index, batt_load == "y")
    return cat_list, duid_catidx


def expand_env(s):
    """Expand a beginning environment variable. e.g. %APPDATA%."""

    if (s is not None) and s[0] == "%":
        a = s.split("%")
        s = "".join((os_getenv(a[1]), a[2]))
    return s


def main(
    mirror_dir=None,
    data_dir=Path(user_data_dir("nemweb", appauthor="NEMWeb")),
    output_dir=Path(user_documents_dir()),
):
    """The NEMWeb.py program.

    Args:
        mirror_dir: Optional directory containing local copies of ZIP files
        data_dir: Directory for storing PKL files (defaults to platform-specific app data)
        output_dir: Directory for output files (defaults to platform-specific documents)
    """
    # Convert mirror_dir to Path if provided
    if mirror_dir:
        mirror_dir = Path(mirror_dir)

    nemweb = NEMWeb(mirror_dir, str(data_dir) + os.path.sep, str(output_dir) + os.path.sep)

    number_of_files = len(nemweb.dispatch_scada.nem_MW_df_list)
    print(f"Assembling 12 months of data from {number_of_files} files...")
    nemweb.df_five_minute = nemweb.build_df_five_minute()
    nemweb.show_elapsed()

    # ds = nemweb.dispatch_scada
    df_five_minute = nemweb.df_five_minute

    from_month = df_five_minute.axes[0][0].strftime("%b %Y")
    to_month = (df_five_minute.axes[0][-1] - timedelta(days=1)).strftime("%b %Y")
    print("Showing interactive graph of raw dispatch_SCADA data. Close graph to continue.")
    row_sum_5min_GW = df_five_minute.sum(1) / 1000
    pmgr = PlotManagerGW5min()
    pmgr.plotGW5min(row_sum_5min_GW, f"Raw Dispatch_SCADA data {from_month} - {to_month}")

    print("Finding and interpolating any short period of missing Dispatch_SCADA data.")
    repair_missing_periods(df_five_minute, row_sum_5min_GW)

    print("Showing interactive graph of repaired dispatch_SCADA data. Close graph to continue.")
    row_sum_5min_GW = df_five_minute.sum(1) / 1000
    pmgr.plotGW5min(row_sum_5min_GW, f"Repaired Dispatch_SCADA data {from_month} - {to_month}")

    nemweb.show_elapsed()
    print("Adding rooftop solar data...")
    start_date = nemweb.df_five_minute.axes[0][0].replace(minute=0)
    rt = nemweb.rooftop_pv
    rt.df_thirty_minute = rt.build_df_thirty_minute(start_date)
    nemweb.add_rooftop_solar_data()
    nemweb.show_elapsed()

    total_num_data_points = (
        nemweb.dispatch_scada.num_data_points + nemweb.rooftop_pv.num_data_points
    )
    print(f" Total number of nemweb.com.au data points: {total_num_data_points:,}")

    print("Showing interactive graph of NEM 5-minute generation. Close graph to continue.")
    row_sum_5min_GW = df_five_minute.sum(1) / 1000
    pmgr.plotGW5min(row_sum_5min_GW, f"NEM Generation {from_month} - {to_month}")

    # Read in 'DUID Categories.csv' from the same directory as this program.
    nemweb.duid_categories_file = Path(__file__).parent / "DUID categories.csv"
    print(f"Reading {nemweb.duid_categories_file}")
    o = read_duid_categories_csv(nemweb.duid_categories_file)
    nemweb.cat_list = o[0]
    nemweb.duid_catidx = o[1]
    print(f"{nemweb.cat_list}\n")

    number_of_duids = len(nemweb.df_five_minute.axes[1]) - 5
    print(f"Tallying categories for {number_of_duids} DUIDs and 5 rooftop solar regions...")
    o = nemweb.build_cat_and_in_out_dfs()
    nemweb.df_categories = o[0]
    nemweb.df_in_out = o[1]
    nemweb.show_elapsed()

    nemweb.save_in_out_and_cat()

    # Show the plots and pause on close.
    pmgr.do_plots()
    pmgr.pause_on_close(
        5
    )  # Wait 5 seconds before closing the plots unless interacts with the plots.
    return nemweb


if __name__ == "__main__":
    if True:
        nemweb = main()
    else:
        mirror_dir = (
            r"C:\Users\Public\Downloads\www.nemweb.com.au"
            "\\REPORTS\\ARCHIVE\\"
        )
        nemweb = main(mirror_dir=mirror_dir)
