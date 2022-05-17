# apsu-magnet-module-assembly-notebooks
This is a Project using Jupyter Notebooks to Support APSU Module Assembly

This project provides one or more Jupyter notebooks to facilitate Module Assembly for the APS Upgrade.
These notebooks requires a number of packages that are installed via conda.

# Notebooks provided in this package are:

## MagnetShimming/shim_calculation.ipynb
This directory contains two notebooks that perform
a constrained optimization to calculate shim values for the magnets on a module to
place the magnet centers on a best fit line.   The first notebook (01-...)
is used to determine the initial magnet shim values and the second
notebook (02-...) is used for subsequent refinements.

The first notebook (01-...) performs an initial calculation to determine
the starting shim pack for module assembly.   Output files are stored as

File name format is:

(Module Name)_(X/Y)_module_shim_pack_00_(DateStamp).xlsx

For Example:
DLMA-1160_X_module_shim_pack_00_202111022-113730.xlsx


The second notebook (02-...) performs subsequent calculations to optmize
the shim configuration.  This can be run multiple times where and an internal
index is incremented to help keep track of the sequence of the calculations.

File Format is:

(Module Name)_(X/Y)_module_shim_pack_(Index)_(DateStamp).xlsx

For Example:
DLMA-1160_X_module_shim_change_01_202111022-114530.xlsx


## ModuleSurvey/MagnetCoordinates.ipynb

This notebook performs data manipulation to support Survey Group activities.

# Instructions for installing Jupyter and required packages

1) First, install miniconda:
https://docs.conda.io/en/latest/miniconda.html

Once you have this installed, you will have the ability to manage separate environments and versions of python
for specific applications.   Miniconda doesn't install a large number of packages so it is easier to tailor
the environment to a particular applications needs.    However, it does mean that you need to install packages
manually.

2) Bring up a shell out of which you can run conda.  This will depend on your operating system.
On Macs and Linux, you do this through your shell.  On Windows, you should have a menu item in
your applications sidebar to bring up a conda shell.  Once you have this set up, navigate to the magnet shim directory and issue the following commands.

> cd W:\apsu-magnet-module-assembly-notebooks (replace W with shared drive designation on your machine)
> conda create --yes --name SharedMagnetShim python=3.9.6
> conda activate SharedMagnetShim
> conda install --yes jupyterlab
> jupyter lab

Open Installation_Script.ipynb and run the first cell (Shift + Enter). 

Close JupyterLab (Cntrl + C) and open anaconda prompt again
## Running the notebooks

Once you have a working conda environment for the notebooks, you bring up your conda shell and issue
the following

> cd W:\apsu-magnet-module-assembly-notebooks\MagnetShimming (Replace W with shared drive designation on your machine)
> conda activate SharedMagnetShim
> jupyter lab   

You can now run the notebooks out of your web browser.   One way is to use the menu [Run]:[Restart Kernal and Run All Cells]

If you don't want to see the little bit of python code, you can use [View]:[Collapse All Code].

At the bottom of your screen,  you will see a run status.  This will either be Idle or Busy depending on whether or not
the script is running

## A Note on Data Files

There are some sample data files stored in the directory CDBInfo and sample Excel files in the main directory.   These are useful
for debugging.   The proper files, when required,  should be put in as input to the program.   When this is run on a network
where the CDB can be reached,  the CDBInfo files will be overwritten by the data from the CDB.
