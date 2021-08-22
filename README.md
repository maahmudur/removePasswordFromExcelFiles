# Remove Password From Excel Files with Known Password
##This script is intended to use before the extraction process if the source files comes with password protected.

The data collection team often collects raw data from the factories where the files are password-protected, however we can collect the password from them too. Since those files can not be read directly using pandas without dealing with errors in a Jupyter Notebook, it's convenient to make copies of those files after removing the password. The password is hardcoded in the script but can be asked to the user while running if necessary.


