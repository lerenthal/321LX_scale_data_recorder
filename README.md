# 321LX_scale_data_recorder
the application connects to Precisa or any other USB/RS232 scales, records weight data on the screen and exports data to a CSV file. the app is python based and is packed with PyInstaller as single exe file for windows.
the gui is based on the TInkter library and serial library for cennectivity. 
the application supports direct SMTP based mail export of csv.
for compliance purposes the application has a crash recovery backend file to make sure no data is lost if the system crashes. 
testing was based on the Precuisa scale but the system is designed to support any USB/RS-232 or RJ-45 scale. 
