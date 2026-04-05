# dyno-jet-log-table-generator-app
This is a quick application I designed to read data from a DynoJet tune log file attached to a 2015 Harley Davidson FHLTCUL with upgraded cams, exhaust, and air filter.  

This app allows you to pigeon-hole the data collected from a log session onto an X Y axis for modifying/creating .pvt tune files for the DynoJet ECM.  

For example, in the log containing 70,000+ rows collected from a 20-minute live tuning session on the open road, you have a row that displays data points for RPM and MAP (manifold ambient pressue) that read 1389 RPM and 43.86 MAP readings but you do not have a table cell in a tune table for those two exact data points collected; you have to decide if that data point gets mapped to 1250 or 1500 RPM and 40 or 45 MAP.
