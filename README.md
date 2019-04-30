QGIS Excel Sync
===============

This is a QGIS plugin that links an Excel file to a Shapefile (or other layer).

A 1:1 relation is created between the two files, for every new row on the layer
a new row will be created in the Excel file. An autoincrementing identification
key will be used to link the individual rows. Derived fields can be configured
so the Excel file can contain information like area, centroid, wkt representation
or other information about the linked feature.
