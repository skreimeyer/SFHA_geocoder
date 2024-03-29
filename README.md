# SFHA Geocoder

This is a tool I used to support my work in a municipal public works department.

Please note that this program uses a geocoding service specific to central 
Arkansas; however, esri server REST interfaces are sometimes very similar.

---

Takes an excel file containing addresses and attempts to geocode all rows with 
the PAGIS geocoding service. Addresses must have a valid pattern to be coded:

    [Number] [Street Name] ... [Suffix]

or:

    Lot [Number] { Block [Number] } [Subdivision Name] ...
    Block [Number] [Subdivision Name] ...

Rows prefixed with "Lot" or "Block" will result in a query to the parcel server.
The address string will be parsed into lot and block numbers and subdivision
name. The first result will be assumed valid, and the parcel will be geocoded
by calculating the parcel centroid.

Address fields which match neither pattern will be skipped. All data will be
copied to a new spreadsheet named ORIGINALFILE_geocoded.xls
