# Architectural Unit Conversion

This function will translate strings such as 10'-6" into architectural units as a double data type in feet like 10.5000000000000.

If there is no ' or " in the string, the default is inches.

Accepted entries are:

```
10'6"
10'6
10'-6"
10'-6
10'-6 1/2"
10'-6 1/2
10'
6
6-1/2
6-1/2"
6 1/2
6 1/2"
1/2
1/2"
10.5'
6.5"
6.5
```
