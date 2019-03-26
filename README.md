# Save all sheets in separate files

Macro to save all the sheets from a spreadsheet in separate files. Filename is automatically generated from sheet name.

Macro has the following customisation:

- Output format (xls, xlsx, ods, csv)
- Location of the output folder
- Output filename can be further customised

## Notes

Path separator is / not \ (even on WIndows)!

## Csv options:

To save in .csv takes an extra parameter: *FilterOptions*

An example is shown below with some options explained:

<pre>
args2(2).Value = "44,34,76,1,,0,false,true,false,false,false"
                   ^ ^  ^  ^  ^  ^     ^    ^    ^     ^-: ?
                   | |  |  |  |  |     |    |    |_______: Save cell formula instead of calculated value, if true
                   | |  |  |  |  |     |    |____________: Save cell content as shown, if true
                   | |  |  |  |  |     |_________________: ?
                   | |  |  |  |  |_______________________: ?
                   | |  |  |  |__________________________: Quote all text cells, if true
                   | |  |  |_____________________________: ?
                   | |  |________________________________: Encoding: 76: utf-8; ANSI: Windows 1252/WInLatin 1; 65535: utf-16
                   | |___________________________________: string delimiter: 34: "; 39: '
                   |_____________________________________: separator: 9: tab; 44: coma; FIX: fixed column width
</pre>

## Licence

This macro is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This macro is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

Copy of the GNU General Public License is availabel at: [http://www.gnu.org/licenses/]()

©2019 - GeoProc.com
