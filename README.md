# unMergeExcelCell.py - unmerge excel file

Using the python package [xlrd](https://github.com/python-excel/xlrd) and [xlwt](https://github.com/python-excel/xlwt) to unmerge the excel file, and then auto fill the merged cells by the first cell in the merged cells.

For general purpose, unMergeExcelCell.py won't keep the style and format from the origin excel file.

## Usage
``` 
    Un-merge excel cell and auto fill with the first cell value in the merged cells
    Usage: unMergeExcelCell.py <excel file>
    excel file: the path for un-merging excel file
```


> Notice: unMergeExcelCell.py will make a new excel file by add the sufix `_unmerged`.

> `e.g. excel.xlsx -> excel_unmerged.xlsx`

**In learning we trust !**
