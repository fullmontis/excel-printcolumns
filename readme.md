Print-Columns - An excel macro for printing in columns
======================================================

## Introduction

Print-Columns allows to and quickly easily fit sheets with a high
number of rows and low number of columns on a single page. This helps
show more data on each papersheet and save paper.

## Installation

Simply copy the code above in a new macro in excel.

## How it works

![how this works](http://i.imgur.com/esGyERs.png)
     
Divides selected cells in blocks of height h, then creates a
new sheet. It then puts each block side by side, up to
PAGECOLS blocks. This is one page.
It then proceeds to create a new page under the first one.
   
Both h and pagecols are set by the user when run.
   
It also automatically sets page width to one page and
adds page breaks after eery new page, for a print friendly format.
 
Important: all selected cells should have the same height to avoid
weird impagination

## License   

This software is placed in the public domain by its creator.   

