# Barcode Function Library Excel Add-In
## Purpose and Features
Custom function library to generate the following 1D and 2D barcodes using autoshapes:
| Barcode Type | Barcodes |
| --- | --- |
| 1D Code | Code 11, Code 39, Code 93, Code 128 |
| 1D UPC/EAN | EAN-2, EAN-5, EAN-8, EAN-13, UPC-A, UPC-E |
| 2D Barcodes | Aztec, Data Matrix, PDF417, QR Code |
## Compatibility
Microsoft Excel 2013+
## Installation
1. Download 'Barcode Fx Library Add-In 1.0.xlam'
2. Follow these [instructions](https://support.office.com/en-us/article/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) from Microsoft
## Usage
- The functions are located under the 'Formulas' tab under the 'Barcode Function Library' group (see images).

| Function Name | Data Types | Author |
| --- | --- | --- |
| `Aztec()` | ASCII | [Alois Zingl](http://members.chello.at/~easyfilter/barcode.html) |
| `Code11()` | Numeric and dash | Eszopicoder |
| `Code39()` | Alphanumeric and -$%./+ | Eszopicoder |
| `Code93()` | Alphanumeric and -$%./+ | Eszopicoder |
| `Code128()` | ASCII | [Alois Zingl](http://members.chello.at/~easyfilter/barcode.html) |
| `DataMatrix()` | ASCII | [Alois Zingl](http://members.chello.at/~easyfilter/barcode.html) |
| `EAN_2()` | Numeric | Eszopicoder |
| `EAN_5()` | Numeric | Eszopicoder |
| `EAN_13()` | Numeric | Eszopicoder |
| `PDF_417()` | ASCII | [Grandzebu](http://grandzebu.net/informatique/codbar-en/pdf417.htm) |
| `QRCode()` | ASCII | [Alois Zingl](http://members.chello.at/~easyfilter/barcode.html) |
| `UPCA()` | Numeric | Eszopicoder |
| `UPCE()` | Numeric | Eszopicoder |
## Sample Images
<img src="Images/Barcode Sample.PNG">
