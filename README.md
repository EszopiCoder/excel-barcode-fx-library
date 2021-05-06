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
- All check digits are auto-calculated.

| Function Name | Data Types | Length | Author |
| --- | --- | --- | --- |
| `Aztec()` | ASCII | Unlimited | [Alois Zingl](http://members.chello.at/~easyfilter/barcode.html) |
| `Code11()` | Numeric and dash | Unlimited | Eszopicoder |
| `Code39()` | Alphanumeric and -$%./+ | Unlimited | Eszopicoder |
| `Code93()` | Alphanumeric and -$%./+ | Unlimited + 2 check digits | Eszopicoder |
| `Code128()` | ASCII | Unlimited | [Alois Zingl](http://members.chello.at/~easyfilter/barcode.html) |
| `DataMatrix()` | ASCII | Unlimited | [Alois Zingl](http://members.chello.at/~easyfilter/barcode.html) |
| `EAN_2()` | Numeric | 2 digits | Eszopicoder |
| `EAN_5()` | Numeric | 5 digits | Eszopicoder |
| `EAN_13()` | Numeric | 12 digits + check digit | Eszopicoder |
| `PDF_417()` | ASCII | Unlimited | [Grandzebu](http://grandzebu.net/informatique/codbar-en/pdf417.htm) |
| `QRCode()` | ASCII | Unlimited | [Alois Zingl](http://members.chello.at/~easyfilter/barcode.html) |
| `UPCA()` | Numeric | 11 + check digit (UPC-A) or 8 digits (EAN-8) | Eszopicoder |
| `UPCE()` | Numeric | ("0" or "1") + 6 digits | Eszopicoder |
## Sample Images
<img src="Images/Barcode Sample.PNG">
