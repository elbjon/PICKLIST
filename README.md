# PICKLIST
A digital pick list complementing a commercial Warehouse Management System (WMS). It has been in daily use for over a decade now.

## Situation

A vegetable delivery service has acquired an organic food store to expand its product range.
Products ordered by customers must now be picked from shop shelves for further processing.
The WMS initially provided only paper picklists, which did not meet basic requirements.
The commercial digital solution also failed, introducing disadvantages compared to the paper picklist. 
*Analog* picking from shop shelves, rather than from a warehouse, was time-consuming and prone to errors.
Additionally, the picking staff was not accustomed to using computers.

## Solution

- **Data Extraction:** Tapping into the WMS for data and preparing the picklist on a PC.
- **Picking Process:** Using mobile notebooks with barcode scanners attached to picking trolleys.
- **Post-Processing:** Updating the WMS data on a PC after the picking process.


Key Features:
Digital Picklist with Clear product information (storage location/shelf number, manufacturer, product name, product quantities, weight, etc. )
Packing is organized according to storage location and divided into delivery tours (two tours at a time) to optimize further processing.
Barcode scanning prevents mispicks.
Implemented as VBA for Excel Scripts mainly.

Issues:
Some extremely rare Barcodes (other than EAN) aren't processed reliably.
Some parts of the code had to be removed before publishing on github.

Future Enhancements:
Adding product images from WMS, ecoinform, or DataNatuRE.
Developing a standalone solution (not based on VBA).
Implementing augmented reality (e.g., Google Glasses) to highlight products in the picker’s view.



WMS: BioOffice
