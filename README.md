# JT_Excel

This package allows webMethods Integration Server to create, read, update,
and write Excel spreadsheets.

## Functionality

Built on top of [Apache POI](https://poi.apache.org/) the `JT_Excel`
package allows you to work with workbooks and sheets in `.xls` and
`.xlsx` format. You can open or create workbooks and sheets. And within
those sheets it is possible to read, update, insert, and delete data.

Typical use-cases are to create spreadsheets with data coming from
backend systems like ERP or CRM. Another popular scenario is to
employ a spreadsheet for bulk-entry of data. Together with validations
that makes for a very efficient interface for power users.
Or you could use the spreadsheets to exchange data with customers.
Individual pricelists come to mind here.

## Limitations

In terms of data formats no formats other than the spreadsheet files
in `.xls` and `.xlsx` format (no CSV etc.) and the normal data
structures of Integration Server (`IData`) are in the game.

So if you want to convert an Excel spreadsheet to JSON or XML, you will need
to write a Flow service that parses the data from the spreadsheet, taking into account in which column and row values are located.

Typically you will have a document list after that with one document
per line and the columns mapped to the fields within that document.
From there you then work just as if the data had come from a database
operation, an EDI document, or any other data source.

As to what formats and content details are supported on the Excel side,
this is determined entirely by what Apache POI covers.

## Installation

You can install `JT_Excel` in two ways.

- There will be releases that come as a `ZIP` file and must be
  install in the traditional way. That means copying it into
  `$IS_HOME/replicate/inbound` and then invoking
  `Package Management / Install Inbound Release`.
- For people who want to be on the bleeding edge, you can always
  just clone or download the Git repository into your workspace
  and then work with it like a developer. For any environment
  other than DEVELOPMENT this is not recommended.

## Samples

There is also a test package (`JT_ExcelTest`) that additionally
serves as a sample package. Please browse its services to get
see how things are done.

## Built-in services

The service that come with `JT_Excel` can roughly be grouped like this:

- `jt.excel.pub.workbook` : Workbook-related operations (read, write, create)
- `jt.excel.pub.sheet` : Sheet-related operations (get, update, insert data etc.)
- `jt.excel.pub.cell` : Single cell-related operations

For the time being there is no details documentation for the individual
services. For inquiries about those please
get in touch with [JahnTech](https://jahntech.com).
