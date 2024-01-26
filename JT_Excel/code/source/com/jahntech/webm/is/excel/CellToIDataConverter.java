package com.jahntech.webm.is.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.wm.data.IData;
import com.wm.data.IDataCursor;
import com.wm.data.IDataFactory;
import com.wm.data.IDataUtil;

/**
 * Convert range of cells from a spreadsheet into a DocumentList ({@link IData})
 */
public class CellToIDataConverter {

	private Sheet sheet;
	private int columnStart;
	private int columnEnd;
	private int rowStart;
	private int rowEnd;
	private boolean isFirstRowAsHeader;
	private int firstDataRow;
	private int numberOfRows;
	private String[] headers;

	/**
	 * Initialize with spreadsheet from which the values shall be taken
	 * 
	 * @param sheet              Sheet from spreadsheet workbook
	 * @param columnStart        Number of start column (index starts with 0). To
	 *                           use auto-detection (using
	 *                           {@link Sheet#getFirstCellNum()}) provide the value
	 *                           {@value JtCell#NO_VALUE_PROVIDED_FOR_POSITION}
	 * @param columnEnd          Number of end column (index starts with 0). To use
	 *                           auto-detection (using
	 *                           {@link Sheet#getLastCellNum()}) provide the value
	 *                           {@value JtCell#NO_VALUE_PROVIDED_FOR_POSITION}
	 * @param rowStartNumber     Number of start column (index starts with 0). To use
	 *                           auto-detection (using
	 *                           {@link Sheet#getFirstRowNum()}) provide the value
	 *                           {@value JtCell#NO_VALUE_PROVIDED_FOR_POSITION}
	 * @param rowEnd             Number of start column (index starts with 0). To
	 *                           use auto-detection (using
	 *                           {@link Sheet#getLastRowNum()}) provide the value
	 *                           {@value JtCell#NO_VALUE_PROVIDED_FOR_POSITION}
	 * @param isFirstRowAsHeader Should the values of the first row be used as field
	 *                           names for the output? If no, the names will be
	 *                           auto-generated starting with
	 *                           {@value HeaderRow#GENERIC_HEADER_BASE} and directly
	 *                           followed by an increasing number. If a cell that is
	 *                           supposed to contain the field name is empty, the
	 *                           fall-back is to apply the logic of the
	 *                           auto-generation.
	 */
	public CellToIDataConverter(Sheet sheet, int columnStart, int columnEnd, int rowStart, int rowEnd,
			boolean isFirstRowAsHeader) {
		super();
		this.sheet = sheet;
		this.columnStart = columnStart;
		this.columnEnd = columnEnd;
		this.rowStart = rowStart;
		this.rowEnd = rowEnd;
		this.isFirstRowAsHeader = isFirstRowAsHeader;

		// Auto-detect start/end values, if necessary
		if (rowStart == JtCell.NO_VALUE_PROVIDED_FOR_POSITION) {
			rowStart = sheet.getFirstRowNum();
		}
		if (rowEnd == JtCell.NO_VALUE_PROVIDED_FOR_POSITION) {
			rowEnd = sheet.getLastRowNum();
		}
		if (columnStart == JtCell.NO_VALUE_PROVIDED_FOR_POSITION) {
			columnStart = sheet.getRow(rowStart).getFirstCellNum();
		}
		if (columnEnd == JtCell.NO_VALUE_PROVIDED_FOR_POSITION) {
			columnEnd = sheet.getRow(rowEnd).getLastCellNum() - 1;
		}

		numberOfRows = rowEnd - rowStart + 1;
		firstDataRow = rowStart;

		// If containing the headers, the first row must be "removed" from data
		// processing
		if (isFirstRowAsHeader) {
			firstDataRow++;
			numberOfRows--;
		}
		headers = determineHeaders();
	}


	/**
	 * Build header names depending how the instance of {@link CellToIDataConverter}
	 * was initialized. So happens by either reading them from the first row of the
	 * specified (or auto-detected) cell range. Or by auto-generating using the
	 * fixed value {@value HeaderRow#GENERIC_HEADER_BASE} and appending the number
	 * of column.
	 * 
	 * @return header names
	 */
	private String[] determineHeaders() {
		String headers[];
		if (isFirstRowAsHeader) {
			Row firstRow = sheet.getRow(rowStart);
			HeaderRow headerRow = new HeaderRow(firstRow, columnStart, columnEnd);
			headers = headerRow.getFieldNames();
		} else {
			int numberOfColumns = columnEnd - columnStart + 1;
			headers = HeaderRow.genericHeaders(numberOfColumns);
		}
		return headers;
	}


	/**
	 * Return cell range as Document List (aka {@link IData})
	 * 
	 * @return document list with cell contents
	 */
	public IData[] getAsDocumentList() {

		IData[] out = new IData[numberOfRows];
		int indexDocumentList = 0;

		IDataCursor rowCursor;

		for (int currentRowNum = firstDataRow; currentRowNum <= rowEnd; currentRowNum++) {

			out[indexDocumentList] = IDataFactory.create();
			rowCursor = out[indexDocumentList].getCursor();

			int indexField = 0;

			for (int currentColumnNum = columnStart; currentColumnNum <= columnEnd; currentColumnNum++) {

				Row currentRowContent = sheet.getRow(currentRowNum);

				// If cell has content ...
				if (currentRowContent != null && currentRowContent.getCell(currentColumnNum) != null) {

					// ... get cell value as string
					Cell cell = sheet.getRow(currentRowNum).getCell(currentColumnNum);
					String cellValue = new JtCell(cell).getValueAsString();

					IDataUtil.put(rowCursor, headers[indexField], cellValue);
				} else {

					// Otherwise fill with empty string
					IDataUtil.put(rowCursor, headers[indexField], "");
				}
				indexField++;
			}
			indexDocumentList++;

			if (rowCursor != null) {
				rowCursor.destroy();
			}
		}
		return out;
	}


	/**
	 * Return first column to read from. If no value was specified during instance
	 * creation, this will return the auto-detected value.
	 * 
	 * @return first column to read from
	 */
	public int getColumnStart() {
		return columnStart;
	}


	/**
	 * Return last column to read from. If no value was specified during instance
	 * creation, this will return the auto-detected value.
	 * 
	 * @return last column to read from
	 */
	public int getColumnEnd() {
		return columnEnd;
	}


	/**
	 * Return first row to read from. If no value was specified during instance
	 * creation, this will return the auto-detected value.
	 * 
	 * @return first row to read from
	 */
	public int getRowStart() {
		return rowStart;
	}


	/**
	 * Return last row to read from. If no value was specified during instance
	 * creation, this will return the auto-detected value.
	 * 
	 * @return last row to read from
	 */
	public int getRowEnd() {
		return rowEnd;
	}

}
