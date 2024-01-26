package com.jahntech.webm.is.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * Header row in a spreadsheet
 */
public class HeaderRow {

	/**
	 * Will be used to name headers that have not individual name. That can either
	 * happen if no header names were provided at all. But it also happens if cells
	 * in a row, that is supposed to contain the header names, are empty
	 */
	public static final String GENERIC_HEADER_BASE = "Field";

	private Row row;
	private int columnStart;
	private int columnEnd;

	/**
	 * Initialize header row
	 * 
	 * @param row         Spreadsheet row to contain header names
	 * @param columnStart First column to use for header names
	 * @param columnEnd   Last column to use for header names
	 */
	public HeaderRow(Row row, int columnStart, int columnEnd) {
		this.row = row;
		this.columnStart = columnStart;
		this.columnEnd = columnEnd;
	}


	/**
	 * Extract values out of row. Empty cells will be represented by
	 * {@value #GENERIC_HEADER_BASE} plus the column number. If multiple column in
	 * the provided range have the same content, an underscore followed by an ever
	 * increasing number will be appended.
	 * 
	 * @return array with header names
	 */
	public String[] getFieldNames() {

		String[] headers = new String[columnEnd - columnStart + 1];

		for (int column = 0; columnStart + column <= columnEnd; column++) {

			Cell cell = row.getCell(column);

			if (cell == null) {
				headers[column] = GENERIC_HEADER_BASE + (column + 1);
			} else {

				// Duplicate check for column names
				String fieldName = new JtCell(cell).getValueAsString();
				boolean duplicateCheck = true;
				int duplicateSuffix = 0;

				while (duplicateCheck) {
					duplicateSuffix++;
					duplicateCheck = false;
					for (int i = 0; i < column; i++) {
						if (headers[i].equals(fieldName)) {
							fieldName = fieldName.concat("_" + duplicateSuffix);
							duplicateCheck = true;
						}
					}
				}
				headers[column] = fieldName;
			}
		}
		return headers;
	}


	/**
	 * Create array with generic header names for a given number of columns. The
	 * values will be {@value #GENERIC_HEADER_BASE} plus the number of the column
	 * (starting at 1)
	 * 
	 * @param numberOfColumns Number of column for which to create the header names
	 * @return array with headers
	 */
	public static String[] genericHeaders(int numberOfColumns) {
		String[] headers = new String[numberOfColumns];
		for (int i = 0; i < headers.length; i++) {
			headers[i] = GENERIC_HEADER_BASE + (i + 1);
		}
		return headers;
	}

}
