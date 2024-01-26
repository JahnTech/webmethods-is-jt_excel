package com.jahntech.webm.is.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * Wrapper around {@link Cell}
 *
 */
public class JtCell {

	/**
	 * Indicates that a coordinate (column or row) of a cell has not been specified.
	 * If this values is found, the appropriate fall-back logic kicks in. How the
	 * latter looks like varies greatly and depends on the place in the code.
	 */
	public static final int NO_VALUE_PROVIDED_FOR_POSITION = -1;

	private Cell cell;

	/**
	 * Initialize with cell
	 * 
	 * @param cell standard POI cell object
	 */
	public JtCell(Cell cell) {
		super();
		this.cell = cell;
	}


	/**
	 * Get cell value as string. This performs the necessary conversions, if
	 * applicable
	 * 
	 * @return cell value as string
	 */
	public String getValueAsString() {
		String value = null;

		switch (cell.getCellType()) {
		case STRING:
			value = cell.getRichStringCellValue().getString();
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				// System.out.println(cell.getDateCellValue());
				value = cell.getDateCellValue().toString();
			} else {
				value = Double.toString(cell.getNumericCellValue());
			}
			break;
		case BOOLEAN:
			value = Boolean.toString(cell.getBooleanCellValue());
			break;
		case FORMULA:
			value = cell.getCellFormula();
			break;
		default:
			value = "";
		}

		return value;
	}

}
