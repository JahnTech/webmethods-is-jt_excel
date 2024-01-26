package com.jahntech.webm.is.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.wm.data.IData;
import com.wm.data.IDataCursor;
import com.wm.data.IDataFactory;
import com.wm.data.IDataUtil;

/**
 * Convert DocumentList ({@link IData}) into a range of cells in a spreadsheet
 */
public class IDataToCellConverter {

	private Sheet sheet;
	private int columnStart;
	private int columnEnd = JtCell.NO_VALUE_PROVIDED_FOR_POSITION;
	private int rowStart;
	private int rowEnd = JtCell.NO_VALUE_PROVIDED_FOR_POSITION;
	private boolean isFirstRowAsHeader;

	private CellStyle styleHeader = null;
	private CellStyle styleData = null;
	private CellStyle styleDataAlternate = null;

	/**
	 * Initialize with spreadsheet into which the Document List shall be taken
	 * 
	 * @param sheet              Sheet from spreadsheet workbook
	 * @param columnStart        Number of start column (index starts with 0)
	 * @param rowStart           Number of start column (index starts with 0)
	 * @param isFirstRowAsHeader Should the field names from the Document List be
	 *                           used to insert a header into the spreadsheet?
	 */
	public IDataToCellConverter(Sheet sheet, int columnStart, int rowStart, boolean isFirstRowAsHeader) {
		super();
		this.sheet = sheet;
		this.columnStart = columnStart;
		this.rowStart = rowStart;
		this.isFirstRowAsHeader = isFirstRowAsHeader;

	}


	/**
	 * Set style for header line
	 * 
	 * @param styleAsObject Cell style for the header line
	 */
	public void setStyleHeader(Object styleAsObject) {
		this.styleHeader = getStyleFromObject(styleAsObject);
	}


	/**
	 * Set style for data lines
	 * 
	 * @param styleAsObject Cell style for data lines
	 */
	public void setStyleData(Object styleAsObject) {
		this.styleData = getStyleFromObject(styleAsObject);
	}


	/**
	 * Optional alternating style for data lines to improve readability
	 * 
	 * @param styleAsObject Cell style for alternating data lines
	 */
	public void setStyleDataAlternate(Object styleAsObject) {
		this.styleDataAlternate = getStyleFromObject(styleAsObject);
	}


	/**
	 * Get cell style from object if not null and add it to the workbook
	 * 
	 * @param styleAsObject Style as provided from the calling Flow service
	 * @return
	 */
	private CellStyle getStyleFromObject(Object styleAsObject) {
		if (styleAsObject != null) {
			CellStyle style = (CellStyle) styleAsObject;
			CellStyle localStyle = sheet.getWorkbook().createCellStyle();
			localStyle.cloneStyleFrom(style);
			return localStyle;
		}
		return null;
	}


	/**
	 * Get row object for row number. If the row does not exist yet, it will be
	 * created
	 * 
	 * @param rowNumber Number of row to get
	 * @return row for row number
	 */
	private Row getValidRow(int rowNumber) {
		Row row;
		row = sheet.getRow(rowNumber);

		try {
			Cell cell = row.getCell(0);
		} catch (Exception e) {
			row = sheet.createRow(rowNumber);
		}

		return row;
	}


	/**
	 * Set cell range from Document List
	 * 
	 * @param documentList Document List with data for the spreadsheet
	 */
	public void setFromDocumentList(IData[] documentList) {
		Cell cell;
		Row row;

		int i = 0;
		int rowNumber = rowStart;
		int columnNumber = columnStart;

		if (documentList != null) {

			for (i = 0; i < documentList.length; i++) {

				IDataCursor docCursor = documentList[i].getCursor();
				row = getValidRow(rowNumber);
				columnNumber = columnStart;

				if (isFirstRowAsHeader && (rowNumber == rowStart)) {
					while (docCursor.next()) {
						String docListFieldName = docCursor.getKey();
						cell = row.createCell(columnNumber);
						cell.setCellValue(docListFieldName);
						if (styleHeader != null) {
							cell.setCellStyle(styleHeader);
						}
						columnNumber++;
					}
					rowNumber++;
					columnNumber = columnStart;
					row = getValidRow(rowNumber);
					docCursor = documentList[i].getCursor();
				}

				// Write data
				while (docCursor.next()) {
					String docListFieldValue = docCursor.getValue().toString();
					cell = row.createCell(columnNumber);
					cell.setCellValue(docListFieldValue);
					if (styleData != null) {
						if (i % 2 == 0 || styleDataAlternate == null) {
							// Even row or no alternate style defined
							cell.setCellStyle(styleData);
						} else {
							// Uneven row and alternate style defined
							cell.setCellStyle(styleDataAlternate);
						}
					}
					columnNumber++;
				}
				docCursor.destroy();
				rowNumber++;
			}
		}

	}


	/**
	 * Return first column containing data.
	 * 
	 * @return first column containing data
	 */
	public int getColumnStart() {
		return columnStart;
	}


	/**
	 * Return last column containing data. This has been calculated when the data
	 * was inserted into the spreadsheet by {@link #setFromDocumentList(IData[])}.
	 * If no data have been inserted yet, it will return
	 * {@value JtCell#NO_VALUE_PROVIDED_FOR_POSITION}.
	 * 
	 * @return last column containing data
	 */
	public int getColumnEnd() {
		return columnEnd;
	}


	/**
	 * Return first row containing data.
	 * 
	 * @return first row containing data
	 */
	public int getRowStart() {
		return rowStart;
	}


	/**
	 * Return last row containing data. This has been calculated when the data was
	 * inserted into the spreadsheet by {@link #setFromDocumentList(IData[])}. If no
	 * data have been inserted yet, it will return
	 * {@value JtCell#NO_VALUE_PROVIDED_FOR_POSITION}.
	 * 
	 * @return last row containing data
	 */
	public int getRowEnd() {
		return rowEnd;
	}

}
