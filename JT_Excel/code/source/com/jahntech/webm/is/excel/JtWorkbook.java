package com.jahntech.webm.is.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Wrapper around {@link Workbook} to handle the different types transparently
 * for the using Java services
 */
public class JtWorkbook {

	/**
	 * Possible types for workbooks
	 */
	public enum Type {
		/**
		 * OLE2-based spreadsheet
		 */
		XLS,
		/**
		 * Newer, XML-based spreadsheet
		 */
		XLSX
	}

	private Workbook workbook;
	private Type type;

	/**
	 * Initialize workbook from file, taking into account the different formats
	 * 
	 * @param file File object for workbook
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public JtWorkbook(File file) throws InvalidFormatException, IOException {
		type = getTypeFromFileName(file.getName());
		readFile(file);
	}


	/**
	 * 
	 * 
	 * @param file
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private void readFile(File file) throws InvalidFormatException, IOException {
		switch (type) {
		case XLS:
			FileInputStream fis = new FileInputStream(file);
			workbook = new HSSFWorkbook(fis);
			break;
		case XLSX:
			workbook = new XSSFWorkbook(file);
		default:
			break;
		}

	}


	/**
	 * Get workbook object
	 * 
	 * @return work book object
	 */
	public Workbook getWorkbook() {
		return workbook;
	}


	/**
	 * Determine the type of the workbook based on the file name
	 * 
	 * @param fileName File name
	 * @return type of workbook
	 */
	public static Type getTypeFromFileName(String fileName) {
		String fileNameSuffixUpperCase = FilenameUtils.getExtension(fileName).toUpperCase();

		if (fileNameSuffixUpperCase.equals(Type.XLSX.toString())) {
			return Type.XLSX;
		} else if (fileNameSuffixUpperCase.equals(Type.XLS.toString())) {
			return Type.XLS;
		} else {
			throw new IllegalStateException("Filename is not valid for an Excel spreadsheet");
		}

	}


	/**
	 * Determine type of workbook from string that matches the filename extension.
	 * Defaults to {@value Type#XLS} if null or empty.
	 * 
	 * @param typeStr Filename extension (not case-sensitive), may be null or empty
	 * @return type of workbook
	 */
	public static Type getType(String typeStr) {
		JtWorkbook.Type type;

		if (typeStr == null || typeStr.equals("")) {
			type = Type.XLS;
		} else {
			type = Type.valueOf(typeStr.toUpperCase());
		}

		return type;
	}

}
