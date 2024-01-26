package jt.excel.pub;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import com.jahntech.webm.is.excel.JtCell;
import org.apache.poi.ss.usermodel.Sheet;
import com.jahntech.webm.is.excel.CellToIDataConverter;
import com.jahntech.webm.is.excel.IDataToCellConverter;
// --- <<IS-END-IMPORTS>> ---

public final class data

{
	// ---( internal utility methods )---

	final static data _instance = new data();

	static data _newInstance() { return new data(); }

	static data _cast(Object o) { return (data)o; }

	// ---( server methods )---




	public static final void cellsToDocumentList (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(cellsToDocumentList)>> ---
		// @sigtype java 3.5
		// [i] object:0:required sheet
		// [i] field:0:optional columnStart
		// [i] field:0:optional columnEnd
		// [i] field:0:optional rowStart
		// [i] field:0:optional rowEnd
		// [i] field:0:optional isFirstRowAsHeader {"true","false"}
		// [o] record:1:required documentList
		// [o] field:0:required columnStart
		// [o] field:0:required columnEnd
		// [o] field:0:required rowStart
		// [o] field:0:required rowEnd
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
			Sheet	sheet = (Sheet) IDataUtil.get( pipelineCursor, "sheet" );
			int	columnStart = IDataUtil.getInt( pipelineCursor, "columnStart", JtCell.NO_VALUE_PROVIDED_FOR_POSITION );
			int	columnEnd = IDataUtil.getInt( pipelineCursor, "columnEnd", JtCell.NO_VALUE_PROVIDED_FOR_POSITION );
			int	rowStart = IDataUtil.getInt( pipelineCursor, "rowStart", JtCell.NO_VALUE_PROVIDED_FOR_POSITION );
			int	rowEnd = IDataUtil.getInt( pipelineCursor, "rowEnd", JtCell.NO_VALUE_PROVIDED_FOR_POSITION );
			boolean isFirstRowAsHeader = IDataUtil.getBoolean(pipelineCursor, "isFirstRowAsHeader", false);
			
			CellToIDataConverter converter = new CellToIDataConverter(sheet, columnStart, columnEnd, rowStart, rowEnd, isFirstRowAsHeader);
			
			IDataUtil.put( pipelineCursor, "documentList", converter.getAsDocumentList() );
			IDataUtil.put( pipelineCursor, "columnStart", "" + converter.getColumnStart() );
			IDataUtil.put( pipelineCursor, "columnEnd", "" + converter.getColumnEnd() );
			IDataUtil.put( pipelineCursor, "rowStart", "" + converter.getRowStart() );
			IDataUtil.put( pipelineCursor, "rowEnd", "" + converter.getRowEnd() );
		} finally {
			pipelineCursor.destroy();
		}
			
		// --- <<IS-END>> ---

                
	}



	public static final void documentListToCells (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(documentListToCells)>> ---
		// @sigtype java 3.5
		// [i] record:1:required documentList
		// [i] object:0:required sheet
		// [i] field:0:optional columnStart
		// [i] field:0:optional rowStart
		// [i] field:0:optional isFirstRowAsHeader {"true","false"}
		// [i] record:0:optional styles
		// [i] - object:0:optional header
		// [i] - object:0:optional data
		// [i] - object:0:optional dataAlternate
		// [o] object:0:required sheet
		// [o] field:0:required columnStart
		// [o] field:0:required columnEnd
		// [o] field:0:required rowStart
		// [o] field:0:required rowEnd
		 
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
		Sheet sheet = (Sheet) IDataUtil.get( pipelineCursor, "sheet");
		IData[]	documentList = IDataUtil.getIDataArray( pipelineCursor, "documentList");
		int	columnStart = IDataUtil.getInt( pipelineCursor, "columnStart", 0);
		int	rowStart = IDataUtil.getInt( pipelineCursor, "rowStart", 0);
		Boolean	isFirstRowAsHeader = IDataUtil.getBoolean( pipelineCursor, "isFirstRowAsHeader", true);
		
		// Styles
		IData style = IDataUtil.getIData(pipelineCursor, "style");
		Object	styleHeaderObj = null;
		Object	styleDataObj = null;
		Object	styleDataAlternateObj = null;
		if (style != null) {
			IDataCursor styleCursor = style.getCursor();
			styleHeaderObj = IDataUtil.get( pipelineCursor, "header" );
			styleDataObj = IDataUtil.get( pipelineCursor, "data" );
			styleDataAlternateObj = IDataUtil.get( pipelineCursor, "dataAlternate" );
			styleCursor.destroy();
		}
		
		// Initialize sheet incl. styles
		IDataToCellConverter converter = new IDataToCellConverter(sheet, columnStart, rowStart, isFirstRowAsHeader);
		converter.setStyleHeader(styleHeaderObj);
		converter.setStyleData(styleDataObj);
		converter.setStyleDataAlternate(styleDataAlternateObj);
		converter.setFromDocumentList(documentList);
		
		IDataUtil.put( pipelineCursor, "sheet", sheet );
		IDataUtil.put( pipelineCursor, "columnStart", "" + converter.getColumnStart());
		IDataUtil.put( pipelineCursor, "columnEnd", "" + converter.getColumnEnd());
		IDataUtil.put( pipelineCursor, "rowStart", "" + converter.getRowStart());
		IDataUtil.put( pipelineCursor, "rowEnd", "" + converter.getRowEnd());
		//		
		} finally {
			pipelineCursor.destroy();
		}
		// pipeline
			
		// --- <<IS-END>> ---

                
	}
}

