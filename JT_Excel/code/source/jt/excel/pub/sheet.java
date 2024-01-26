package jt.excel.pub;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import org.apache.poi.ss.usermodel.Workbook;
// --- <<IS-END-IMPORTS>> ---

public final class sheet

{
	// ---( internal utility methods )---

	final static sheet _instance = new sheet();

	static sheet _newInstance() { return new sheet(); }

	static sheet _cast(Object o) { return (sheet)o; }

	// ---( server methods )---




	public static final void getByIndex (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(getByIndex)>> ---
		// @sigtype java 3.5
		// [i] object:0:required workbook
		// [i] field:0:required index
		// [o] object:0:required sheet
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
			Workbook workbook = (Workbook) IDataUtil.get( pipelineCursor, "workbook" );
			int	index = IDataUtil.getInt( pipelineCursor, "index", -1 );
			IDataUtil.put( pipelineCursor, "sheet", workbook.getSheetAt(index) );
		} finally {
			pipelineCursor.destroy();
		}
			
		// --- <<IS-END>> ---

                
	}



	public static final void getByName (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(getByName)>> ---
		// @sigtype java 3.5
		// [i] object:0:required workbook
		// [i] field:0:required name
		// [o] object:0:required sheet
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
			Workbook workbook = (Workbook) IDataUtil.get( pipelineCursor, "workbook" );
			String name = IDataUtil.getString( pipelineCursor, "name" );
			IDataUtil.put( pipelineCursor, "sheet", workbook.getSheet(name) );
		} finally {
			pipelineCursor.destroy();
		}
			
		// --- <<IS-END>> ---

                
	}
}

