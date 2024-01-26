package jt.excel.pub;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.jahntech.webm.is.excel.JtWorkbook;
import com.jahntech.webm.is.excel.JtWorkbook.Type;
// --- <<IS-END-IMPORTS>> ---

public final class workbook

{
	// ---( internal utility methods )---

	final static workbook _instance = new workbook();

	static workbook _newInstance() { return new workbook(); }

	static workbook _cast(Object o) { return (workbook)o; }

	// ---( server methods )---




	public static final void close (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(close)>> ---
		// @sigtype java 3.5
		// [i] object:0:required workbook
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
			Workbook workbook = (Workbook) IDataUtil.get( pipelineCursor, "workbook" );
			try {
				workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
				throw new ServiceException(e);
			}
			
			
		} finally {
			pipelineCursor.destroy();
		}
			
		// --- <<IS-END>> ---

                
	}



	public static final void create (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(create)>> ---
		// @sigtype java 3.5
		// [i] field:0:optional version {"xls","xlsx"}
		// [o] object:0:required workbook
		// pipeline 
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
			String	version = IDataUtil.getString( pipelineCursor, "version" );
			JtWorkbook.Type type = JtWorkbook.getType(version);
		
			Workbook workbook = null;
			
			try {
				
				if(type == Type.XLS){
					workbook = new HSSFWorkbook();
				} else {
					workbook = new XSSFWorkbook();
				}
			} catch (Exception e) {
				e.printStackTrace();
				throw new ServiceException(e);
			}
		
			IDataUtil.put( pipelineCursor, "workbook", workbook );
		} finally {
			pipelineCursor.destroy();
		}
			
		// --- <<IS-END>> ---

                
	}



	public static final void open (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(open)>> ---
		// @sigtype java 3.5
		// [i] field:0:required filePath
		// [o] object:0:required workbook
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
			String	filePath = IDataUtil.getString( pipelineCursor, "filePath" );
		
			JtWorkbook jtwb = null;
			try {
				jtwb = new JtWorkbook(new File(filePath));
			} catch (InvalidFormatException | IOException e) {
				e.printStackTrace();
				throw new ServiceException(e);
			}
		
			IDataUtil.put( pipelineCursor, "workbook", jtwb.getWorkbook() );
			
		} finally {
			pipelineCursor.destroy();
		}
			
		// --- <<IS-END>> ---

                
	}



	public static final void save (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(save)>> ---
		// @sigtype java 3.5
		// [i] object:0:required workbook
		// [i] field:0:required filePath
		// [o] object:0:required workbook
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
			Workbook workbook = (Workbook) IDataUtil.get( pipelineCursor, "workbook" );
			String	filePath = IDataUtil.getString( pipelineCursor, "filePath" );
			pipelineCursor.destroy();
			
			
			FileOutputStream fileOut;
			
			if (filePath.endsWith(".xls") && workbook instanceof XSSFWorkbook) {
				filePath = filePath + "x";
			}
			
			try {
				fileOut = new FileOutputStream(filePath);
				workbook.write(fileOut);
				fileOut.close();
			} catch (FileNotFoundException e) {
				e.printStackTrace();
				throw new ServiceException(e);
			} catch (IOException e) {
				e.printStackTrace();
				throw new ServiceException(e);
			}
			
			IDataUtil.put( pipelineCursor, "workbook", workbook );
		} finally {
		pipelineCursor.destroy();
		}
			
		// --- <<IS-END>> ---

                
	}
}

