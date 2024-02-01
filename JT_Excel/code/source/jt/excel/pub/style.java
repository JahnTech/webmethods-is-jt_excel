package jt.excel.pub;

// -----( IS Java Code Template v1.2

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
// --- <<IS-END-IMPORTS>> ---

public final class style

{
	// ---( internal utility methods )---

	final static style _instance = new style();

	static style _newInstance() { return new style(); }

	static style _cast(Object o) { return (style)o; }

	// ---( server methods )---




	public static final void applyToCells (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(applyToCells)>> ---
		// @sigtype java 3.5
		// [i] object:0:required sheet
		// [i] object:0:required style
		// [i] field:0:required columnStart
		// [i] field:0:required rowStart
		// [i] field:0:optional columnEnd
		// [i] field:0:optional rowEnd
		// [o] object:0:required sheet
		// [o] object:0:required style
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		try {
			Sheet sheet = (Sheet) IDataUtil.get( pipelineCursor, "sheet" );
			CellStyle style = (CellStyle) IDataUtil.get( pipelineCursor, "style" );
			int	columnStart = getIntegerFromString(IDataUtil.getString( pipelineCursor, "columnStart" ),0);
			int	rowStart = getIntegerFromString(IDataUtil.getString( pipelineCursor, "rowStart" ),0);
			int	columnEnd = getIntegerFromString(IDataUtil.getString( pipelineCursor, "columnEnd" ),columnStart);
			int	rowEnd = getIntegerFromString(IDataUtil.getString( pipelineCursor, "rowEnd" ),rowStart);
			
			CellStyle tmpStyle = sheet.getWorkbook().createCellStyle();
			tmpStyle.cloneStyleFrom(style);
			
			Cell cell = null;
			
			for (int row = rowStart; row <= rowEnd; row++ ){
				for (int column = columnStart; column <= columnEnd; column++){
					cell = getValidRow(sheet, row).getCell(column);
					cell.setCellStyle(tmpStyle);
				}
			}
			
			IDataUtil.put( pipelineCursor, "sheet", sheet );
			IDataUtil.put( pipelineCursor, "style", style );
		} finally {
			pipelineCursor.destroy();
		}
		// pipeline	
			
		// --- <<IS-END>> ---

                
	}



	public static final void define (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(define)>> ---
		// @sigtype java 3.5
		// [i] object:0:required workbook
		// [i] record:0:optional border
		// [i] - field:0:optional top {"none","thin","medium","thick"}
		// [i] - field:0:optional bottom {"none","thin","medium","thick"}
		// [i] - field:0:optional left {"none","thin","medium","thick"}
		// [i] - field:0:optional right {"none","thin","medium","thick"}
		// [i] record:0:optional font
		// [i] - field:0:optional bold {"true","false"}
		// [i] - field:0:optional underlined {"none","single","double"}
		// [i] - field:0:optional italic {"true","false"}
		// [i] - field:0:optional align {"left","right","center"}
		// [i] - field:0:optional size {"8","10","12","14","16"}
		// [i] record:0:optional color
		// [i] - field:0:optional foreground {"WHITE","BLACK","GREY_25_PERCENT","GREY_40_PERCENT","GREY_50_PERCENT","GREY_80_PERCENT","LIGHT_BLUE","LIGHT_CORNFLOWER_BLUE","LIGHT_GREEN","LIGHT_ORANGE","LIGHT_TURQUOISE","LIGHT_YELLOW"}
		// [o] object:0:required workbook
		// [o] object:0:required style
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		
		try {
			Workbook wb = (Workbook) IDataUtil.get( pipelineCursor, "workbook" );
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			
			// font begin
			IData	fontDoc = IDataUtil.getIData( pipelineCursor, "font" );
			if ( fontDoc != null) {
				
				IDataCursor fontCursor = fontDoc.getCursor();
				Boolean	bold = IDataUtil.getBoolean( fontCursor, "bold", false);
				String	underlined = IDataUtil.getString( fontCursor, "underlined" );
				Boolean	italic = IDataUtil.getBoolean( fontCursor, "italic", false );
				String	align = IDataUtil.getString( fontCursor, "align" );
				String	size = IDataUtil.getString( fontCursor, "size" );
				
				
				
				if (bold) { 
				// font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
					font.setBold(true);
				}
				if (underlined != null) {
					if (underlined.equals("single")) {
						font.setUnderline(HSSFFont.U_SINGLE);
					}
					if (underlined.equals("double")) {
						font.setUnderline(HSSFFont.U_DOUBLE);
					}
				}
				if (italic) {
					font.setItalic(true);
				}
				if (align != null) {
					if (align.equals("left")) {
						style.setAlignment(HorizontalAlignment.LEFT);
					}
					if (align.equals("right")) {
						style.setAlignment(HorizontalAlignment.RIGHT);
					}
					if (align.equals("center")) {
						style.setAlignment(HorizontalAlignment.CENTER);
					}
				}
				if (size!=null) {
					font.setFontHeightInPoints(Short.parseShort(size));
				}
				
				fontCursor.destroy();
				style.setFont(font);
			}
			// font end
			
			
			// border begin
			IData	border = IDataUtil.getIData( pipelineCursor, "border" );
			if ( border != null)
			{
				IDataCursor borderCursor = border.getCursor();
				String	top = IDataUtil.getString( borderCursor, "top" );
				String	bottom = IDataUtil.getString( borderCursor, "bottom" );
				String	left = IDataUtil.getString( borderCursor, "left" );
				String	right = IDataUtil.getString( borderCursor, "right" );
				borderCursor.destroy();
				
				if (top != null) {
					if (top.equals("thin")) {
						style.setBorderTop(BorderStyle.THIN);
					    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
					} else if (top.equals("medium")){
						style.setBorderTop(BorderStyle.MEDIUM);
					    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
					} else if (top.equals("thick")){
						style.setBorderTop(BorderStyle.THICK);
					    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
					}
				}
				
				if (bottom!=null){
					if (bottom.equals("thin")) {
						style.setBorderBottom(BorderStyle.THIN);
					    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
					} else if (bottom.equals("medium")){
						style.setBorderBottom(BorderStyle.MEDIUM);
					    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
					} else if (bottom.equals("thick")){
						style.setBorderBottom(BorderStyle.THICK);
					    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
					}
				}
				
				if (left!=null){
					if (left.equals("thin")) {
						style.setBorderLeft(BorderStyle.THIN);
					    style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
					} else if (left.equals("medium")){
						style.setBorderLeft(BorderStyle.MEDIUM);
					    style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
					} else if (left.equals("thick")){
						style.setBorderLeft(BorderStyle.THICK);
					    style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
					}
				}
				
				if (right!=null){
					if (right.equals("thin")) {
						style.setBorderRight(BorderStyle.THIN);
					    style.setRightBorderColor(IndexedColors.BLACK.getIndex());
					} else if (right.equals("medium")){
						style.setBorderRight(BorderStyle.MEDIUM);
					    style.setRightBorderColor(IndexedColors.BLACK.getIndex());
					} else if (right.equals("thick")){
						style.setBorderRight(BorderStyle.THICK);
					    style.setRightBorderColor(IndexedColors.BLACK.getIndex());
					}
				}
			}
			// border end
			
			// color begin
			IData	color = IDataUtil.getIData( pipelineCursor, "color" );
			if ( color != null)
			{
				IDataCursor colorCursor = color.getCursor();
					String	foreground = IDataUtil.getString( colorCursor, "foreground" );
				colorCursor.destroy();
				
				if(foreground!=null){
					if(foreground.equals("WHITE")){
		//						Color color1 = new HSSFColor().;
						
						style.setFillForegroundColor(HSSFColorPredefined.WHITE.getIndex());
					} else if (foreground.equals("BLACK")){
						style.setFillForegroundColor(HSSFColorPredefined.BLACK.getIndex());
					} else if (foreground.equals("GREY_25_PERCENT")){
						style.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
					} else if (foreground.equals("GREY_40_PERCENT")){
						style.setFillForegroundColor(HSSFColorPredefined.GREY_40_PERCENT.getIndex());
					} else if (foreground.equals("GREY_50_PERCENT")){
						style.setFillForegroundColor(HSSFColorPredefined.GREY_50_PERCENT.getIndex());
					} else if (foreground.equals("GREY_80_PERCENT")){
						style.setFillForegroundColor(HSSFColorPredefined.GREY_80_PERCENT.getIndex());
					} else if (foreground.equals("LIGHT_BLUE")){
						style.setFillForegroundColor(HSSFColorPredefined.LIGHT_BLUE.getIndex());
					}  else if (foreground.equals("LIGHT_CORNFLOWER_BLUE")){
						style.setFillForegroundColor(HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex());
					} else if (foreground.equals("LIGHT_GREEN")){
						style.setFillForegroundColor(HSSFColorPredefined.LIGHT_GREEN.getIndex());
					} else if (foreground.equals("LIGHT_ORANGE")){
						style.setFillForegroundColor(HSSFColorPredefined.LIGHT_GREEN.getIndex());
					} else if (foreground.equals("LIGHT_TURQUOISE")){
						style.setFillForegroundColor(HSSFColorPredefined.LIGHT_TURQUOISE.getIndex());
					} else if (foreground.equals("LIGHT_YELLOW")){
						style.setFillForegroundColor(HSSFColorPredefined.LIGHT_YELLOW.getIndex());
					} 
				    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				}
			}
			// color end
			
			IDataUtil.put( pipelineCursor, "workbook", wb );
			IDataUtil.put( pipelineCursor, "style", style );
		} finally {
			pipelineCursor.destroy();
		}
		// pipeline
			
		// --- <<IS-END>> ---

                
	}

	// --- <<IS-START-SHARED>> ---
	public static int getIntegerFromString(String value, int defaultValue){
		int i = defaultValue;
		
		try {
			i = Integer.parseInt(value);
		} catch (NumberFormatException e) {
			// TODO Auto-generated catch blok
			
		}
		
		return i;
	}
	
	public static org.apache.poi.ss.usermodel.Row getValidRow(Sheet mSheet, int mRow){
		Row mZeile;		
		
		//System.out.println("req:"+mRow+" last:"+mSheet.getLastRowNum());
		
		if( mRow <= mSheet.getLastRowNum() )
		{
			mZeile = mSheet.getRow(mRow);
		}
		else	
		{
			mZeile = mSheet.createRow(mRow);	
		}
	
		return mZeile;
	}
		
	// --- <<IS-END-SHARED>> ---
}

