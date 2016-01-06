package main.java.com.pji.excel.serviceImpl;

import java.util.List;

import main.java.com.pji.excel.en.HeaderFooterEn;
import main.java.com.pji.excel.service.HeaderFooterInterface;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

//created by arief cungip
//bekasi simple code
/**
 * @author arief.pramudyo
 * @since 12/21/2015
 * @version 1.0
 */
public class HeaderFooterImpl implements HeaderFooterInterface {

	@Override
	public void addHeader(WritableSheet sheet, Integer column, Integer row,
			String valueHeader, WritableCellFormat cellFormat) {
		try {
			Label label;
			if (cellFormat != null)
				label = new Label(column, row, valueHeader, cellFormat);
			else
				label = new Label(column, row, valueHeader);
			sheet.addCell(label);

/*			CellView cv = sheet.getColumnView(column);
			cv.setAutosize(true);
			sheet.setColumnView(column, cv);
*/
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@Override
	public Integer addHeaderMatriks(WritableSheet sheet,
			List<main.java.com.pji.excel.model.HeaderFooter> model,
			Integer rightPosition) {

		Integer rowLeft = 0;
		Integer rowCenter = 0;
		Integer rowRight = 0;
		for (int i = 0; i < model.size(); i++) {
			if (model.get(i).positionHeader
					.equalsIgnoreCase(HeaderFooterEn.Right
							.getHeaderFooterType())) {
				if (rightPosition != null)
					addHeader(sheet, rightPosition, rowRight, model.get(i)
							.getHeader(), model.get(i)
							.getWritableCellFormatHeader());
				else
					addHeader(sheet, 2, rowRight, model.get(i)
							.getHeader(), model.get(i)
							.getWritableCellFormatHeader());
				rowRight++;
			} else if (model.get(i).positionHeader
					.equalsIgnoreCase(HeaderFooterEn.Center
							.getHeaderFooterType())) {
				try {
					if(rightPosition!=null)
					sheet.mergeCells(1, rowCenter, rightPosition - 1, rowCenter);
				} catch (WriteException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				addHeader(sheet, 1, rowCenter, model.get(i).getHeader(), model
						.get(i).getWritableCellFormatHeader());
				rowCenter++;
			} else {
				addHeader(sheet, 0, rowCenter, model.get(i).getHeader(), model
						.get(i).getWritableCellFormatHeader());
				rowCenter++;
			}
		}

		if (rowLeft > rowCenter) {
			if (rowRight > rowLeft) {
				return rowRight;
			} else {
				return rowLeft;
			}
		} else {
			if (rowRight > rowCenter) {
				return rowRight;
			} else {
				return rowCenter;
			}
		}

	}

}
