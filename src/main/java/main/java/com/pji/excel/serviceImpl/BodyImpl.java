package main.java.com.pji.excel.serviceImpl;

import java.util.ArrayList;
import java.util.List;

import main.java.com.pji.excel.service.BodyInterface;
import jxl.CellView;
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
public class BodyImpl implements BodyInterface {

	/**
	 * @author arief.pramudyo
	 * @param sheet
	 *            ,column,row,valueCaption,cellFormat
	 * @since 12/21/2015
	 */
	@Override
	public void addCaption(WritableSheet sheet, Integer column, Integer row,
			String valueCaption, WritableCellFormat cellFormat) {
		// TODO Auto-generated method stub
		try {
			Label label;
			if (row == null)
				row = 0;

			if (cellFormat != null)
				label = new Label(column, row, valueCaption, cellFormat);
			else
				label = new Label(column, row, valueCaption);

			sheet.addCell(label);
			/*
			 * for (int i = 0; i < sheet.getColumns(); i++) { Cell[] cells =
			 * sheet.getColumn(i); int longestStrLen = -1;
			 * 
			 * if (cells.length == 0) continue;
			 * 
			 * Find the widest cell in the column. for (int j = 0; j <
			 * cells.length; j++) { if (cells[j].getContents().length() >
			 * longestStrLen) { String str = cells[j].getContents(); if (str ==
			 * null || str.isEmpty()) continue; longestStrLen =
			 * str.trim().length(); } }
			 * 
			 * If not found, skip the column. if (longestStrLen == -1) continue;
			 * 
			 * If wider than the max width, crop width if (longestStrLen > 255)
			 * longestStrLen = 255;
			 * 
			 * CellView cv = sheet.getColumnView(i); cv.setSize(longestStrLen *
			 * 256 + 100); Every character is 256 units wide, so scale it.
			 * 
			 * sheet.setColumnView(i, cv); }
			 */
			CellView cv = sheet.getColumnView(column);
			cv.setAutosize(true);
			sheet.setColumnView(column, cv);
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@Override
	public void addLabel(WritableSheet sheet, Integer column, Integer row,
			String valueLabel, WritableCellFormat cellFormat) {
		try {
			Label label;
			if (cellFormat != null)
				label = new Label(column, row, valueLabel, cellFormat);
			else
				label = new Label(column, row, valueLabel);
			sheet.addCell(label);
			/*
			 * CellView cv = sheet.getColumnView(column); cv.setAutosize(true);
			 * sheet.setColumnView(column, cv);
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
	public void generateAddBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption, Integer headerRowSize) {
		WritableCellFormat formatCellLabel = new WritableCellFormat();
		WritableCellFormat formatCellCaption = new WritableCellFormat();
		if (writableCellFormatLabel != null) {
			formatCellLabel = writableCellFormatLabel;
		} else {
			formatCellLabel = null;
		}
		if (writableCellFormatCaption != null) {
			formatCellCaption = writableCellFormatCaption;
		} else {
			formatCellCaption = null;
		}
		int rowLabel = 0;
		int colLabel = 0;
		int colCaption = 0;
		int rowCaption = 0;
		if (headerRowSize != null) {
			rowCaption = headerRowSize;
			rowLabel = rowCaption;
		}
		if (listLabel != null) {
			if (row == false) {
				if (listCaption != null) {
					rowLabel = rowLabel + 1;
					for (String caption : listCaption) {
						addCaption(sheet, colCaption, rowCaption, caption,
								formatCellCaption);
						colCaption++;
					}
				}
				for (List<String> labelRow : listLabel) {
					for (String labelColumn : labelRow) {
						addLabel(sheet, colLabel, rowLabel, labelColumn,
								formatCellLabel);
						colLabel++;
					}
					colLabel=0;
					rowLabel++;
				}
			} else {
				if (listCaption != null) {
					List<String> inputColumnEmpty = new ArrayList<String>();
					rowLabel = rowLabel + 1;
					for (String caption : listCaption) {
						addCaption(sheet, colCaption, rowCaption, caption, formatCellCaption);
						colCaption++;
					}
					int j=listCaption.size()-listLabel.size();
					for(int i = 0;i<j;i++){
						listLabel.add(inputColumnEmpty);
					}
				}
				for (List<String> labelColumn : listLabel) {
					if(!labelColumn.isEmpty()){
					for (String labelRow : labelColumn) {
						addLabel(sheet, colLabel, rowLabel, labelRow,
								formatCellLabel);
						rowLabel++;
					}
					rowLabel=rowLabel-labelColumn.size();
					}else{
						for(int i = 0;i<listLabel.get(0).size();i++){
							addLabel(sheet, colLabel, rowLabel, "",
									formatCellLabel);
							rowLabel++;	
						}
						rowLabel=rowLabel-listLabel.get(0).size();
					}
						
					colLabel++;
				}
			}
		}
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel) {
		generateAddBody(sheet, row, listLabel, null, null, null, null);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatCaption) {
		generateAddBody(sheet, row, listLabel, null, null,
				writableCellFormatCaption, null);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel, Integer headerRowSize) {
		// TODO Auto-generated method stub
		generateAddBody(sheet, row, listLabel, null, null, null, headerRowSize);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel, Integer headerRowSize) {
		// TODO Auto-generated method stub
		generateAddBody(sheet, row, listLabel, writableCellFormatLabel, null,
				null, headerRowSize);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption, Integer headerRowSize) {
		// TODO Auto-generated method stub
		generateAddBody(sheet, row, listLabel, writableCellFormatLabel,
				listCaption, writableCellFormatCaption, headerRowSize);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption, Integer headerRowSize) {
		// TODO Auto-generated method stub
		generateAddBody(sheet, row, listLabel, writableCellFormatLabel,
				listCaption, null, headerRowSize);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel, List<String> listCaption,
			WritableCellFormat writableCellFormatCaption, Integer headerRowSize) {
		// TODO Auto-generated method stub
		generateAddBody(sheet, row, listLabel, null, listCaption,
				writableCellFormatCaption, headerRowSize);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption) {
		// TODO Auto-generated method stub
		generateAddBody(sheet, row, listLabel, writableCellFormatLabel,
				listCaption, writableCellFormatCaption, null);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel, List<String> listCaption) {
		// TODO Auto-generated method stub
		generateAddBody(sheet, row, listLabel, writableCellFormatLabel,
				listCaption, null, null);
	}

	@Override
	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel, List<String> listCaption,
			WritableCellFormat writableCellFormatCaption) {
		// TODO Auto-generated method stub
		generateAddBody(sheet, row, listLabel, null, listCaption,
				writableCellFormatCaption, null);
	}
}
