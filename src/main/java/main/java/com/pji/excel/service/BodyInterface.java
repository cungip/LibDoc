package main.java.com.pji.excel.service;

import java.util.List;

import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;

//created by arief cungip
//bekasi simple code
/**
 * @author arief.pramudyo
 * @since 12/21/2015
 * @version 1.0
 */
public interface BodyInterface {
	public void addCaption(WritableSheet sheet, Integer column, Integer row,
			String valueCaption, WritableCellFormat cellFormat);

	public void addLabel(WritableSheet sheet, Integer column, Integer row,
			String valueLabel, WritableCellFormat cellFormat);

	public void generateAddBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption, Integer headerRowSize);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel, Integer headerRowSize);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel, Integer headerRowSize);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption, Integer headerRowSize);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption, Integer headerRowSize);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel, List<String> listCaption,
			WritableCellFormat writableCellFormatCaption, Integer headerRowSize);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel,
			List<String> listCaption,
			WritableCellFormat writableCellFormatCaption);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel,
			WritableCellFormat writableCellFormatLabel, List<String> listCaption);

	public void addBody(WritableSheet sheet, Boolean row,
			List<List<String>> listLabel, List<String> listCaption,
			WritableCellFormat writableCellFormatCaption);
}
