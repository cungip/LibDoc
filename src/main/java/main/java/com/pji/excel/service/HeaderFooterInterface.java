package main.java.com.pji.excel.service;

import java.util.List;

import main.java.com.pji.excel.model.HeaderFooter;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;

//created by arief cungip
//bekasi simple code
/**
 * @author arief.pramudyo
 * @since 12/21/2015
 * @version 1.0
 */
public interface HeaderFooterInterface {
	public void addHeader(WritableSheet sheet,Integer column,Integer row,String valueLabel,WritableCellFormat cellFormat);
	public Integer addHeaderMatriks(WritableSheet sheet, List<HeaderFooter> model,Integer rightPosition);
	}
