package main.java.com.pji.excel.serviceImpl;

import main.java.com.pji.excel.model.Body;
import main.java.com.pji.excel.service.ToolInterface;

//created by arief cungip
//bekasi simple code
/**
 * @author arief.pramudyo
 * @since 12/21/2015
 * @version 1.0
 */
public class ToolImpl implements ToolInterface {
	
	public Body setBodyCell(Integer row, Integer col, String val, Body body) {
		body.getValueLabel().add(val);
		body.getRowLabel().add(row);
		body.getColumnLabel().add(col);
		return body;
	}

}
