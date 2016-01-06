package main.java.com.pji.excel.model;

import java.util.List;

//created by arief cungip
//bekasi simple code
/**
 * @author arief.pramudyo
 * @since 12/21/2015
 * @version 1.0
 */
public class Body {
	
	public List<String> valueCaption;
	public List<String> valueLabel;
	public List<Integer> columnCaption; 
	public List<Integer> columnLabel; 
	public List<Integer> rowCaption;
	public List<Integer> rowLabel;
	
	public List<Integer> getColumnCaption() {
		return columnCaption;
	}

	public void setColumnCaption(List<Integer> columnCaption) {
		this.columnCaption = columnCaption;
	}

	public List<Integer> getColumnLabel() {
		return columnLabel;
	}

	public void setColumnLabel(List<Integer> columnLabel) {
		this.columnLabel = columnLabel;
	}

	public List<Integer> getRowCaption() {
		return rowCaption;
	}

	public void setRowCaption(List<Integer> rowCaption) {
		this.rowCaption = rowCaption;
	}

	public List<Integer> getRowLabel() {
		return rowLabel;
	}

	public void setRowLabel(List<Integer> rowLabel) {
		this.rowLabel = rowLabel;
	}

	public List<String> getValueCaption() {
		return valueCaption;
	}

	public void setValueCaption(List<String> valueCaption) {
		this.valueCaption = valueCaption;
	}

	public List<String> getValueLabel() {
		return valueLabel;
	}

	public void setValueLabel(List<String> valueLabel) {
		this.valueLabel = valueLabel;
	}

	@Override
	public String toString() {
		return "Body [valueCaption=" + valueCaption + ", valueLabel="
				+ valueLabel + ", columnCaption=" + columnCaption
				+ ", columnLabel=" + columnLabel + ", rowCaption=" + rowCaption
				+ ", rowLabel=" + rowLabel + "]";
	}

	
}
