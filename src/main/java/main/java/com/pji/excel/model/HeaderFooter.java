package main.java.com.pji.excel.model;

import jxl.write.WritableCellFormat;

//created by arief cungip
//bekasi simple code
/**
 * @author arief.pramudyo
 * @since 12/21/2015
 * @version 1.0
 */
public class HeaderFooter {
public String header;
public String positionHeader;
public WritableCellFormat writableCellFormatHeader;
public String footer;
public String positionFooter;
public WritableCellFormat writableCellFormatFooter;

public HeaderFooter(){}

public HeaderFooter(String header,String footer,String positionHeader,String positionFooter,WritableCellFormat writableCellFormatHeader,WritableCellFormat writableCellFormatFooter){
	this.header=header;
	this.footer=footer;
	this.writableCellFormatHeader=writableCellFormatHeader;
	this.footer=footer;
	this.positionFooter=positionFooter;
	this.writableCellFormatFooter=writableCellFormatFooter;
}

public String getHeader() {
	return header;
}

public void setHeader(String header) {
	this.header = header;
}

public String getPositionHeader() {
	return positionHeader;
}

public void setPositionHeader(String positionHeader) {
	this.positionHeader = positionHeader;
}

public WritableCellFormat getWritableCellFormatHeader() {
	return writableCellFormatHeader;
}

public void setWritableCellFormatHeader(
		WritableCellFormat writableCellFormatHeader) {
	this.writableCellFormatHeader = writableCellFormatHeader;
}

public String getFooter() {
	return footer;
}

public void setFooter(String footer) {
	this.footer = footer;
}

public String getPositionFooter() {
	return positionFooter;
}

public void setPositionFooter(String positionFooter) {
	this.positionFooter = positionFooter;
}

public WritableCellFormat getWritableCellFormatFooter() {
	return writableCellFormatFooter;
}

public void setWritableCellFormatFooter(
		WritableCellFormat writableCellFormatFooter) {
	this.writableCellFormatFooter = writableCellFormatFooter;
}

@Override
public String toString() {
	return "HeaderFooterModel [header=" + header + ", positionHeader="
			+ positionHeader + ", writableCellFormatHeader="
			+ writableCellFormatHeader + ", footer=" + footer
			+ ", positionFooter=" + positionFooter
			+ ", writableCellFormatFooter=" + writableCellFormatFooter + "]";
}

}
