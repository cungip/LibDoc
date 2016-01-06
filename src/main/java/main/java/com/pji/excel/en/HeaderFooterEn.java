package main.java.com.pji.excel.en;

public enum HeaderFooterEn {
	Left("left"),
	Center("center"),
	Right("right");
	
	public final String headerFooterType;

	private HeaderFooterEn(String headerFooterType) {
		this.headerFooterType=headerFooterType;
	}
	
	public String getHeaderFooterType() {
		return headerFooterType;
	}
	
	

}
