package main.java.com.pji.excel.en;

public enum HeaderJson {
	JSON("application/json"),
	Xml("application/xml"),
	URLENCODED("application/x-www-form-urlencoded");
	
	public final String headerType;

	private HeaderJson(String headerType) {
		this.headerType=headerType;
	}
	
	public String getHeaderType() {
		return headerType;
	}
	
	

}
