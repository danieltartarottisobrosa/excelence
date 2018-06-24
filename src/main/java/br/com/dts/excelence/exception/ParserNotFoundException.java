package br.com.dts.excelence.exception;

public class ParserNotFoundException extends RuntimeException {
	private static final long serialVersionUID = 1065827952146657324L;
	
	private final Class<?> type;
	
	public ParserNotFoundException(Class<?> type) {
		this.type = type;
	}
	
	public Class<?> getType() {
		return type;
	}
}
