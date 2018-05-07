package com.qiqi.excel;

import java.io.Serializable;

public class ExcelDTO implements Serializable {

	private static final long serialVersionUID = 1L;

	private String companyName;
	
	private String orderType;
	
	private int totalNum;
	
	private int cancenlTotalNum;
	
	private int completedTotalNum;
	
	private String zeroString = "/";

	public String getZeroString() {
		return zeroString;
	}


	public int getTotalNum() {
		return totalNum;
	}


	public void setTotalNum(int totalNum) {
		this.totalNum = totalNum;
	}


	public int getCancenlTotalNum() {
		return cancenlTotalNum;
	}


	public void setCancenlTotalNum(int cancenlTotalNum) {
		this.cancenlTotalNum = cancenlTotalNum;
	}


	public int getCompletedTotalNum() {
		return completedTotalNum;
	}


	public void setCompletedTotalNum(int completedTotalNum) {
		this.completedTotalNum = completedTotalNum;
	}


	public String getCompanyName() {
		return companyName;
	}

	public void setCompanyName(String companyName) {
		this.companyName = companyName;
	}

	public String getOrderType() {
		return orderType;
	}

	public void setOrderType(String orderType) {
		this.orderType = orderType;
	}
	
	
}
