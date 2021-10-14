package com.dc.eventpoi.test;

import java.math.BigDecimal;

public class Person {
    private byte[] headImage;
	private String no;
	private String name;
	private Integer age;
	private BigDecimal salary;
	private BigDecimal oldSalary;
	private String remark;
	private Integer caigouNum;
	
    public Integer getCaigouNum() {
		return caigouNum;
	}
	public void setCaigouNum(Integer caigouNum) {
		this.caigouNum = caigouNum;
	}
	public byte[] getHeadImage() {
        return headImage;
    }
    public void setHeadImage(byte[] headImage) {
        this.headImage = headImage;
    }
    public String getRemark() {
        return remark;
    }
    public void setRemark(String remark) {
        this.remark = remark;
    }
    
	public String getNo() {
		return no;
	}
	public void setNo(String no) {
		this.no = no;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	
	public Integer getAge() {
		return age;
	}
	public void setAge(Integer age) {
		this.age = age;
	}
	public BigDecimal getSalary() {
		return salary;
	}
	public void setSalary(BigDecimal salary) {
		this.salary = salary;
	}
	public BigDecimal getOldSalary() {
		return oldSalary;
	}
	public void setOldSalary(BigDecimal oldSalary) {
		this.oldSalary = oldSalary;
	}
	
}
