package com.dc.eventpoi.test;

public class ProductInfo {
    private String no;
    private String name;
    private String xinghao;
    private String desc;
    private byte[] headImage;
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
    public String getXinghao() {
        return xinghao;
    }
    public void setXinghao(String xinghao) {
        this.xinghao = xinghao;
    }
    public String getDesc() {
        return desc;
    }
    public void setDesc(String desc) {
        this.desc = desc;
    }
    
    
}
