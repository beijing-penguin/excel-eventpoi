package com.dc.eventpoi.test.temp;

import com.dc.eventpoi.test.temp.read.StreamReadBaseCallBack;

public interface ReadCallBack extends RegCallBack{
	public void callBack(StreamReadBaseCallBack baseCallBack);
}
