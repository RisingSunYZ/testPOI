package com.bian.testPOI;


import java.util.HashMap;
import java.util.Map;

public class App {

	private static Map<String,String> dataSource = null;
	
	static{
		dataSource = new HashMap<String, String>();
		dataSource.put("name", "yz");
		dataSource.put("addr", "gaoxin");
		dataSource.put("score", "89");
		
		dataSource.put("name", "zwl");
		dataSource.put("addr", "jingkai");
		dataSource.put("score", "73");
	}
	
//	@Test
//	private void testWritePOI(){
//		System.out.println(dataSource.get("name"));
//	}
	
	public static void main(String[] args) {
		System.out.println("111111111");
	}
	
}
