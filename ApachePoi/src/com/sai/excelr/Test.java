package com.sai.excelr;

public class Test {
	
	
	public String domainName(String str){
		String[] st = new String[2]; 
		st=str.split("@");
		
		
		return "www."+st[1];
	}
	
	public static void main(String[] args) {
		
		Test t = new Test();
		System.out.println(t.domainName("saigopalp@agtpltd.com"));
		
	}

}
