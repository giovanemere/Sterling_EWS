package com.aossas.o365;

import org.junit.Test;

public class UtilMailTest {
	
	@Test
	public void testMethod() {
		try {
			UtilMail test = new UtilMail("juan.andres@aossas0.onmicrosoft.com", "Colpatria_2019");
			test.getMailList("Inbox", "eduart.doria@aossas.com");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
		
	}
}
