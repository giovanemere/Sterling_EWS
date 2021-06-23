package com.aossas.o365;

import com.aossas.o365.ReadInbox;

import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceVersionException;


import org.junit.Assert;
import org.junit.Ignore;
import org.junit.Test;

public class ReadInboxTest  {
	/**
	 * Initialising a ReadInbox with missing Domain should lead to an Exception.
	 * @throws Exception 
	*/ 
    @Ignore
    @Test (expected = Exception.class)
    public void setReadInboxNegativeDomain() throws Exception
    {
    	
    	@SuppressWarnings("unused")
		ReadInbox inbox = new ReadInbox("usuario@dominio.com", "password");
    }

	/**
	 * Initialising a ReadInbox with missing Domain should lead to an Exception.
	 * @throws Exception 
	*/ 
    @Ignore
    @Test (expected = Exception.class)
    public void setReadInboxNegativeUser() throws Exception
    {
    	@SuppressWarnings("unused")
    	ReadInbox inbox = new ReadInbox("usuario@aossas0.onmicrosoft.com", "password");
    }

	/**
	 * Initialising a ReadInbox with wrong Domain should lead to an Exception.
	 * @throws Exception 
	*/ 
    @Ignore
    @Test (expected = Exception.class)
    public void setReadInboxNegativePassword() throws Exception
    {
    	@SuppressWarnings("unused")
    	ReadInbox inbox = new ReadInbox("juan.andres@aossas0.onmicrosoft.com", "password");
    }
    

	/**
	 * Initialising a ReadInbox with something should lead to an IllegalArgumentException.
	 * @throws IllegalArgumentException 
	*/ 
    @Ignore
    @Test (expected = IllegalArgumentException.class)
    public void readMailFolderIllegalArgumentException() throws IllegalArgumentException, ServiceVersionException, ServiceLocalException, Exception 
    {
    	ReadInbox inbox = new ReadInbox("usuario@dominio.com", "password");
		inbox.readMailInbox();	
    }

	/**
	 * Initialising a ReadInbox with something should lead to an ServiceVersionException.
	 * @throws ServiceVersionException 
	*/ 
    @Ignore
    @Test (expected = ServiceVersionException.class)
    public void readMailFolderServiceVersionException() throws IllegalArgumentException, ServiceVersionException, ServiceLocalException, Exception
    {
    	ReadInbox inbox = new ReadInbox("usuario@dominio.com", "password");
    	inbox.readMailInbox();	
    }

	/**
	 * Initialising a ReadInbox with something should lead to an ServiceLocalException.
	 * @throws ServiceLocalException 
	*/ 
    @Ignore
    @Test (expected = ServiceLocalException.class)
    public void readMailFolderServiceLocalException() throws IllegalArgumentException, ServiceVersionException, ServiceLocalException, Exception
    {
    	ReadInbox inbox = new ReadInbox("juan.andres@aossas0.onmicrosoft.com", "Colpatria_2019");
    	inbox.readMailInbox();
    }

	/**
	 * Initialising a ReadInbox with missing Domain should lead to an Exception.
	 * @throws Exception 
	*/ 
    @Ignore
    @Test (expected = Exception.class)
    public void readMailFolderException() throws IllegalArgumentException, ServiceVersionException, ServiceLocalException, Exception
    {
    	ReadInbox inbox = new ReadInbox("usuario@dominio.com", "password");
    	inbox.readMailInbox();
    }

    @Ignore
    @Test
    public void setReadInbox() 
    {
    	try {
    		//TODO reemplazar con las credenciales correctas
	    	ReadInbox inbox = new ReadInbox("juan.andres@aossas0.onmicrosoft.com", "Colpatria_2019");
	    	inbox.readMailInbox();
    		Assert.assertTrue(true);
    	}catch(Exception e) {
    		Assert.assertTrue(false);
    	}
    }
    
    @Test
    public void readMailFoder() 
    {
    	try {
    		//TODO reemplazar con las credenciales correctas
	    	ReadInbox inbox = new ReadInbox("juan.andres@aossas0.onmicrosoft.com", "Colpatria_2019");
	    	inbox.readMailFoder("Inbox", "eduart.doria@aossas.com");
    		Assert.assertTrue(true);
    	}catch(Exception e) {
    		Assert.assertTrue(false);
    	}
    }
    
}
