package com.aossas.o365;
import com.sun.xml.internal.messaging.saaj.util.Base64;
//byte array to file 
	import java.io.File; 
	import java.io.FileOutputStream; 
	import java.io.OutputStream; 
	
	
public class UtilFunctions {
	public static String encodeToBase64(byte[] plainText) {
		//Base64 codec = new Base64();
		byte[] encoded = Base64.encode(plainText);
		return new String(encoded);
	}
	
	public static String encodeToBase64(String plainText) {
		return encodeToBase64(plainText.getBytes());
	}
	
	public static String decodeToBase64(String encoded) {
		//Base64 codec = new Base64();
		String decoded = Base64.base64Decode(encoded);
		return decoded;
	}
	
	 // Method which write the bytes into a file 
	
	public static void writeByte(String fullfilePath, byte[] bytes) 
    { 
		File file = new File (fullfilePath);
		writeByte(file, bytes);
    }
	
    public static void writeByte(File file, byte[] bytes) 
    { 
        try { 
  
            // Initialize a pointer 
            // in file using OutputStream 
            OutputStream 
                os 
                = new FileOutputStream(file); 
  
            // Starts writing the bytes in it 
            os.write(bytes); 
            System.out.println("Successfully"
                               + " byte inserted"); 
  
            // Close the file 
            os.close(); 
        } 
  
        catch (Exception e) { 
            System.out.println("Exception: " + e); 
        } 
    } 
}
