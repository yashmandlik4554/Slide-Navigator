package model;

import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.ServerSocket;
import java.net.Socket;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class Main {
	
	public static String presentationName = null;
	public static Process process = null;
	
	public Main() {
		// TODO Auto-generated constructor stub		
		try {
			ServerSocket serverSocket = new ServerSocket(8686);
			//while(true) {
				System.out.println("Waiting for connection");
				Socket socket = serverSocket.accept();
				System.out.println("Connected");
				new ClientHandler(socket);
			//}			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}			

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		new Main();
	}
}

