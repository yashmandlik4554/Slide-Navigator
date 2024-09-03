package model;

import java.net.Socket;

public class Connect {

	public Connect() {
		// TODO Auto-generated constructor stub
		
		try {
			Socket s = new Socket("192.168.1.2", 8684);
		}catch (Exception e) {
			// TODO: handle exception
		}
	}

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		new Connect();
	}

}
