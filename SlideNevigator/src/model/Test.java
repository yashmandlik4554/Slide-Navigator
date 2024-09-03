package model;

import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.awt.Robot;
import java.awt.event.InputEvent;

public class Test {

	public Test(String cmd) {
		// TODO Auto-generated constructor stub
		
		try {		
			Robot robot = null; //Used to capture the screen
			
			//Get default screen device
            GraphicsEnvironment gEnv=GraphicsEnvironment.getLocalGraphicsEnvironment();
            GraphicsDevice gDev=gEnv.getDefaultScreenDevice();
            
          //Prepare Robot object
            robot = new Robot(gDev);
            
            //robot.keyPress(39);		//right key
            robot.mouseMove(500, 300);
            robot.mousePress(InputEvent.BUTTON1_MASK);
            
            //robot.keyPress(37);		//left key
		}catch(Exception e) {
			System.out.println(e.getMessage());
		}
		
	}

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		new Test("cmd /c  D:/SlideNavigator/abc.pptx");
	}

}
