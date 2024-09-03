package model;

import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.Socket;
import java.util.ArrayList;

import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class ClientHandler extends Thread
{
	DataOutputStream dataOutputStream = null;
	DataInputStream dataInputStream = null;
	Socket socket = null;
	String path = "D:/SlideNavigator/";
	
	public ClientHandler(Socket socket) {
		// TODO Auto-generated constructor stub		
		this.socket = socket;
		try {
			dataOutputStream = new DataOutputStream(socket.getOutputStream());
			dataInputStream = new DataInputStream(socket.getInputStream());
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		start();
	}
	
	public void run() {
		try {			
				File file = new File(path);
				String fileNames[]=file.list();
				String data = "";
				for(int i=0;i<fileNames.length;i++) {
					data = data+","+fileNames[i];
				}			
				write(data);
			
				while(true) {
					//System.out.println("w8ing for cmd");				
					String command = dataInputStream.readUTF();				
					commandInterpretor(command.trim());
				}
		}catch(Exception e) {
			System.out.println("run "+e.getMessage());
		}
	}
	
	public void write(String data) {
		try {
			dataOutputStream.writeUTF(data.trim());			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void commandInterpretor(String command) {
		try {
			if(command.equals("create")) {							//create presentation
				String slideName = dataInputStream.readUTF();
				new CreatePresentation(path+slideName);
			} else if(command.equals("nslides")) {					//number of slides in presentation
				System.out.println("nslides");
				String slideName = dataInputStream.readUTF();
				System.out.println(slideName);
				if(!slideName.endsWith("pptx"))
					slideName = slideName + ".pptx";
			    File f=new File(path+slideName);
			    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(f));	      
			    //XSLFSlide slides[] = ppt.getSlides();			    
			    //write(slides.length+"");
			    
			    ArrayList<XSLFSlide> list = (ArrayList<XSLFSlide>) ppt.getSlides();
			    write(list.size()+"");
			    
			    Main.presentationName = slideName;
			    
			    try {		
					Process process = Runtime.getRuntime().exec("cmd /c D:/SlideNavigator/"+slideName);					
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}
			    
			} else if(command.equals("deletes")) {					//delete slide in presentation
				try {		
					Process process = Runtime.getRuntime().exec("taskkill /F /IM POWERPNT.EXE");
					Thread.sleep(2000);
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}
				String slideName = dataInputStream.readUTF();
				if(!slideName.endsWith("pptx"))
					slideName = slideName + ".pptx";
				String slideNumber = dataInputStream.readUTF();
			    File f=new File(path+slideName);
			    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(f));
			    ppt.removeSlide(Integer.parseInt(slideNumber));
			    FileOutputStream out = new FileOutputStream(f);
			    ppt.write(out);
			    out.close();
			    try {
					Process process = Runtime.getRuntime().exec("cmd /c D:/SlideNavigator/"+slideName);					
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}
			    System.out.println("slide deleted");
			} else if(command.equals("deletep")) {					//delete presentation
				String slideName = dataInputStream.readUTF();
			    File f=new File(path+slideName);
			    f.delete();
			    System.out.println("Presentation deleted");
			} else if(command.equals("creates")) {					//create slide in presentation
				
				String slideName = dataInputStream.readUTF();
				String matter = dataInputStream.readUTF();
				
				try {
					Process process = Runtime.getRuntime().exec("taskkill /F /IM POWERPNT.EXE");
					Thread.sleep(2000);
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}
				
				//String slideTitle = dataInputStream.readUTF();
			    File f=new File(path+slideName);
			    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(f));	
			    
			    //getting the slide master object
			    ArrayList<XSLFSlideMaster> list = (ArrayList<XSLFSlideMaster>) ppt.getSlideMasters();
			      
			    //getting the slide master object
			    XSLFSlideMaster slideMaster = list.get(0);    
			   // XSLFSlideMaster slideMaster = ppt.getSlideMasters()[0];
			    //get the desired slide layout 
			    XSLFSlideLayout titleLayout = slideMaster.getLayout(SlideLayout.TITLE);
			    
			    //creating a slide with title layout
			    XSLFSlide slide1 = ppt.createSlide(titleLayout);
			    
			    //selecting the place holder in it 
			    XSLFTextShape title1 = slide1.getPlaceholder(0); 
			      
			    //setting the title init
			    title1.setText(matter);
			      
			    //ppt.createSlide();
			    FileOutputStream out = new FileOutputStream(f);
			    ppt.write(out);
			    out.close();
			    try {		
					Main.process = Runtime.getRuntime().exec("cmd /c D:/SlideNavigator/"+slideName);
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}
			    System.out.println("Slide Created");
			} else if(command.equals("slideContaint")) {
								
				String slideName = dataInputStream.readUTF();
				if(!slideName.endsWith("pptx"))
					slideName = slideName + ".pptx";
				String slideNumber = dataInputStream.readUTF();
				String cmd = dataInputStream.readUTF();
				
			    File f=new File(path+slideName);
			    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(f));
			    readPPT(ppt, Integer.parseInt(slideNumber), slideName);	
			    
			    
			    System.out.println("Navigation Command : "+cmd);
			    
			    /*Robot robot = null; //Used to capture the screen
				
				//Get default screen device
	            GraphicsEnvironment gEnv=GraphicsEnvironment.getLocalGraphicsEnvironment();
	            GraphicsDevice gDev=gEnv.getDefaultScreenDevice();
	            
	            //Prepare Robot object
	            robot = new Robot(gDev);
	            robot.keyPress(116);
			    
			    if(cmd.equalsIgnoreCase("left")) {
			    	 robot.keyPress(37);		//left key
			    } else if(cmd.equalsIgnoreCase("right")) {
			    	 robot.keyPress(39);		//right key
			    }*/
			    
			    /*try {		
					Process process = Runtime.getRuntime().exec("taskkill /F /IM POWERPNT.EXE");
					Thread.sleep(1000);
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}		    	
		         
		    	FileOutputStream out = new FileOutputStream(f);
			    ppt.write(out);
			    out.close();
		    	
		    	try {		
					Main.process = Runtime.getRuntime().exec("cmd /c D:/SlideNavigator/"+slideName);
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}*/
			    
			} else if(command.equals("slideshow")) {
				//up arrow - 38
				//down arrow - 40
				
				String slideName = dataInputStream.readUTF();
				if(!slideName.endsWith("pptx"))
					slideName = slideName + ".pptx";
				String slideNumber = dataInputStream.readUTF();
				String cmd = dataInputStream.readUTF();

				File f=new File(path+slideName);
				XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(f));
				readPPT(ppt, Integer.parseInt(slideNumber), slideName);	

				System.out.println("Navigation Command : "+cmd);

				Robot robot = null; //Used to capture the screen

				//Get default screen device
				GraphicsEnvironment gEnv=GraphicsEnvironment.getLocalGraphicsEnvironment();
				GraphicsDevice gDev=gEnv.getDefaultScreenDevice();

				//Prepare Robot object
				robot = new Robot(gDev);
				//robot.keyPress(116);
				
				if(cmd.trim().length()>0) {
					
				} else {
					System.out.println("Slide Show Number : " + slideNumber);
					if(slideNumber.trim().length()>0) {
						int n = Integer.parseInt(slideNumber.trim());
						for(int i=0;i<n;i++) {
							robot.keyPress(40);
							robot.keyRelease(116);
						}
					}
				}
				
				
				robot.keyPress(KeyEvent.VK_SHIFT);
				robot.keyPress(116);				
				robot.keyRelease(116);
				robot.keyRelease(KeyEvent.VK_SHIFT);

				if(cmd.equalsIgnoreCase("left")) {
					robot.keyPress(37);		//left key
				} else if(cmd.equalsIgnoreCase("right")) {
					robot.keyPress(39);		//right key
				}
				
			}  else if(command.equals("edit")) {
				System.out.println("edit");				
	
				String slideName = dataInputStream.readUTF();
				if(!slideName.endsWith("pptx"))
					slideName = slideName + ".pptx";
				String slideNumber = dataInputStream.readUTF();
				String matterStr = dataInputStream.readUTF();
				System.out.println("Slide Matter : " + matterStr);

				try {		
					Process process = Runtime.getRuntime().exec("taskkill /F /IM POWERPNT.EXE");
					Thread.sleep(1000);
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}		    	

				System.out.println(path+slideName);
				//creating presentation
				File f=new File(path+slideName);
				XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(f));
			      //XMLSlideShow ppt = new XMLSlideShow();	
				ArrayList<XSLFSlide> slides = (ArrayList<XSLFSlide>) ppt.getSlides();
				//XSLFSlide slides[] = ppt.getSlides();
				
				//String slideContent = slides[0].getPlaceholder(1).getText();
				
				try {
					XSLFSlide slide = slides.get(Integer.parseInt(slideNumber));
					slide.getPlaceholder(1).getText();
				//slides[Integer.parseInt(slideNumber)].getPlaceholder(1).setText(matterStr);
					slide.getPlaceholder(1).setText(matterStr);
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}
			      //create a file object
			      FileOutputStream out = new FileOutputStream(f);
			      ppt.write(out);
			      out.close();
			      			      
			      System.out.println("updated");

				try {		
					Main.process = Runtime.getRuntime().exec("cmd /c D:/SlideNavigator/"+slideName);
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}



			} else if(command.equals("close")) {							//close presentation
				try {		
					Process process = Runtime.getRuntime().exec("taskkill /F /IM POWERPNT.EXE");
				}catch(Exception e) {
					System.out.println(e.getMessage());
				}
			}
			else if(command.equals("left")) {							//close presentation
				Robot robot = null; //Used to capture the screen
				
				//Get default screen device
	            GraphicsEnvironment gEnv=GraphicsEnvironment.getLocalGraphicsEnvironment();
	            GraphicsDevice gDev=gEnv.getDefaultScreenDevice();
	            
	          //Prepare Robot object
	            robot = new Robot(gDev);
	            
	            //robot.keyPress(39);		//right key
	            
	            robot.keyPress(37);		//left key
			}
			else if(command.equals("right")) {							//close presentation
				Robot robot = null; //Used to capture the screen
				
				//Get default screen device
	            GraphicsEnvironment gEnv=GraphicsEnvironment.getLocalGraphicsEnvironment();
	            GraphicsDevice gDev=gEnv.getDefaultScreenDevice();
	            
	          //Prepare Robot object
	            robot = new Robot(gDev);
	            
	            robot.keyPress(39);		//right key
	            
	            //robot.keyPress(37);		//left key
			}
			else if(command.equals("esc")) {							//close presentation
				Robot robot = null; //Used to capture the screen
				
				//Get default screen device
	            GraphicsEnvironment gEnv=GraphicsEnvironment.getLocalGraphicsEnvironment();
	            GraphicsDevice gDev=gEnv.getDefaultScreenDevice();
	            
	          //Prepare Robot object
	            robot = new Robot(gDev);
	            
	            robot.keyPress(27);		//esc key
	            
	            //robot.keyPress(37);		//left key
			}
		} catch(Exception e) {
			System.out.println("Errr : "+e.getMessage());
		}
	}
	
	public void readPPT(XMLSlideShow ppt, int n, String slideName) { 
		
        CoreProperties props = ppt.getProperties().getCoreProperties();
        String title = props.getTitle();
        System.out.println("Title: " + title);
        
        String slideText = "";   
        ArrayList<XSLFSlide> slides = (ArrayList<XSLFSlide>) ppt.getSlides();
        
        //XSLFSlide slides[] = ppt.getSlides();        
        System.out.println("Starting slide...");
        ArrayList<XSLFShape> shapes = (ArrayList<XSLFShape>) slides.get(n).getShapes();        	
    	for (XSLFShape shape: shapes) {
    		if (shape instanceof XSLFTextShape) {
    	        XSLFTextShape textShape = (XSLFTextShape)shape;
    	        String text = textShape.getText();
    	        System.out.println("Text: " + text);
    	        slideText = slideText + "\n" + text;    	        
    		}
    	}
    	write(slideText);
    	
    	 /*XSLFSlide selectesdslide = slides[n];
         //bringing it to the top
         ppt.setSlideOrder(selectesdslide, 0);
    	
    	
    	
        
        /*for (XSLFSlide slide: ppt.getSlides()) {
        	System.out.println("Starting slide...");
        	XSLFShape[] shapes = slide.getShapes();        	
        	for (XSLFShape shape: shapes) {
        		if (shape instanceof XSLFTextShape) {
        	        XSLFTextShape textShape = (XSLFTextShape)shape;
        	        String text = textShape.getText();
        	        System.out.println("Text: " + text);
        		}
        	}
        }*/
	}	
}