package model;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class CreatePresentation {
	
	CreatePresentation(String name) {
		try {
			//creating a new empty slide show
		      XMLSlideShow ppt = new XMLSlideShow();
		      
		      //creating an FileOutputStream object
		      File file =new File(name+".pptx");
		      FileOutputStream out = new FileOutputStream(file);
		      
		      //saving the changes to a file
		      ppt.write(out);
		      System.out.println("Presentation created successfully");
		      out.close();
			
		}catch(Exception e) {
			System.err.println("error while creating presentation");
		}
	}
   public static void main(String args[]) throws IOException{
   
      
   }
}