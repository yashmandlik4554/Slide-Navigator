package model;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class TitleLayout {

	   public static void main(String args[]) throws IOException{
	   
	      //creating presentation
	      XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("D:\\SlideNavigator\\example1.pptx"));	    
	      	      	      
	      ArrayList<XSLFSlideMaster> list = (ArrayList<XSLFSlideMaster>) ppt.getSlideMasters();
	      
	      //getting the slide master object
	      XSLFSlideMaster slideMaster = list.get(0);      
	      
	      //get the desired slide layout 
	      XSLFSlideLayout titleLayout = slideMaster.getLayout(SlideLayout.TITLE);
	                                                     
	      //creating a slide with title layout
	      XSLFSlide slide1 = ppt.createSlide(titleLayout);
	      
	      //selecting the place holder in it 
	      XSLFTextShape title1 = slide1.getPlaceholder(0); 
	      
	      XSLFTextShape matter = slide1.getPlaceholder(1);
	      
	      String str = matter.getText();
	      
	      matter.setLineColor(Color.red);
	      
	      //setting the title init 
	     // title1.setText("Tutorials point");	      
	      //title1.getText();
	      
	      matter.setText(str + "this is mater");
	      
	      //create a file object
	      File file=new File("D:\\SlideNavigator\\Presentation2.pptx");
	      FileOutputStream out = new FileOutputStream(file);
	      
	      //save the changes in a PPt document
	      ppt.write(out);
	      System.out.println("slide cretated successfully");
	      out.close();  
	   }
	}