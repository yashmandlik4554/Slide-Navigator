package model;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class Deleteslide { 
   
   public static void main(String args[]) throws IOException{
   
      //Opening an existing slide
      File file=new File("D:\\SlideNavigator\\example1.pptx");
      XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));
      
      ArrayList<XSLFSlide> slides = (ArrayList<XSLFSlide>) ppt.getSlides();
      
      //XSLFSlide slides[] = ppt.getSlides();
      System.out.println(slides.size());
      
      /*XSLFTextShape shape = slides[0].getPlaceholder(0);
      
      XSLFTextShape shape = slides[2].getPlaceholder(1);
     
      
      slides[0].getPlaceholder(1).setText(slides[0].getPlaceholder(1).getText()+"done");
      					
      					
     /*System.out.println(slides[1].getPlaceholder(0).getText());
     System.out.println(slides[2].getPlaceholder(0).getText());
     System.out.println(slides[3].getPlaceholder(0).getText());*/
      
      //deleting a slide
      //ppt.removeSlide(1);
      
      
      
      //creating a file object
      FileOutputStream out = new FileOutputStream(file);
      
      //Saving the changes to the presentation
      ppt.write(out);
      out.close();	
   }
}