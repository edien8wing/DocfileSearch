/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package docfilesearch;
import java.io.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 *
 * @author Administrator
 */
public class Docfilesearch {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        int beg=0;
        int end=60;
        String url="L:\\回复\\O&O DiskRecovery\\DOC";
        String titleName="";
        String [] lst=readAllFiles(url);
        for(int i=0;i<lst.length;i++){
            File thisFile= new File(lst[i]);
            String currentTxt=doc2String(thisFile).trim().replaceAll("[^0-9a-zA-Z\u4e00-\u9fa5.，,。？“”]+","");
            System.out.println(i+":    "+lst[i]+"        length:"+currentTxt.length());
            System.out.println("开始"+beg+"到"+end);
            if(end>currentTxt.length()){
                titleName=currentTxt.substring(beg, currentTxt.length());
            }
            else{ 
                titleName=currentTxt.substring(beg,end);
            }
            System.out.println(titleName);
            System.out.println("完整");
            System.out.println(currentTxt);
            System.out.println("***********************************************");
            thisFile.renameTo(new File("L:\\backup\\"+i+" "+titleName+".doc"));
        }
    }
    public static String[] readAllFiles(String url){
        File dir = new File(url);
        if(dir.isFile())
            return null;
        String [] lst = dir.list();
        for(int i = 0;i<lst.length;i++){
            lst[i]=url+"\\"+lst[i];
        }
        return lst;
    }
    
	public static String doc2String(File file) {
		String result = "";
		try {
			FileInputStream fis = new FileInputStream(file);
			HWPFDocument doc = new HWPFDocument(fis);
			StringBuilder txt=doc.getText();
                        result=txt.toString();
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;

        }
        public static String docx2String(File file)  {
        
        String result = "";
        try {
            FileInputStream fis = new FileInputStream(file);
            XWPFDocument xdoc = new XWPFDocument(fis);
            XWPFWordExtractor extractor = new XWPFWordExtractor(xdoc);
            result = extractor.getText();
            
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

}
