/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package main.java.com.mycompany.printo;

/**
 *
 * @author Shivam
 */

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import java.io.*;
import java.io.File;
import java.io.OutputStream;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
//import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class main {
    public static void main(String[] args){
        try {
            Scanner sc= new Scanner(System.in);
            main a = new main();
            String src, dst, tempdst = "normal.pdf";
            int x;
            System.out.println("Enter input file source like D:\\a.txt ");
            src = sc.next();
            System.out.println("\nEnter destination file source like D:\\a.pdf ");
            dst = sc.next();
            System.out.println("\nEnter File type \n1.HTML\n2.Image\n3.Text\n4.Excel\n5.Word");
            x = sc.nextInt();
            switch (x){
                case 1:
                    a.html2pdf(src,tempdst);
                    watermrk(dst,tempdst);
                    break;
                case 2:
                    a.img2pdf(src,tempdst);
                    watermrk(dst,tempdst);
                    break;
                case 3:
                    a.txt2pdf(src,tempdst);
                    watermrk(dst,tempdst);
                    break;
                case 4:
                    a.excel2pdf(src, tempdst);
                    watermrk(dst,tempdst);
                    break;
                case 5:
                    a.word2pdf(src, tempdst);
                    watermrk(dst,tempdst);
                    break;
                default:
                    System.out.print("Enter a correct option : ");
                    break;
            }
        } catch (IOException | DocumentException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);              
        }
      
    }
    
    private void html2pdf(String src, String dst) {
        try {
            Document document = new Document();
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(dst));
            document.open();
            XMLWorkerHelper.getInstance().parseXHtml(writer, document, new FileInputStream(src));
            document.close();
        } catch (FileNotFoundException | DocumentException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void img2pdf(String src, String dst){
        FileOutputStream fileOutputStream = null;
        try {
            Document pdfdoc = new Document();
            fileOutputStream = new FileOutputStream(dst);
            PdfWriter writer = null;
            writer = PdfWriter.getInstance(pdfdoc, fileOutputStream);
            writer.open();
            pdfdoc.open();
            pdfdoc.add(com.itextpdf.text.Image.getInstance(src));
            System.out.println("Converted...");
            pdfdoc.close();
            writer.close();
        } catch (FileNotFoundException | DocumentException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fileOutputStream.close();
            } catch (IOException ex) {
                Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
    public void txt2pdf(String src, String dst)  {
        try {
            Document pdfDoc = new Document(PageSize.A4);
            PdfWriter.getInstance(pdfDoc, new FileOutputStream(dst));
            pdfDoc.open();
            Font myfont = new Font();
            myfont.setStyle(Font.NORMAL);
            myfont.setSize(11);
            pdfDoc.add(new Paragraph("\n"));
            BufferedReader br = new BufferedReader(new FileReader(src));
            String strLine;
            while ((strLine = br.readLine()) != null) {
                Paragraph para = new Paragraph(strLine + "\n", myfont);
                para.setAlignment(Element.ALIGN_JUSTIFIED);
                pdfDoc.add(para);
            }	
            pdfDoc.close();
            br.close();
            System.out.println("\n\nConverted...");
        } catch (FileNotFoundException | DocumentException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void word2pdf(String src, String dst)  {     
        try {
            InputStream in = new FileInputStream(new File(src));
            OutputStream out = new FileOutputStream(new File(dst));
           
            // 1) Load DOCX into XWPFDocument
            XWPFDocument document = new XWPFDocument(in);
            // 2) Prepare Pdf options
            PdfOptions options = PdfOptions.create();
           System.out.print(in);
            // 3) Convert XWPFDocument to Pdf
            PdfConverter.getInstance().convert(document, out, options);
            System.out.println("\n\nConverted...");
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        } 
    
    }
    
    public void excel2pdf(String src, String dst){
        try {
            Document pdfdoc = new Document();
            XSSFExcelExtractor Eextractor = null ; 
            FileOutputStream fileOutputStream = new FileOutputStream(dst);
            PdfWriter writer = null;
            writer = PdfWriter.getInstance(pdfdoc, fileOutputStream);
            writer.open();
            pdfdoc.open();
            File file = new File(src);
            FileInputStream fis = new FileInputStream(file);
            XSSFWorkbook book = new XSSFWorkbook(fis);
			
//            HSSFWorkbook book = new HSSFWorkbook(fis);  for older versions of ms office
            Eextractor = new XSSFExcelExtractor(book);
            String filedata = Eextractor.getText();
            pdfdoc.add(new Paragraph(filedata));
            System.out.println("Converted...");
            pdfdoc.close();
            writer.close();
        } catch (IOException | DocumentException ex) {
            Logger.getLogger(main.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public static void watermrk(String dst, String tempdst) throws IOException, DocumentException {

        // read existing pdf
        PdfReader reader = new PdfReader(tempdst);
        PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(dst));

        // text watermark
        Font FONT = new Font(Font.FontFamily.HELVETICA, 34, Font.BOLD, new GrayColor(0.5f));
        Phrase p = new Phrase("PRINTO", FONT);

        // properties
        PdfContentByte over;
        Rectangle pagesize;
        float x, y;

        // loop over every page
        int n = reader.getNumberOfPages();
        for (int i = 1; i <= n; i++) {

            // get page size and position
            pagesize = reader.getPageSizeWithRotation(i);
            x = (pagesize.getLeft() + pagesize.getRight()) / 2;
            y = (pagesize.getTop() + pagesize.getBottom()) / 2;
            over = stamper.getOverContent(i);
            over.saveState();

            // set transparency
            PdfGState state = new PdfGState();
            state.setFillOpacity(0.2f);
            over.setGState(state);


                ColumnText.showTextAligned(over, Element.ALIGN_CENTER, p, x, y, 0);
            

            over.restoreState();
        }
        stamper.close();
        reader.close();
    }
}
