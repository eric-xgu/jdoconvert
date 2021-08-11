package com.ringo;


import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) {
        String conv=args[0];
        if(conv.equals("pdf")){
            String word=args[1];
            String pdf=args[2];
            System.out.println("源文件："+word);
            System.out.println("pdf文件"+pdf);
            FileConvert.DocToPdf(word, pdf);
        }else if(conv.equals("doc")){
            String word=args[2];
            String pdf=args[1];
            FileConvert.PdfToDoc(pdf, word);
        }
        //FileConvert.DocToPdf(wordFile, pdfFile);
        //PdfDocument doc=new PdfDocument();
        //doc.loadFromFile("D:/新建DOCX文档.docx");
        //doc.saveToFile("D:/a.pdf", FileFormat.PDF);
    }
}

