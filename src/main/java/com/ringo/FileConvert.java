package com.ringo;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.*;

public class FileConvert {
    public static void PdfToDoc(String pdfPath,String doc1){
        PDDocument doc = null;
        OutputStream fos = null;
        Writer writer = null;
        PDFTextStripper stripper = null;
        try {
            doc = PDDocument.load(new File(pdfPath));
            fos = new FileOutputStream(doc1);
            writer = new OutputStreamWriter(fos, "UTF-8");
            stripper = new PDFTextStripper();
            int pageNumber = doc.getNumberOfPages();
            stripper.setSortByPosition(true);
            stripper.setStartPage(1);
            stripper.setEndPage(pageNumber);
            stripper.writeText(doc, writer);
            writer.close();
            doc.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("end..");
    }

    public static void DocToPdf(String doc,String pdf){
        ActiveXComponent app = null;
        String wordFile = doc;
        String pdfFile = pdf;
        System.out.println("开始转换...");
        // 开始时间
        long start = System.currentTimeMillis();
        try {
            // 打开word
            app = new ActiveXComponent("Word.Application");
            // 设置word不可见,很多博客下面这里都写了这一句话，其实是没有必要的，因为默认就是不可见的，如果设置可见就是会打开一个word文档，对于转化为pdf明显是没有必要的
            //app.setProperty("Visible", false);
            // 获得word中所有打开的文档
            Dispatch documents = app.getProperty("Documents").toDispatch();
            System.out.println("打开文件: " + wordFile);
            // 打开文档
            Dispatch document = Dispatch.call(documents, "Open", wordFile, false, true).toDispatch();
            // 如果文件存在的话，不会覆盖，会直接报错，所以我们需要判断文件是否存在
            File target = new File(pdfFile);
            if (target.exists()) {
                target.delete();
            }
            System.out.println("另存为: " + pdfFile);
            // 另存为，将文档报错为pdf，其中word保存为pdf的格式宏的值是17
            Dispatch.call(document, "SaveAs", pdfFile, 17);
            // 关闭文档
            Dispatch.call(document, "Close", false);
            // 结束时间
            long end = System.currentTimeMillis();
            System.out.println("转换成功，用时：" + (end - start) + "ms");
        }catch(Exception e) {
            e.getMessage();
            System.out.println("转换失败"+e.getMessage());
        }finally {
            // 关闭office
            app.invoke("Quit", 0);
        }
    }
}
