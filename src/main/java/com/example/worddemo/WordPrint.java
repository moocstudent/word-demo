package com.example.worddemo;

import com.spire.doc.*;
import java.awt.print.*;
public class WordPrint {

    public static void main(String[] args) throws Exception {
        //加载文档
        Document doc = new Document();
        doc.loadFromFile("创建表格.docx");

        PrinterJob loPrinterJob = PrinterJob.getPrinterJob();
        PageFormat loPageFormat = loPrinterJob.defaultPage();

        //设置打印纸张大小
        Paper loPaper = loPageFormat.getPaper();
        loPaper.setSize(600, 500);
        loPageFormat.setPaper(loPaper);

        //删除默认页边距
        loPaper.setImageableArea(0, 0, loPageFormat.getWidth(), loPageFormat.getHeight());
        //设置打印份数
        loPrinterJob.setCopies(1);
        loPrinterJob.setPrintable(doc, loPageFormat);
        //设置打印对话框
        if (loPrinterJob.printDialog()) {
            //执行打印
            try {
                loPrinterJob.print();
            } catch (PrinterException e)

            {
                e.printStackTrace();
            }
        }
    }
}