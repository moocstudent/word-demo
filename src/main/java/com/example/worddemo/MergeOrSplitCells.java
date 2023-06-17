package com.example.worddemo;

import com.spire.doc.*;

public class MergeOrSplitCells {
    public static void main(String[] args){
        //创建Document类的对象
        Document doc = new Document();
        Section sec = doc.addSection();

        //添加一个4行4列的表格
        Table tb= sec.addTable(true);
        tb.resetCells(4,4);

        //调用方法纵向合并第1列中的第2、3个单元格
        tb.applyVerticalMerge(0,1,2);
        //调用方法横向合并第1行中的第2、3个单元格
        tb.applyHorizontalMerge(0,1,2);

        //调用方法获取第4行中的第4个单元格，拆分成2列3行
        tb.getRows().get(3).getCells().get(3).splitCell(2,3);

        //保存文档
        doc.saveToFile("合并拆分单元格.docx",FileFormat.Docx_2010);
    }
}