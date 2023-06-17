package com.example.worddemo;

import com.spire.doc.*;
import com.spire.doc.documents.TableRowHeightType;

import java.awt.*;

/**
 *
 */
public class AddRow {
    public static void main(String[] args) {
        //加载测试文档
        Document doc = new Document();
//        doc.loadFromFile("sample.docx");

        //获取表格
        Section section = doc.addSection();
//        section.addColumn(1.1f,1.2f);
//        Section section = doc.getSections().get(0);
        Table table = section.addTable(true);
        table.resetCells(10, 10);

        //将第一行设置为表格标题
        TableRow row = table.getRows().get(0);
        row.isHeader(true);
        row.setHeight(20);
        row.setHeightType(TableRowHeightType.Exactly);
        row.getRowFormat().setBackColor(Color.gray);

//保存文档
        doc.saveToFile("基础行.docx", FileFormat.Docx_2013);
        doc.dispose();
    }
}