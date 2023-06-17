package com.example.worddemo;

import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;
import com.spire.doc.documents.TextSelection;
import com.spire.doc.fields.TextRange;


public class drawCombox {
    public static void main(String[] args) {

        //新建word 文档
        Document doc = new Document();
        Section section = doc.addSection();
        Paragraph para = section.addParagraph();
        para.setText("指定字符替换成复选框 symbol1, symbol2, symbol3.");

        //设置段落中字体样式
        ParagraphStyle style= new ParagraphStyle(doc);
        style.setName("paraStyle");
        style.getCharacterFormat().setFontName("宋体");
        style.getCharacterFormat().setFontSize(11f);
        doc.getStyles().add(style);
        para.applyStyle("paraStyle");

        //复选框打勾
        TextSelection selection1 = doc.findString("symbol1",true,true);
        TextRange tr1 = selection1.getAsOneRange();
        tr1.getCharacterFormat().setFontName("Wingdings 2");
        //除了16进制，也可以用10进制来表示这个符号，复选框打勾是82
        doc.replace(selection1.getSelectedText(), "\u0052", true, true);
        //doc.replace(selection1.getSelectedText(),String.valueOf(((char)82)), true, true);

        //复选框打叉
        TextSelection selection2 = doc.findString("symbol2",true,true);
        TextRange tr2 = selection2.getAsOneRange();
        tr2.getCharacterFormat().setFontName("Wingdings 2");
        //16进制复选框打叉是0053，10进制是83
        doc.replace(selection2.getSelectedText(), "\u0053", true, true);

        //复选框不勾选
        TextSelection selection3 = doc.findString("symbol3", true, true);
        TextRange tr3 = selection3.getAsOneRange();
        tr3.getCharacterFormat().setFontName("Wingdings 2");
        //16进制复选框不勾选是00A3，10进制是163
        doc.replace(selection3.getSelectedText(), "\u00A3", true, true);

        doc.saveToFile("勾选复选框.docx",FileFormat.Docx_2013);
    }
}