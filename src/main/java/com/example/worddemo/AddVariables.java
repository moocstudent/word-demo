package com.example.worddemo;

import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.formatting.CharacterFormat;

public class AddVariables {
    public static void main(String[] args) {

        //创建Document
        Document document = new Document();

        //添加一个节
        Section section = document.addSection();

        //添加一个段落
        Paragraph paragraph = section.addParagraph();

        //设置文本格式
        CharacterFormat characterFormat = paragraph.getStyle().getCharacterFormat();
        characterFormat.setFontName("微软雅黑");
        characterFormat.setFontSize(14);

        //设置页边距
        section.getPageSetup().getMargins().setTop(80f);

        //添加变量到段落
        paragraph.appendField("($space1)", FieldType.Field_Doc_Variable);
        paragraph.appendText("是物质的永恒运动、变化的持续性、顺序性的表现，包含时刻和时段两个概念。\r\n");
        paragraph.appendField("($space2)", FieldType.Field_Doc_Variable);
        paragraph.appendText("是人类用以描述物质运动过程或事件发生过程的一个参数，确定");
        paragraph.appendField("($space3)", FieldType.Field_Doc_Variable);
        paragraph.appendText("，是靠不受外界影响的物质周期");
        paragraph.appendField("($space4)", FieldType.Field_Doc_Variable);
        paragraph.appendText("变化的");
        paragraph.appendField("($space5)", FieldType.Field_Doc_Variable);
        paragraph.appendText("。");

        //获取变量集合
        VariableCollection variableCollection = document.getVariables();

        //给添加的变量赋值
        variableCollection.add("($space1)", "时间");
        variableCollection.add("($space2)", "时间");
        variableCollection.add("($space3)", "时间");
        variableCollection.add("($space4)", "物质周期");
        //只有这个答案被替换为了括号
        variableCollection.add("($space5)", "(    )");

        //更新文档中的域
        document.isUpdateFields(true);

        //保存文档
        document.saveToFile("占位符替换答案.docx", FileFormat.Auto);
        document.dispose();
    }
}