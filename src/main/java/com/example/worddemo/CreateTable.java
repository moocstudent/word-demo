package com.example.worddemo;

import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

import java.awt.*;

public class CreateTable {

    public static void main(String[] args) {

        //创建一个 Document 对象
        Document document = new Document();

        //添加一个节
        Section section = document.addSection();

        //定义表格数据
        String[] header = {"国家", "首都", "大陆", "国土面积", "人口数量"};
        String[][] data =
                {
                        new String[]{"玻利维亚", "拉巴斯", "南美", "1098575", "7300000"},
                        new String[]{"巴西", "巴西利亚", "南美", "8511196", "150400000"},
                        new String[]{"加拿大", "渥太华", "北美", "9976147", "26500000"},
                        new String[]{"智利”“圣地亚哥",  "南美", "756943", "13200000"},
                        new String[]{"哥伦比亚", "波哥大", "南美", "1138907", "33000000"},
                        new String[]{"古巴", "哈瓦那", "北美", "114524", "10600000"},
                        new String[]{"厄瓜多尔", "基多", "南美", "455502", "10600000"},
                        new String[]{"萨尔瓦多", "圣萨尔瓦多", "北美", "20865", "5300000"},
                        new String[]{ "圭亚那", "乔治城","南美", "214969", "800000"},

                };

        //添加表格
        Table table = section.addTable(true);
        table.resetCells(data.length + 1, header.length);

        //将第一行设置为表格标题
        TableRow row = table.getRows().get(0);
        row.isHeader(true);
        row.setHeight(20);
        row.setHeightType(TableRowHeightType.Exactly);
        row.getRowFormat().setBackColor(Color.gray);
        for (int i = 0; i < header.length; i++) {
            row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            Paragraph p = row.getCells().get(i).addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
            TextRange txtRange = p.appendText(header[i]);
            txtRange.getCharacterFormat().setBold(true);
        }

        //将数据添加到其余行
        for (int r = 0; r < data.length; r++) {
            TableRow dataRow = table.getRows().get(r + 1);
            dataRow.setHeight(25);
            dataRow.setHeightType(TableRowHeightType.Exactly);
            dataRow.getRowFormat().setBackColor(Color.white);
            for (int c = 0; c < data[r].length; c++) {
                dataRow.getCells().get(c).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                dataRow.getCells().get(c).addParagraph().appendText(data[r][c]);
            }
        }

        //设置单元格的背景颜色
        for (int j = 1; j < table.getRows().getCount(); j++) {
            if (j % 2 == 0) {
                TableRow row2 = table.getRows().get(j);
                for (int f = 0; f < row2.getCells().getCount(); f++) {
                    row2.getCells().get(f).getCellFormat().setBackColor(new Color(173, 216, 230));
                }
            }
        }

        //保存文件
        document.saveToFile("创建表格.docx", FileFormat.Docx_2013);
    }
}
