package com.supcon;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

public class LoadFromStream {
    public static void main( String[] args ) {
        //加载文档
        Document doc = new Document();
        String docxPathStr = "D://workspace//java//wordexport//docs//废水异常排放申请表.docx";
        File docxFile = new File(docxPathStr);
        if (!docxFile.exists()) {//文件不存在
            System.out.println("文件不存在");
            return;
        }
        try {
            InputStream input = new FileInputStream(docxPathStr);
            doc.loadFromStream(input, FileFormat.Docx_2013);
            System.out.println("读取word文件成功");
            //doc.addSection();
            //要替换第一个出现的指定文本，只需在替换前调用setReplaceFirst方法来指定只替换第一个出现的指定文本
            //doc.setReplaceFirst(true);
            //调用方法用新文本替换原文本内容
            System.out.println("开始替换占位符");
            doc.replace("${APPLY_COMPANY}", "中控技术股份有限公司", false, true);
            doc.replace("${JOB_CONTENT}", "作业内容", false, true);
            doc.replace("${PHONE}", "13738048773", false, true);
            doc.replace("${FILLEDBY}", "陈宏杰", false, true);
            doc.replace("${APPLY_TIME}", "2022.12.22", false, true);
            doc.replace("${FILL_DATE}", "2022.12.22", false, true);
            //保存文档
            doc.saveToFile("C://Users//rico//Desktop//ReplaceAllText.docx", FileFormat.Docx_2013);
            System.out.println("保存成功");
            doc.dispose();
        } catch (FileNotFoundException e) {
            System.out.println("找不到文件");
            e.printStackTrace();
        }
    }
}
