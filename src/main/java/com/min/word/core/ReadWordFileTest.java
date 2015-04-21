package com.min.word.core;

import java.io.FileInputStream;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class ReadWordFileTest {
	public static void main(String[] args) throws Exception{
		System.out.println("---------------- Read File Start ------------------");
		XWPFDocument document = new XWPFDocument(new FileInputStream("test.docx"));
		XWPFWordExtractor we = new XWPFWordExtractor(document);
		System.out.println(we.getText());
		System.out.println("---------------- Read File End ------------------");
	}
}
