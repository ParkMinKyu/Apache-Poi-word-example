/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */
package com.min.word.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.net.URL;
import java.net.URLConnection;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STUnderline;

public class MakeWordFileTest {
	public static void main(String[] args) throws Exception{
		String fileName = "test.docx";
		System.out.println("---------- Word Create Start ------------");
		//기본 빈 word 파일 생성
		XWPFDocument document = new XWPFDocument();
		FileOutputStream out = new FileOutputStream(new File(fileName));
		System.out.println("---------- Create Blank Success ------------");
		
		//Paragraph 생성
		XWPFParagraph paragraph = document.createParagraph();
		System.out.println("---------- Create Paragraph Success ------------");
		
		//border 적용
		paragraph.setBorderBottom(Borders.BASIC_BLACK_DASHES);
		paragraph.setBorderLeft(Borders.BASIC_BLACK_DASHES);
		paragraph.setBorderRight(Borders.BASIC_BLACK_DASHES);
		paragraph.setBorderTop(Borders.BASIC_BLACK_DASHES);
		System.out.println("---------- Create Border Success ------------");
		
		XWPFRun run = paragraph.createRun();
		run.setText("At tutorialspoint.com, we strive hard to " +
				   "provide quality tutorials for self-learning " +
				   "purpose in the domains of Academics, Information " +
				   "Technology, Management and Computer Programming Languages.");
		System.out.println("---------- Text Write to File ------------");
		
		//Table 생성
		XWPFTable table = document.createTable();
		//row추가
		XWPFTableRow rowOne = table.getRow(0);
		rowOne.getCell(0).setText("Col One, Row One");
		rowOne.addNewTableCell().setText("Col Tow, Row One");
		rowOne.addNewTableCell().setText("Col Three, Row One");
		//row추가
		XWPFTableRow rowTow = table.createRow();
		rowTow.getCell(0).setText("Col One, Row Tow");
		rowTow.getCell(1).setText("Col Tow, Row Tow");
		rowTow.getCell(2).setText("Col Three, Row Tow");
		//row추가
		XWPFTableRow rowThree = table.createRow();
		rowThree.getCell(0).setText("Col One, Row Three");
		rowThree.getCell(1).setText("Col Tow, Row Three");
		rowThree.getCell(2).setText("Col Three, Row Three");
		System.out.println("---------- Create Table Success ------------");
		
		//Add Image
		XWPFParagraph imageParagraph = document.createParagraph();
		XWPFRun imageRun = imageParagraph.createRun();
		imageRun.addPicture(new FileInputStream("test.png"), XWPFDocument.PICTURE_TYPE_PNG,"test.png", Units.toEMU(300), Units.toEMU(300));
		System.out.println("---------- Create Image Success ------------");
		
		//Hyperlink
		XWPFParagraph hyperlink = document.createParagraph();
		String id = hyperlink.getDocument().getPackagePart().addExternalRelationship("http://niee.kr", XWPFRelation.HYPERLINK.getRelation()).getId();
		CTR ctr = CTR.Factory.newInstance();
		CTHyperlink ctHyperlink = hyperlink.getCTP().addNewHyperlink();
		ctHyperlink.setId(id);
		
		CTText ctText = CTText.Factory.newInstance();
		ctText.setStringValue("Hyper-Link TEST");
		ctr.setTArray(new CTText[]{ctText});
		
		//설정 하이퍼링크 스타일
		CTColor color = CTColor.Factory.newInstance();
		color.setVal("0000FF");
		CTRPr ctrPr = ctr.addNewRPr();
		ctrPr.setColor(color);
		ctrPr.addNewU().setVal(STUnderline.SINGLE);
		
		//글꼴 설정
		CTFonts fonts = ctrPr.isSetRFonts() ? ctrPr.getRFonts() : ctrPr.addNewRFonts();
		fonts.setAscii("마이크로소프트 雅黑");
		fonts.setEastAsia("마이크로소프트 雅黑");
		fonts.setHAnsi("마이크로소프트 雅黑");

		//글꼴 크기 설정
		CTHpsMeasure sz = ctrPr.isSetSz() ? ctrPr.getSz() : ctrPr.addNewSz();
		sz.setVal(new BigInteger("24"));
		ctHyperlink.setRArray(new CTR[]{ctr});
		hyperlink.setAlignment(ParagraphAlignment.LEFT);
		hyperlink.setVerticalAlignment(TextAlignment.CENTER);
		System.out.println("---------- Create Hyperlink Success ------------");
		
		//Font style
		XWPFParagraph fontStyle = document.createParagraph();
		
		//set Bold an Italic
		XWPFRun boldAnItalic = fontStyle.createRun();
		boldAnItalic.setBold(true);
		boldAnItalic.setItalic(true);
		boldAnItalic.setText("Bold an Italic");
		boldAnItalic.addBreak();
		
		//set Text Position
		XWPFRun textPosition = fontStyle.createRun();
		textPosition.setText("Set Text Position");
		textPosition.setTextPosition(100);
		
		//Set Strike through and font Size and Subscript
		XWPFRun otherStyle = fontStyle.createRun();
		otherStyle.setStrike(true);
		otherStyle.setFontSize(20);
		otherStyle.setSubscript(VerticalAlign.SUBSCRIPT);
		otherStyle.setText(" Set Strike through and font Size and Subscript");
		System.out.println("---------- Set Font Style ------------");
		
		//Set Alignment Paragraph
		XWPFParagraph alignment = document.createParagraph();
		//Alignment to Right
		alignment.setAlignment(ParagraphAlignment.RIGHT);
		
		XWPFRun alignRight = alignment.createRun();
		alignRight.setText("At tutorialspoint.com, we strive hard to " +
				   "provide quality tutorials for self-learning " +
				   "purpose in the domains of Academics, Information " +
				   "Technology, Management and Computer Programming " +
				   "Languages.");
		
		//Alignment to Center
		alignment = document.createParagraph();
		//Alignment to Right
		alignment.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun alignCenter = alignment.createRun();
		alignCenter.setText("The endeavour started by Mohtashim, an AMU " +
				   "alumni, who is the founder and the managing director " +
				   "of Tutorials Point (I) Pvt. Ltd. He came up with the " +
				   "website tutorialspoint.com in year 2006 with the help" +
				   "of handpicked freelancers, with an array of tutorials" +
				   " for computer programming languages. ");
		System.out.println("---------- Set Alignment ------------");
		
		//word 파일 저장
		document.write(out);
		out.close();
		System.out.println("---------- Save File Name : " + fileName + " ------------");
		System.out.println("---------- Word Create End ------------");
	}
}
