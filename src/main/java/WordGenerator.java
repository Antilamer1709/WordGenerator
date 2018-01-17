import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;


import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

public class WordGenerator {
    public static void main(String[] args) throws Exception {
        createDocument();
    }

    private static void createDocument() throws Exception {
        //Blank Document
        XWPFDocument document = new XWPFDocument();
        FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Antilamer\\Documents\\POI\\create_table.docx"));

        initDocument(document);

        document.write(out);
        out.close();
        System.out.println("create_table.docx written successully");
    }

    private static void initDocument(XWPFDocument document) {
        initHeader(document);
        //create table
        XWPFTable table = document.createTable();
        initTableWidth(table);
        createTableRows(table);
    }

    private static void initHeader(XWPFDocument document) {
        //create paragraph
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);

        addHeaderRun(paragraph, "Tytuł programu: ", "KOKOSY");
        addHeaderRun(paragraph, "Data emisji: ", "12-01-2017");
        addHeaderRun(paragraph, "Tytuł: ", "Jakoś tam");
        addHeaderRun(paragraph, "Numer odcinka: ", "14");
        addHeaderRun(paragraph, "Numer serii: ", "12");
        addHeaderRun(paragraph, "Autor: ", "Dobry ziomek");

        addBorderBottom(paragraph);
    }

    private static void addHeaderRun(XWPFParagraph paragraph, String label, String text) {
        //Set Bold an Italic
        XWPFRun paragraphOneRunOne = paragraph.createRun();
        paragraphOneRunOne.setBold(true);
        paragraphOneRunOne.setText(label);

        XWPFRun paragraphOneRunTwo = paragraph.createRun();
        paragraphOneRunTwo.setText(text);
        paragraphOneRunTwo.addBreak();
    }

    private static void addBorderBottom(XWPFParagraph paragraph) {
        paragraph.setBorderBottom(Borders.SINGLE);
    }

    private static void initTableWidth(XWPFTable table) {
        CTTbl ctTable = table.getCTTbl();
        CTTblPr pr = ctTable.getTblPr();
        CTTblWidth tblW = pr.getTblW();
        tblW.setW(BigInteger.valueOf(5000));
        tblW.setType(STTblWidth.PCT);
        pr.setTblW(tblW);
        ctTable.setTblPr(pr);

        //align center
        CTJc jc = pr.addNewJc();
        jc.setVal(STJc.RIGHT);
        pr.setJc(jc);
    }

    private static void createTableRows(XWPFTable table) {
        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("col one, row one");
        tableRowOne.addNewTableCell().setText("col two, row one");
        tableRowOne.addNewTableCell().setText("col three, row one");

        //create second row
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("col one, row two");
        tableRowTwo.getCell(1).setText("col two, row two");
        tableRowTwo.getCell(2).setText("col three, row two");

        //create third row
        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText("col one, row three");
        tableRowThree.getCell(1).setText("col two, row three");
        tableRowThree.getCell(2).setText("col three, row three");
    }
}
