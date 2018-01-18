import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

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
        createHeader(document);
        createAppearanceAgreements(document);
//        createWithoutAgreements(document); //todo
        createExternalMaterials(document);
    }

    private static void createHeader(XWPFDocument document) {
        //create paragraph
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);

        addHeaderRun(paragraph, "Tytuł programu: ", "KOKOSY");
        addHeaderRun(paragraph, "Data emisji: ", "12-01-2017");
        addHeaderRun(paragraph, "Tytuł: ", "Jakoś tam");
        addHeaderRun(paragraph, "Numer odcinka: ", "14");
        addHeaderRun(paragraph, "Numer serii: ", "12");
        addHeaderRun(paragraph, "Autor: ", "Dobry ziomek");

        paragraph.setBorderBottom(Borders.THICK);
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

    private static void createAppearanceAgreements(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        addLabel(paragraph, "Zgody na wizerunek:");
        createAATable(document);
        document.createParagraph().setBorderBottom(Borders.THICK);
    }

    private static void addLabel(XWPFParagraph paragraph, String label) {
        //create paragraph
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun paragraphOneRunOne = paragraph.createRun();
        paragraphOneRunOne.setBold(true);
        paragraphOneRunOne.setText(label);
        paragraphOneRunOne.addBreak();
    }

    private static void createAATable(XWPFDocument document) {
        XWPFTable table = document.createTable();
        initTableWidth(table);
        createAATableRows(table);
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

    private static void createAATableRows(XWPFTable table) {
        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("Nazwisko i imię");
        tableRowOne.addNewTableCell().setText("Obostrzerzenia");
        tableRowOne.addNewTableCell().setText("Uwagi");
        paintHeaderRows(tableRowOne.getTableCells());
    }

    private static void paintHeaderRows(List<XWPFTableCell> tableCells) {
        for (XWPFTableCell cell : tableCells) {
            cell.setColor("d8d8d8"); //todo create custom enum
        }
    }

    private static void createExternalMaterials(XWPFDocument document) {
        createQuotationRights(document);
        //todo others
    }

    private static void createQuotationRights(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        addLabel(paragraph, "Materiały zewnętzne - prawo cytatu:");

        XWPFTable table = document.createTable();
        initTableWidth(table);
        createQuotationTableRows(table);
        initQTableColumnSize(table);
    }

    private static void createQuotationTableRows(XWPFTable table) {
        initQuotationHeader(table);
        initQuotationBody(table);
    }

    private static void initQuotationHeader(XWPFTable table) {
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("Nazwisko i imię");
        tableRowOne.addNewTableCell().setText("Uwagi");
        tableRowOne.addNewTableCell().setText("Tytuł");
        tableRowOne.addNewTableCell().setText("Producent");

        paintHeaderRows(tableRowOne.getTableCells());
        mergeCell(tableRowOne.getCell(1), BigInteger.valueOf(4));
        mergeCell(tableRowOne.getCell(2), BigInteger.valueOf(2));

        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("Reżyser");
        tableRowTwo.getCell(1).setText("Scenarysta");
        tableRowTwo.getCell(2).setText("D. produkcji");
        tableRowTwo.getCell(3).setText("Czas Trwania");
        tableRowTwo.addNewTableCell().setText("Licencja do");
        tableRowTwo.addNewTableCell().setText("TC początek");
        tableRowTwo.addNewTableCell().setText("TC koniec");
        tableRowTwo.addNewTableCell().setText("Właściciel praw");
        paintHeaderRows(tableRowTwo.getTableCells());
    }

    private static void mergeCell(XWPFTableCell cell, BigInteger value) {
        if (cell.getCTTc().getTcPr() == null) cell.getCTTc().addNewTcPr();
        if (cell.getCTTc().getTcPr().getGridSpan() == null) cell.getCTTc().getTcPr().addNewGridSpan();
        cell.getCTTc().getTcPr().getGridSpan().setVal(value);
    }

    private static void initQuotationBody(XWPFTable table) {
        XWPFTableRow tableRowOne = table.createRow();
        tableRowOne.getCell(0).setText("some");
        tableRowOne.getCell(1).setText("example ");
        tableRowOne.getCell(2).setText("test");
        tableRowOne.getCell(3).setText("text");

        mergeCell(tableRowOne.getCell(1), BigInteger.valueOf(4));
        mergeCell(tableRowOne.getCell(2), BigInteger.valueOf(2));

        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("class");
        tableRowTwo.getCell(1).setText("java ");
        tableRowTwo.getCell(2).setText("idea");
        tableRowTwo.getCell(3).setText("double");
        tableRowTwo.addNewTableCell().setText("long");
        tableRowTwo.addNewTableCell().setText("static");
        tableRowTwo.addNewTableCell().setText("final");
        tableRowTwo.addNewTableCell().setText("case");
    }

    private static void initQTableColumnSize(XWPFTable table) {
        setColumnSize(table.getRow(0).getCell(0), 900);
        setColumnSize(table.getRow(0).getCell(1), 5000);
        setColumnSize(table.getRow(0).getCell(2), 5000);
        setColumnSize(table.getRow(0).getCell(3), 1100);
        setColumnSize(table.getRow(1).getCell(7), 1100);
    }

    private static void setColumnSize(XWPFTableCell cell, int size) {
        CTTblWidth tblWidth = cell.getCTTc().addNewTcPr().addNewTcW();
        tblWidth.setW(BigInteger.valueOf(size));
        //STTblWidth.DXA is used to specify width in twentieths of a point.
        tblWidth.setType(STTblWidth.DXA);
    }
}
