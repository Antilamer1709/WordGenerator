import java.io.File;
import java.io.FileOutputStream;
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
        createHeader(document);
        createAppearanceAgreements(document);
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
        paragraph.setBorderBottom(Borders.THICK);
    }

    private static void createAppearanceAgreements(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        addLabel(paragraph, "Zgody na wizerunek:");
        createAATable(document);
        paragraph.setBorderBottom(Borders.THICK);
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
        painHeaderRow(tableRowOne.getTableCells());
    }

    private static void painHeaderRow(List<XWPFTableCell> tableCells) {
        for (XWPFTableCell cell : tableCells) {
            cell.setColor("c9c9c9"); //todo create custom enum
        }
    }
}
