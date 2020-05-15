import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import FileUtilities.testUtilities;
import java.util.Iterator;

public class dataDrivenExcelTest {

    public static void main(String args[]) throws IOException, ClassNotFoundException {
        // file input stream argument for the XSSFWorkbook object creation
        dataDrivenExcelTest obj = new dataDrivenExcelTest();
        FileInputStream fis = null;
        File file = new File(dataDrivenExcelTest.class.getResource("datasets\\testcases.xlsx").getFile());
        String path = file.getAbsolutePath();
        path = URLDecoder.decode(path, StandardCharsets.UTF_8);
        System.out.println(path);
        try {
            fis = new FileInputStream(path);
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
        }

        XSSFWorkbook excelBook = new XSSFWorkbook(fis);

        System.out.println("total sheets"+ excelBook.getNumberOfSheets());
        XSSFSheet sheet = excelBook.getSheetAt(0);
        System.out.println("sheet name is : "+sheet.getSheetName());
        for(int i=0; i < excelBook.getNumberOfSheets();i++) {
            sheet = excelBook.getSheetAt(i);
            System.out.println("sheet name is : "+sheet.getSheetName());
            System.out.println("last row number : "+sheet.getLastRowNum());
        }

        String cellValue_0_0 = sheet.getRow(0).getCell(0).getStringCellValue();
        System.out.println("cellValue_0_0 : "+cellValue_0_0);
        Iterator<Row> rowIt = sheet.iterator();
        //Row row1 = row.next();

        while(rowIt.hasNext()) {
            Row rowNext = rowIt.next();
            System.out.println("Row number : "+rowNext.getRowNum());
            Iterator<Cell> cellIt = rowNext.cellIterator();
            System.out.println("Cell values : ");
            while (cellIt.hasNext()) {
                System.out.print(cellIt.next().getStringCellValue()+"  ");
            }
            System.out.println();
        }

        FileInputStream fis2 = new FileInputStream(path);
        obj.findStringInExcel("datasets\\testcases.xlsx","United Kingdom");
    }

    public  void findStringInExcel(String fileName, String text)
            throws IOException, ClassNotFoundException {
        int foundRowNo=0;
        int foundCellNo=0;

        FileInputStream fis = testUtilities.getStreamObjectForFile( this.getClass().getName(),fileName);
        XSSFWorkbook excel = new XSSFWorkbook(fis);
        int totalSheets = excel.getNumberOfSheets();
        for(int i=0;i< totalSheets;i++) {
            XSSFSheet sheet  = excel.getSheetAt(i);
            Iterator<Row> rowIt = sheet.iterator(); // sheet is a collection of rows
            while(rowIt.hasNext()) {
                Row row = rowIt.next();
                Iterator<Cell> cellIt = row.cellIterator(); // row is a collection of cells
                while(cellIt.hasNext()) {
                    Cell cell = cellIt.next();
                    if(cell.getCellType()== CellType.STRING) { // check the cell value if it is STRING value

                    } else if (cell.getCellType()== CellType.NUMERIC) { // check the cell value if it is NUMERIC value

                    }
                    if(cell.getStringCellValue().equals(text)){
                        foundRowNo = row.getRowNum();
                        foundCellNo = cell.getColumnIndex();
                        System.out.println("Found String : "+text+" at Sheet:"+sheet.getSheetName()
                                +" at Row:"+foundRowNo+", Column:"+foundCellNo+" in the Excel Sheet: "+fileName);
                        break;
                    }
                }
            }
        }

    }
}
