package apache_poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ApachePOIExcelRead {

    private static final String FILE_NAME = "/Users/sskim/Downloads/POI_Excel_Example.xlsx";

    public static void main(String[] args) {

        try {
            //엑셀파일 위치 읽어들임
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            //엑셀파일을 위치 넣어줌
            Workbook workbook = new XSSFWorkbook(excelFile);
            //0번째 sheet를 불러옴
            Sheet datatypeSheet = workbook.getSheetAt(0);
            //열 반복자 생성
            Iterator<Row> rowIterator = datatypeSheet.iterator();

            while (rowIterator.hasNext()) {

                //열이 hasNext() 끝날때까지 반복하면서 현재 열을 얻음.
                Row currentRow = rowIterator.next();
                //현재 열의 Cell 즉 엑셀 시트의 데이터 입력 부분 셀을 얻음.
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {
                    //해당 셀을 반복
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue()+"\t\t\t\t");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue()+"\t\t\t\t");
                    }
                }
                System.out.println();
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
