package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by lenovo on 2017/9/17.
 */
public class ExcelUtils {

    public static void main(String[] args){
        String path = "D:\\alltext\\excel\\1702.xlsx";
        List<List<String>> result = new ExcelUtils().readExcel(path);
        System.out.println(result.size());
        for (int i = 0;i < result.size();i++){
            List<String>model = result.get(i);
            for (int j = 0;j < model.size();j++){
                System.out.print(j + " : " + model.get(j) + "-->");
            }
            System.out.println();
        }
    }

    public List<List<String>> readExcel(String path) {
        List<List<String>> result = new ArrayList<>();
        try {
            InputStream inputStream = new FileInputStream(path);
            //HSSFWorkbook 标识整个 Excel
            Workbook hssfWorkbook = null;
            try {
                hssfWorkbook = new HSSFWorkbook(inputStream);
            }catch (Exception e){
                inputStream = new FileInputStream(path);
                hssfWorkbook = new XSSFWorkbook(inputStream);
            }
            int size = hssfWorkbook.getNumberOfSheets();

            for (int numberSheet = 0;numberSheet < size;numberSheet++){
                //HSSFSheet 标识某一页
                Sheet hssfSheet = hssfWorkbook.getSheetAt(numberSheet);
                if (hssfSheet == null){
                    continue;
                }
                for (int rowNumber = 1;rowNumber <= hssfSheet.getLastRowNum();rowNumber++){
                    Row hssfRow = hssfSheet.getRow(rowNumber);
                    int minRowCellIndex = hssfRow.getFirstCellNum();
                    int maxRowCellIndex = hssfRow.getLastCellNum();
                    List<String> rowList = new ArrayList<>();
                    for (int cellIndex = minRowCellIndex;cellIndex < maxRowCellIndex;cellIndex++){
                        Cell hssfCell = hssfRow.getCell(cellIndex);
                        if (hssfCell == null){
                            continue;
                        }
                        rowList.add(getStringValue(hssfCell));
                    }
                    result.add(rowList);
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return result;
    }

    private String getStringValue(Cell hssfCell) {
        switch (hssfCell.getCellType()){
            case Cell.CELL_TYPE_BOOLEAN:
                return hssfCell.getBooleanCellValue() ? "TRUE" : "FALSE";
            case Cell.CELL_TYPE_FORMULA:
                return hssfCell.getCellFormula();
            case Cell.CELL_TYPE_NUMERIC:
                hssfCell.setCellType(Cell.CELL_TYPE_STRING);
                return hssfCell.getStringCellValue();
            case Cell.CELL_TYPE_STRING:
                return hssfCell.getStringCellValue();
                default:
                    return "";
        }
    }
}
