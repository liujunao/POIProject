package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.util.List;
import java.util.Map;

/**
 * Created by lenovo on 2017/9/17.
 */

//dataTable == sheetName; database = fileName
public class FetchBDToExcel {

    Common common = new Common();
    Workbook workbook = null;

    //往 Excel 中写入
    public void fetchBDToExcel(String database,String dataTable,String path){

        String fileDir = path + "\\" + database + ".xls";
        createExcel(fileDir,dataTable);
        File file = new File(fileDir);
        try {
            //创建 workbook
            workbook = new HSSFWorkbook(new FileInputStream(file));
        } catch (IOException e) {
            e.printStackTrace();
        }
        //流
        FileOutputStream fileOutputStream = null;
        Sheet sheet = workbook.getSheet(dataTable);
        System.out.println(sheet);
        //获取表格的总行数
        int rowCount = sheet.getLastRowNum() + 1;
        //获取表头的列数
        int columnCount = sheet.getRow(0).getLastCellNum();
        List<Map<String,Object>>list = common.query(dataTable,null);
        //获取表头行对象
        Row titleRow = sheet.getRow(0);
        System.out.println(titleRow);
        if (titleRow != null){
            for (int rowId = 0;rowId < list.size();rowId++){
                Map<String,Object> map = list.get(rowId);
                Row row = sheet.createRow(rowId + 1);
                for (int columnIndex = 0;columnIndex < columnCount;columnIndex++){
                    String key = titleRow.getCell(columnIndex).toString();
                    Cell cell = row.createCell(columnIndex);
                    cell.setCellValue(map.get(key) == null ? null : map.get(key).toString());
                }
            }
        }
        try {
            fileOutputStream = new FileOutputStream(fileDir);
            workbook.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                fileOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    //判断文件的 Sheet 是否存在
    public boolean sheetExist(String fileDir,String sheetName){
        boolean flag = false;
        File file = new File(fileDir);
        if (file.exists()){
            //创建workbook
            try {
                workbook = new HSSFWorkbook(new FileInputStream(file));
                //添加 WorkSheet（不添加sheet时，生成的xls文件打开时会报错）
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet != null){
                    flag = true;
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return flag;
    }

    //创建新的 Excel
    public void createExcel(String fileDir,String sheetName){
        workbook = new HSSFWorkbook();
        //添加 Worksheet
        Sheet sheet = workbook.createSheet(sheetName);
        FileOutputStream fileOutputStream = null;

        List<String> list = common.queryColumnName(sheetName);
        //添加表头
        Row row = workbook.getSheet(sheetName).createRow(0);
        for (int i = 0;i < list.size();i++){
            Cell cell = row.createCell(i);
            cell.setCellValue(String.valueOf(list.get(i) != null ? list.get(i) : ""));
        }
        try {
            if (common.fileExist(fileDir)){
                common.deleteExcel(fileDir);
            }
            fileOutputStream = new FileOutputStream(fileDir);
            workbook.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                fileOutputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
