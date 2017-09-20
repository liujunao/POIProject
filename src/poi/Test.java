package poi;

import java.util.Scanner;

/**
 * Created by lenovo on 2017/9/17.
 */
public class Test {

    public static void main(String[] args) throws Exception {
        Scanner scanner = new Scanner(System.in);
        FetchExcelToDB fetchExcelToDB = new FetchExcelToDB();
        FetchBDToExcel fetchBDToExcel = new FetchBDToExcel();
        System.out.println("将Excel文件转存入数据库，请输入1：");
        System.out.println("将数据库内容转存入Excel文件，请输入2：");
        System.out.println();
        int done = scanner.nextInt();
        System.out.println("输入文件存储路径 ： ");
        scanner = new Scanner(System.in);
        String path = scanner.nextLine();
        if (done == 1){
            System.out.println("输入存入到数据库时的数据库名 ： ");
            scanner = new Scanner(System.in);
            String database = scanner.nextLine();
            boolean flag = fetchExcelToDB.fetchExcelToMysql(path,database);
            if (flag){
                System.out.println("操作成功！");
            }else {
                System.out.println("操作失败！");
            }
        }else if (done == 2){
            System.out.println("输入存入到Excel表时的表名 ： ");
            scanner = new Scanner(System.in);
            String excelName = scanner.nextLine();
            boolean flag = fetchBDToExcel.fetchBDToExcel(excelName,path);
            if (flag){
                System.out.println("操作成功！");
            }else {
                System.out.println("操作失败！");
            }
        }

    }
}
