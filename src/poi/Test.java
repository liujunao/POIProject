package poi;

/**
 * Created by lenovo on 2017/9/17.
 */
public class Test {

    public static void main(String[] args) throws Exception {
        Common common = new Common();
        String value = "jdbc:mysql://localhost:3306/hello_1?useSSL=true&serverTimezone=UTC";
        String key = "url";
        String path = "D:\\workCode\\POIProject\\src\\db.properties";
        common.revise(path,key,value);
        FetchBDToExcel fetchBDToExcel = new FetchBDToExcel();
        String file = "D:\\alltext\\hello_1.xls";
        fetchBDToExcel.fetchBDToExcel("hello_1","hellotable_1","D:\\alltext");
//        fetchBDToExcel.createExcel(file,"hellotable_1");

//        System.out.println(common.query("hellotable_1",null));
    }
}
