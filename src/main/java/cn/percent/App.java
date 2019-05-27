package cn.percent;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

/**
 * Hello world!
 *
 */
public class App 
{
    static HashSet<String> set0 = new HashSet<>();
    static HashSet<String> set1 = new HashSet<>();
    static HashSet<String> set2 = new HashSet<>();
    static HashSet<String> set3 = new HashSet<>();
    static HashSet<String> set4 = new HashSet<>();
    public static void main( String[] args )
    {
        //excel文件路径
        String excelPath = "E:\\sortingData.xlsx";
        try {
            //String encoding = "GBK";
            File excel = new File(excelPath);
            //判断文件是否存在
            if (excel.isFile() && excel.exists()) {
                //.是特殊字符，需要转义！！！！！
                String[] split = excel.getName().split("\\.");
                Workbook wb;
                //根据文件后缀（xls/xlsx）进行判断
                if ( "xls".equals(split[1])){
                    //文件流对象
                    FileInputStream fis = new FileInputStream(excel);
                    wb = new HSSFWorkbook(fis);
                }else if ("xlsx".equals(split[1])){
                    wb = new XSSFWorkbook(excel);
                }else {
                    System.out.println("文件类型错误!");
                    return;
                }

                //读取sheet 0
                Sheet sheet = wb.getSheetAt(0);
                //第一行是列名，所以不读
                int firstRowIndex = sheet.getFirstRowNum();
                int lastRowIndex = sheet.getLastRowNum();
//                System.out.println("firstRowIndex: "+firstRowIndex);
                /**第一行到最后一行遍历*/
                Set<String> allCell = new HashSet<>();
                for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {
                    Row row = sheet.getRow(rIndex);
                    int cellNum = row.getLastCellNum();
                    if(row != null){
                        int firstCellIndex = row.getFirstCellNum();
                        int lastCellIndex = row.getLastCellNum();
                        for (int i = firstCellIndex; i <= lastCellIndex-1; i++) {
                                Cell cell = row.getCell(i);
                            String cellStr = cell.toString();
                            allCell.add(cellStr);
                                if (cellStr.equals("")){

                                }else{
                                    switch (i){
                                        case 0:
                                            set0.add(cellStr);
                                            break;
                                        case 1:
                                            set1.add(cellStr);
                                            break;
                                        case 2:
                                            set2.add(cellStr);
                                            break;
                                        case 3:
                                            set3.add(cellStr);
                                            break;
                                        case 4:
                                            set4.add(cellStr);
                                            break;
                                        default:
                                            System.out.println("hava no cell input");
                                    }
                                }
                        }
                    }
                }

                List<Set> owner = new ArrayList<>();
                owner.add(set0);owner.add(set1);owner.add(set2);owner.add(set3);owner.add(set4);
                //-------------比较set之间的数据String存储元素，Set存储拥有元素的excel名
                HashMap<String, Set> matchList = new HashMap<>(16);
                for (int i = 0; i <= owner.size()-2; i++) {//0,1,2,3
                    Set ownerNameSet = new HashSet();
//                    获得List中第i个set
                    Set matcher = owner.get(i);
                    Iterator it = matcher.iterator();
                    while(it.hasNext()){
                        String value = it.next().toString();
                        boolean flag = false;
                    for (int j = i+1; j <= owner.size()-1; j++) {
                        Set elseMatcher = owner.get(j);
                        if(matchList.keySet().contains(value)){
                            continue;
                        }
                        if(elseMatcher.contains(value)){
                            ownerNameSet.add(ownerName(j));
                            ownerNameSet.add(ownerName(i));
                            flag = true;
                            }
                        }
                    if (flag == true){
                        if(matchList.entrySet().contains(value)){
                        }else{
                            matchList.put(value, ownerNameSet);
                            ownerNameSet = new HashSet<>();
                        }
                    }
                    }
                }
//                System.out.println(matchList);
                Iterator it = matchList.entrySet().iterator();
                while(it.hasNext()){
                    Map.Entry entry = (Map.Entry) it.next();
                    String key = entry.getKey().toString();
                    String value = entry.getValue().toString().replace(",", ";");
                    //System.out.printf("%-15s",key);
                    System.out.println(key+","+value);
                }

//                System.out.println("finish");
//                System.out.println(allCell);
            } else {
                System.out.println("找不到指定的文件");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 利用ArrayList的下标找出owner（匹配列）的名字
     * @param owner ArrayList的下标
     * @return ownerName 匹配列的名字
     */
    private static String ownerName(int owner) {
        String ownerName = "noName";
        switch (owner){
            case 0:
                ownerName = "14所";
                break;
            case 1:
                ownerName = "53";
                break;
            case 2:
                ownerName = "722";
                break;
            case 3:
                ownerName = "54ws";
                break;
            case 4:
                ownerName = "54dzz";
                break;
            default:
        }
        return ownerName;
    }
}
