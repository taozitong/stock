package tool;

import entity.ExcelLine;
import lombok.extern.flogger.Flogger;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * ExcelTool
 *
 * @author: taogang
 * @create: 2020-05-21 14:34
 * @description: excel的操作工具类
 */
@Slf4j
public class ExcelTool {
    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";
    
    /**
     * getAllExcelContext
     * @param filePath
     * @return: java.util.List<java.util.Map<java.lang.String,java.lang.Object>>
     * @author:taogang
     * @Date: 2020/5/21 15:42
     * @Description: 把excel的内容转成list里面放上列名和列内容的map方便控制
     */
    public static List<ExcelLine> getAllExcelContext(String filePath){
        List<ExcelLine> result = new ArrayList<ExcelLine>();
        Workbook workbook = getWorkbook(filePath);
        Sheet sheet = workbook.getSheetAt(0);
        //列名存储
        List<String> columnList = new ArrayList<String>();
        //有效列数
        int columNum = 0;
        
        //循环获取行内容
        for(int rowNum = 0;rowNum < sheet.getPhysicalNumberOfRows();rowNum ++){
            Row row = sheet.getRow(rowNum);
            if(null == row){
                continue;
            }
            ExcelLine line = null;
            if(rowNum != 0){
                line = new ExcelLine();
                line.setColumn(columnList);
                line.setData(new ArrayList<String>());
                line.setLineNum(rowNum);
            }
            
            //循环获取每列数据
            for(int cellNum = 0;cellNum < row.getPhysicalNumberOfCells();cellNum ++){
                Cell cell = row.getCell(cellNum);
                //设置列名
                if(rowNum == 0){
                    columnList.add(CellObjectToString(cell));
                    //当列是空的时候或者空字符串的时候
                    if(null == cell || "".equals(CellObjectToString(cell))){
                        break;
                    }
                    columNum++;
                }else{
                    List<String> dataList = line.getData();
                    if(cellNum >= columNum){
                        break;
                    }
                    
                    if(null == cell){
                        dataList.add("");
                    }else{
                        dataList.add(CellObjectToString(cell));
                    }
                
                }
            }
            if(rowNum != 0){
                result.add(line);
            }
        }
        
        
        try{
            workbook.close();
        }catch (IOException e){
            log.error("excel关闭失败的文件地址："+filePath);
            log.error("excel工作空间关闭失败，失败原因："+e.getMessage());
        }
        
        return result;
    }
    
    /**
     * CellObjectToString
     * @param cell
     * @return: java.lang.String
     * @author:taogang
     * @Date: 2020/5/21 16:01
     * @Description: 把每列的数据都转成String
     */
    private static String CellObjectToString(Cell cell) {
        if(cell == null){
            return "";
        }
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:   //数字
                Double doubleValue = cell.getNumericCellValue();
                // 格式化科学计数法，取一位整数
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING:    //字符串
                returnValue = cell.getStringCellValue();
                break;
            case BOOLEAN:   //布尔
                Boolean booleanValue = cell.getBooleanCellValue();
                returnValue = booleanValue.toString();
                break;
            case BLANK:     // 空值
                break;
            case FORMULA:   // 公式
                returnValue = cell.getCellFormula();
                break;
            case ERROR:     // 故障
                break;
            default:
                break;
        }
        return returnValue;
    }
    
    /**
     * getWorkbook
     * @param filePath 文件地址
     * @return: org.apache.poi.ss.usermodel.Workbook
     * @author:taogang
     * @Date: 2020/5/21 15:26
     * @Description: 获取excel文件的工作空间
     * 1、使用路径获取文件的后缀名和文件
     * 2、判断文件是否存在，不存在的话结束并打印错误日志
     * 3、解析excle并返回工作空间
     * 4、无论解析是否异常都关闭流和excel的工作空间
     */
    public static Workbook getWorkbook(String filePath) {
        //获取Excel后缀名
        String fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
        // 获取Excel文件
        File excelFile = new File(filePath);
        Workbook workbook = null;
        FileInputStream inputStream;
        if(!excelFile.exists()){
            log.error(filePath+"，这个excle文件不存在！");
            return workbook;
        }
        
        try{
            // 获取Excel工作簿
            inputStream = new FileInputStream(excelFile);
            //根据文件的后缀名区分是哪种excel
            if (fileType.equalsIgnoreCase(XLS)) {
                workbook = new HSSFWorkbook(inputStream);
            } else if (fileType.equalsIgnoreCase(XLSX)) {
                workbook = new XSSFWorkbook(inputStream);
            }
        }catch (Exception e){
            log.error("解析excel异常，错误信息："+e.getMessage());
            log.error("错误文件路径:"+filePath);
        }
        
        return workbook;
    }
    /**
     * WriteExcel
     * @param path  写入文件的路径
     * @param elList 写入的数据
     * @return: void
     * @author:taogang
     * @Date: 2020/5/21 17:25
     * @Description: 按照给定的数据和地址写入excel文件
     */
    public static void WriteExcel(String path,List<ExcelLine> elList){
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        List<String> column = elList.get(0).getColumn();
        for(int rowNum = 0;rowNum<= elList.size();rowNum++){
            Row row = sheet.createRow(rowNum);
            //设置头
            if(rowNum == 0){
                writeDataToRow(column,row);
            }else{
                for(int i=0;i< elList.size();i++){
                    writeDataToRow(elList.get(i).getData(),row);
                }
            }
        }
        
        // 以文件的形式输出工作簿对象
        FileOutputStream fileOut = null;
        try {
            File exportFile = new File(path);
            if (!exportFile.exists()) {
                exportFile.createNewFile();
            }
        
            fileOut = new FileOutputStream(path);
            workbook.write(fileOut);
            fileOut.flush();
        } catch (Exception e) {
            log.error("输出Excel时发生错误，错误原因：" + e.getMessage());
        } finally {
            try {
                if (null != fileOut) {
                    fileOut.close();
                }
                if (null != workbook) {
                    workbook.close();
                }
            } catch (IOException e) {
                log.error("关闭输出流时发生错误，错误原因：" + e.getMessage());
            }
        }
    
    }
    /**
     * setDefaultExcelHeardStyle
     * @param workbook
     * @return: org.apache.poi.ss.usermodel.CellStyle
     * @author:taogang
     * @Date: 2020/5/21 17:36
     * @Description: 设置默认头的样式
     */
    private static CellStyle setDefaultExcelHeardStyle(Workbook workbook){
        // 构建头单元格样式
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // 下边框
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex()); // 左边框
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex()); // 右边框
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex()); // 上边框
        //设置背景颜色
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //粗体字设置
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }
    
    /**
     * convertDataToRow
     * @param data
     * @param row
     * @return: void
     * @author:taogang
     * @Date: 2020/5/21 17:42
     * @Description: 写入内容到excel里面
     */
    private static void writeDataToRow(List<String> data, Row row){
        Cell cell;
        for(int cluNum = 0;cluNum < data.size();cluNum++){
            cell = row.createCell(cluNum + 1);
            cell.setCellValue(data.get(cluNum));
        }
    }

}
