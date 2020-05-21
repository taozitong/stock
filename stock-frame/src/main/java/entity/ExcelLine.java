package entity;

import lombok.Data;

import java.util.List;

/**
 * ExcelLine
 *
 * @author: taogang
 * @create: 2020-05-21 16:21
 * @description: excel的每行数据
 */
@Data
public class ExcelLine {
    //列名
    private List<String> colomn;
    //行数据
    private List<String> data;
    //行号从零开始
    private int lineNum;
}
