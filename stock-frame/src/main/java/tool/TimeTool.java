package tool;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;

/**
 * @ClassName: TimeTool
 * @Author: taogang
 * @Date: 2020/12/21
 * @Description:
 */
public class TimeTool {

    public static final String DEFAULT = "yyyyMMddHHmmss";

    public static final String DATE_8 = "yyyyMMdd";

    public static final String YEAR = "year";

    public static final String MONTH = "month";

    public static final String WEEK = "week";

    public static final String DAY = "day";

    public static final String HOUR = "hour";

    /**
     * getNow
     * @param patten 时间格式
     * @return java.lang.String
     * @author taogang
     * @date 2020/12/21
     * @desc 根据给定的格式获取当前的时间字符串
     */
    public static String getNow(String patten){
        return DateTimeFormatter.ofPattern(patten).format(LocalDateTime.now());
    }

    /**
     * getNow
     * @return java.lang.String
     * @author taogang
     * @date 2020/12/21
     * @desc 使用默认方式给定时间字符串
     */
    public static String getNow(){
        return getNow(DEFAULT);
    }

    /**
     * getCycle
     * @param cycle 周期
     * @param patten 时间格式
     * @param num 负的是当前时间之前，正的是之后
     * @return java.lang.String
     * @author taogang
     * @date 2020/12/21
     * @desc 根据给定的周期获取前后的时间格式日期字符串
     */
    public static String getCycle(String cycle,String patten,int num){
        LocalDateTime now = LocalDateTime.now();
        switch (cycle){
            case YEAR:
                now = now.minusYears(num);
                break;
            case MONTH:
                now = now.minusMonths(num);
                break;
            case WEEK:
                now = now.minusWeeks(num);
                break;
            case DAY:
                now = now.minusDays(num);
                break;
            case HOUR:
                now = now.minusHours(num);
                break;
        }
        return  DateTimeFormatter.ofPattern(patten).format(now);
    }

    public static String getAroundYear(int num){
        return getCycle(YEAR,DEFAULT,num);
    }

    public static String getAroundMonth(int num){
        return getCycle(MONTH,DEFAULT,num);
    }

    public static String getAroundWeek(int num){
        return getCycle(WEEK,DEFAULT,num);
    }

    public static String getAroundDay(int num){
        return getCycle(DAY,DEFAULT,num);
    }

    public static String getAroundHour(int num){
        return getCycle(HOUR,DEFAULT,num);
    }

    /**
     * getCycleFirst
     * @param cycle
     * @param patten
     * @return java.lang.String
     * @author taogang
     * @date 2020/12/21
     * @desc 获取每个周期里面最开始的时间
     */
    public static String getCycleFirst(String cycle,String patten){
        LocalDateTime now = LocalDateTime.now();
        switch (cycle){
            case YEAR:
                now = now.withDayOfYear(1);
                break;
            case MONTH:
                now = now.withDayOfMonth(1);
                break;
            case WEEK:
                now = now.with(DayOfWeek.MONDAY);
                break;
            case DAY:
                now = LocalDateTime.of(LocalDate.now(), LocalTime.MIN);
                break;
        }
        return  DateTimeFormatter.ofPattern(patten).format(now);
    }

    public static String getYearFirst(){
        return getCycleFirst(YEAR,DEFAULT);
    }

    public static String getMonthFirst(){
        return getCycleFirst(MONTH,DEFAULT);
    }

    public static String getWeekFirst(){
        return getCycleFirst(WEEK,DEFAULT);
    }

    public static String getDayFirst(){
        return getCycleFirst(DAY,DEFAULT);
    }

}
