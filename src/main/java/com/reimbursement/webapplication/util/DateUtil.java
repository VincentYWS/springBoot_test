package com.reimbursement.webapplication.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class DateUtil {


    //時間轉化類型
    public static String DateTypeConversion(String date) throws ParseException {
        Date endDate =new SimpleDateFormat("yyyy-MM-dd").parse(date);
        String endDatestr = new SimpleDateFormat("dd/MM/yyyy").format(endDate);
        return endDatestr;
    }

    public static  String DateaddOne(String dateStr) throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        Date startDate = sdf.parse(dateStr);
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(startDate);
        calendar.add(Calendar.DAY_OF_MONTH, +1);//+1今天的时间加一天
        startDate = calendar.getTime();
        return sdf.format(startDate);
    }

}