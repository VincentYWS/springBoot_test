package com.reimbursement.webapplication.service.impl;

import com.reimbursement.webapplication.entity.ChargeDetails;
import com.reimbursement.webapplication.entity.CostDetail;
import com.reimbursement.webapplication.entity.Reimbursement;
import com.reimbursement.webapplication.entity.ReimbursementDetailTable;
import com.reimbursement.webapplication.service.ReimbursementService;
import com.reimbursement.webapplication.util.DateUtil;
import com.reimbursement.webapplication.util.ExcelUtil;
import com.reimbursement.webapplication.util.ZipUtil;
import org.apache.poi.hssf.usermodel.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
@Service
public class ReimbursementServiceImpl implements ReimbursementService {

    private final String  XRL = "XRL";
    private final String  DLS = "DLS";

    /**
     *  獲取上傳的excel
     * @param mfile 文件
     */
    public Map<String,List<ReimbursementDetailTable>> getUploadFile(MultipartFile mfile) throws Exception {
        //文件是否爲空
        if(!mfile.isEmpty()){
            //文件名
            String name = mfile.getOriginalFilename();
            //獲取文件名後綴
            String suffix=name.substring(name.lastIndexOf(".")+1,name.length());

            // 根据版本选择创建Workbook的方式
            Workbook workbook = ExcelUtil.getWorkbook(mfile.getInputStream(),suffix);

            //當前excle子表數量
            int sum=workbook.getNumberOfSheets();

            //當前標題
            List<String> titles=new ArrayList<String>();

            //暫時先這樣寫，如後期需要用到數據庫的時候，把這塊作為系統參數使用(費用類型)
            List<String> costTypes=new ArrayList<String>();

            costTypes.add("TRAFFIC");
            costTypes.add("MEAL");
            costTypes.add("other");

            //暫時先這樣寫，如後期需要用到數據庫的時候，把這塊作為系統參數使用(錢類型)
            List<String> currencys=new ArrayList<String>();
            currencys.add("HKD");
            currencys.add("MOP");
            currencys.add("CNY");

            //是否提前出發
            final String advance_finally = "In Advance";

            //保存結果集
            List<ReimbursementDetailTable> ReimbursementDetailTableList=new ArrayList<ReimbursementDetailTable>();

            //先區分好相對應的項目
            Map<String,List<ReimbursementDetailTable>> map = new HashMap<String,List<ReimbursementDetailTable>>();

            String itemName = "";

            for (int i=0;i<sum;i++){

                //excel表項目名稱
                String[] sheetNameHead = workbook.getSheetName(i).split("-");

                if(itemName.equals("")){
                    itemName = sheetNameHead[0];
                }

                //費用明細表
                ReimbursementDetailTable reimbursementDetailTable = null;

                //做成區間數據
                //區間時間
                String initDate = "";
                String addDate = "";


                //excel表子表
                Sheet sheet=workbook.getSheetAt(i);
                boolean boo = true;

                int count = 0 ;
                for (Row row : sheet) {

                    //從第7行開始獲取
                    if(count < 7 ){
                        count ++;
                        continue;
                    }

                    if(titles.size() == 0){
                        for(int j = 1 ;j<row.getLastCellNum()-1;j++){
                            String title=ExcelUtil.getValue(sheet.getRow(count).getCell(j)).toString().replaceAll(" ","");
                            if(j == 1 && (title == null || title.equals(""))){
                                break;
                            }
                            titles.add(ExcelUtil.getValue(sheet.getRow(count).getCell(j)).toString().replaceAll(" ",""));
                            boo = false;
                        }
                        count ++;
                    }

                    if(!boo){
                        boo = true;
                        continue;
                    }

                    //通過excle表列表數據返回列表對象（導入表的對象類型）
                    Reimbursement entity=(Reimbursement)ExcelUtil.getDataList(Reimbursement.class,row,titles);

                    //保持列表數據
                    if(!Objects.isNull(entity.getDate()) && !entity.getDate().equals("")){

                        //獲取excel表類型
                        entity.setAllowance(ExcelUtil.getValue(sheet.getRow(5).getCell(2)).toString());
                        entity.setProject(ExcelUtil.getValue(sheet.getRow(2).getCell(2)).toString());
                        //項目名稱
                        entity.setProjectName(sheetNameHead[0]);
                        //每一次進來去看下期間是否滿足條件
                        for(ReimbursementDetailTable reimbursementDetailTable1:ReimbursementDetailTableList){

                            SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
                            long startTime = sdf.parse(reimbursementDetailTable1.getStartDate()).getTime(); //開始時間
                            long endTime = sdf.parse(reimbursementDetailTable1.getEndDate()).getTime(); //結束時間
                            long currentTime=sdf.parse(DateUtil.DateTypeConversion(entity.getDate())).getTime(); //本次時間

                            //算區間
                            if(currentTime >= startTime && currentTime <=  endTime && reimbursementDetailTable1.getProjectName().equals(entity.getProjectName()) ){
                                reimbursementDetailTable = reimbursementDetailTable1;
                                initDate = DateUtil.DateTypeConversion(entity.getDate());
                                addDate = DateUtil.DateTypeConversion(entity.getDate());
                                break;
                            }
                        }
                        //保存開始時間
                        String start = DateUtil.DateTypeConversion(entity.getDate());

                        if(!DLS.equals(entity.getProjectName()) && !"".equals(initDate) && !initDate.equals(start)) {
                            addDate = DateUtil.DateaddOne(initDate);
                        }

                        //entity.getProjectName().equals(reimbursementDetailTable.getProjectName()  //項目名稱是否一樣
                        if(entity.getAllowance().equals("MEAL") && ("".equals(initDate) || !start.equals(addDate))){//是否需要新創建區間

                            //初始化所有數據
                            reimbursementDetailTable = new ReimbursementDetailTable();
                            reimbursementDetailTable.setStartDate(DateUtil.DateTypeConversion(entity.getDate()));
                            reimbursementDetailTable.setProjectName(entity.getProjectName());
                            reimbursementDetailTable.setEndDate(DateUtil.DateTypeConversion(entity.getDate()));
                            reimbursementDetailTable.setTravelDate("9:00-18:00");
                            reimbursementDetailTable.setTravelReason(entity.getProject());
                            if(entity.getDescription().indexOf("(") != -1 || entity.getDescription().indexOf(")") != -1){
                                String advance=entity.getDescription().substring(entity.getDescription().indexOf("(")+1,entity.getDescription().indexOf(")"));
                                if(advance.toUpperCase().contains(advance_finally.toUpperCase())){
                                    reimbursementDetailTable.setAdvance(advance);
                                }
                            }else if(entity.getDescription().indexOf("）") != -1 || entity.getDescription().indexOf("（") != -1){
                                String advance=entity.getDescription().substring(entity.getDescription().indexOf("（")+1,entity.getDescription().indexOf("）"));
                                if(advance.toUpperCase().contains(advance_finally.toUpperCase())){
                                    reimbursementDetailTable.setAdvance(advance);
                                }
                            }
                            //初始化費用類型
                            for(String costType:costTypes){
                                for(String currency:currencys){
                                    CostDetail costDetail= new CostDetail();
                                    costDetail.setCostType(costType);
                                    costDetail.setCurrency(currency);
                                    if(entity.getAllowance().equals(costType) && currency.equals(entity.getCurrency())){
                                        costDetail.setMoney(entity.getAmount()+ "");
                                    }
                                    reimbursementDetailTable.getCostTypeList().add(costDetail);
                                }
                            }

                            if(!reimbursementDetailTable.getProjectName().equals(itemName)){
                                map.put(itemName,ReimbursementDetailTableList);
                                ReimbursementDetailTableList =new ArrayList<ReimbursementDetailTable>();
                                itemName = sheetNameHead[0];
                            }

                            ReimbursementDetailTableList.add(reimbursementDetailTable);

                        }else{
                            String dateStr = start;

                            int z = 0;
                            while (true){
                                if(Objects.isNull(reimbursementDetailTable)){
                                    for(ReimbursementDetailTable reimbursementDetailTable1:ReimbursementDetailTableList){
                                        if(dateStr.equals(reimbursementDetailTable1.getStartDate())){
                                            reimbursementDetailTable = reimbursementDetailTable1;
                                            break;
                                        }
                                    }
                                    dateStr = DateUtil.DateaddOne(dateStr);
                                }else{
                                    break;
                                }
                                if(z == 1000){
                                    break;
                                }
                                z++;
                            }
                            //項目名稱是否一樣
                            if(entity.getProjectName().equals(reimbursementDetailTable.getProjectName())){
                                for(CostDetail costDetail:reimbursementDetailTable.getCostTypeList()){
                                    if(!initDate.equals(start) && start.equals(addDate)){
                                        reimbursementDetailTable.setEndDate(DateUtil.DateTypeConversion(entity.getDate()));
                                    }
                                    if(entity.getAllowance().contains(costDetail.getCostType()) && entity.getCurrency().equals(costDetail.getCurrency())){
                                        if(Objects.isNull(costDetail.getMoney())){
                                            costDetail.setMoney("0");
                                        }

                                        costDetail.setMoney(Double.valueOf(costDetail.getMoney()) + entity.getAmount() + "");

                                    }

                                }
                            }
                        }
                        initDate = DateUtil.DateTypeConversion(entity.getDate());

                    }

                }

                //清空標題list
                titles.clear();
            }

            //保存當前項目結果集
            map.put(itemName,ReimbursementDetailTableList);

            return map;
        }else {
            return null;
        }

    }

    /*
    * 下載excel表
    * */
    public List<ChargeDetails> downloadFile(Map<String, List<ReimbursementDetailTable>> map,String box) {
        return newMultipleExcel_zip(map,box);
    }

    /**
     * 将結果集轉化為excel
     * @param map :返回結果集
     * @return chargeDetailsArrayList :返回視圖
     */
    public  List<ChargeDetails> newMultipleExcel_zip(Map<String, List<ReimbursementDetailTable>> map,String box) {
        if(map == null){
            return null;
        }
        //返回視圖顯示
        List<ChargeDetails> chargeDetailsArrayList=new ArrayList<ChargeDetails>();

        //获取所有匹配的文件（注：此處坑點，被坑了一個上午） //摸版路徑後期放數據庫
        ClassPathResource resource  = new ClassPathResource("templates/dome/template.xlsx");

        //獲取文件路徑
        String path=new ClassPathResource("templates").getPath();

        //當前時間
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
        String dateString = formatter.format(new Date());


        //有數據庫之後存放于數據庫
        //excel路徑
       final String url_excel=path+ "/excel/"+dateString+"/"+UUID.randomUUID()+"/";
        //zip路徑
       final String url_zip = path+"/zip/"+dateString+"/";

        //新建文件夾
        File file_excel = new File(url_excel);
        if (!file_excel.exists()) {
            file_excel.mkdirs();
        }
        File file_zip = new File(url_zip);
        if (!file_zip.exists()) {
            file_zip.mkdirs();
        }

        //文件名
        String file_name;

        try{
            //開始新增excel
            for(String str : map.keySet()){

                BufferedInputStream fs = new BufferedInputStream(resource.getInputStream());
                //读取excel模板
                XSSFWorkbook wb = new XSSFWorkbook(fs);
                //读取了模板内所有sheet内容
                XSSFSheet sheet = wb.getSheetAt(0);
                XSSFCell cell;

                List<ReimbursementDetailTable> list=map.get(str);

               if(list.get(0).getProjectName().equals(XRL)){
                   cell = sheet.getRow(2).getCell(2);
                   cell.setCellValue("香港");
               }else if(list.get(0).getProjectName().equals(DLS)){
                   cell = sheet.getRow(2).getCell(2);
                   cell.setCellValue("澳門");
               }

                //重新按照開始時間排序
                list.sort((s1, s2) -> {
                    SimpleDateFormat format = new SimpleDateFormat("dd/MM/yyyy");
                    try {
                        Date dt1 = format.parse(s1.getStartDate());
                        Date dt2 = format.parse(s2.getStartDate());
                        if (dt1.getTime() > dt2.getTime())
                        {
                            return 1;
                        }else if(dt1.getTime() < dt2.getTime()){
                            return -1;
                        }else{
                            return  0;
                        }
                    } catch (Exception e){
                        return 0;
                    }
                });



               file_name = "出差申请费用明细登记表_V2.1_"+list.get(0).getProjectName()+"_"+list.get(0).getTravelReason().substring(list.get(0).getTravelReason().indexOf("(")+1,list.get(0).getTravelReason().indexOf(")"))+".xlsx";

               int i = 0;

               //把數據寫入到excel表裡面 需要優化，暫時不知道這麼入手
               for (ReimbursementDetailTable entity:list){
                   int j = 0;

                   cell = sheet.getRow(7+i).getCell(1);
                   if(!Objects.isNull(entity.getAdvance())){
                       cell.setCellValue(DateUtil.DateaddOne(entity.getStartDate()));
                   }else{
                       cell.setCellValue(entity.getStartDate());
                   }

                   cell = sheet.getRow(7+i).getCell(2);
                   cell.setCellValue(entity.getEndDate());
                   cell = sheet.getRow(7+i).getCell(3);
                   cell.setCellValue(entity.getTravelDate());
                   cell = sheet.getRow(7+i).getCell(4);
                   if(entity.getTravelReason().indexOf("（") != -1){
                       cell.setCellValue(entity.getTravelReason().split("（")[0]);
                   }else{
                       cell.setCellValue(entity.getTravelReason().split("\\(")[0]);
                   }


                   //是否需要幫忙計算
                   if("true".equals(box)){
                       DecimalFormat df = new DecimalFormat("#.00");

                       for (CostDetail costDetail:entity.getCostTypeList()){
                           cell = sheet.getRow(7+i).getCell(5+j);
                           if(!Objects.isNull(costDetail.getMoney())){
                               cell.setCellValue(df.format(Double.valueOf(costDetail.getMoney())));
                           }else{
                               cell.setCellValue("");
                           }
                           j++;
                       }
                   }
                   i++;
               }

               //重新新建個excel表重新保存
                wb.write(new FileOutputStream(url_excel+file_name));

                //釋放wb
                wb.close();

            }
        }catch (Exception e){

            e.printStackTrace();
            return null;
        }

        //打包zip
        ZipUtil.fileToZip(url_excel,url_zip,"DetailRecord_"+dateString);

        //返回視圖
        ChargeDetails chargeDetails=new ChargeDetails();
        chargeDetails.setFile_url(url_zip+"DetailRecord_"+dateString+".zip");
        chargeDetails.setName("DetailRecord_"+dateString+".zip");
        chargeDetailsArrayList.add(chargeDetails);

        return chargeDetailsArrayList;
    }






}