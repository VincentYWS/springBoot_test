package com.reimbursement.webapplication.controller;

import com.reimbursement.webapplication.entity.ChargeDetails;
import com.reimbursement.webapplication.entity.ReimbursementDetailTable;
import com.reimbursement.webapplication.service.ReimbursementService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URLDecoder;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

@Controller
public class HelloContriller {

    @Autowired
    ReimbursementService reimbursementService;

    /**
     * 顯示報銷表格界面
     * @return
     */
    @GetMapping("/reimbursement")
    public String Reimbursement(Model model) {
        return "index";
    }

    @PostMapping("/reimbursement")
    public String RembursementDetail(@RequestParam(value="filename") MultipartFile file,String box, Model model) {
        try{
            //批量导入。参数：文件名，文件。
            Map<String,List<ReimbursementDetailTable>> map =reimbursementService.getUploadFile(file);

            List<ChargeDetails> chargeDetailsList=reimbursementService.downloadFile(map,box);

    //        model.addAttribute("list",chargeDetailsList);

    //        List<ChargeDetails> chargeDetailsList=reimbursementService.getUploadFile_two(file);

            System.out.println(chargeDetailsList);
            model.addAttribute("list", chargeDetailsList);
            if(chargeDetailsList == null){
                return "error";
            }else{
                return "index";
            }

        }catch (Exception e){
            return "error";
        }

    }

    /**
     * 描述：下载外部案件导入模板
     * @throws Exception
     */
    @RequestMapping(value = "/downloadExcel")
    @ResponseBody
    public void downloadExcel(HttpServletResponse res, HttpServletRequest req, String name) throws Exception {

        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
        String dateString = formatter.format(new Date())+"/";
        String fileName = name;
        ServletOutputStream out;
        res.setContentType("multipart/form-data");
        res.setCharacterEncoding("UTF-8");
        res.setContentType("text/html");
        String filePath = new ClassPathResource("/templates/zip/" + dateString + fileName).getPath();
        String userAgent = req.getHeader("User-Agent");
        if (userAgent.contains("MSIE") || userAgent.contains("Trident")) {
            fileName = java.net.URLEncoder.encode(fileName, "UTF-8");
        } else {
            // 非IE浏览器的处理：
            fileName = new String((fileName).getBytes("UTF-8"), "ISO-8859-1");
        }
        filePath = URLDecoder.decode(filePath, "UTF-8");
        res.setHeader("Content-Disposition", "attachment;fileName=" + fileName);
            FileInputStream inputStream = new FileInputStream(filePath);
        out = res.getOutputStream();
        int b = 0;
        byte[] buffer = new byte[1024];
        while ((b = inputStream.read(buffer)) != -1) {
            // 4.写到输出流(out)中
            out.write(buffer, 0, b);
        }
        inputStream.close();

        if (out != null) {
            out.flush();
            out.close();
        }

    }
}