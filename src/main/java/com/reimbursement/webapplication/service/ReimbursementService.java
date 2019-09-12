package com.reimbursement.webapplication.service;

import com.reimbursement.webapplication.entity.ChargeDetails;
import com.reimbursement.webapplication.entity.ReimbursementDetailTable;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.text.ParseException;
import java.util.List;
import java.util.Map;

public interface ReimbursementService {

    /**
     *  獲取上傳的excel
     */
    Map<String,List<ReimbursementDetailTable>> getUploadFile(MultipartFile file) throws  Exception;

    /**
     * 下載excel文件
     * */
    List<ChargeDetails> downloadFile(Map<String, List<ReimbursementDetailTable>> map,String box) throws IOException, InvalidFormatException, ParseException;

}
