package com.kun.sdk.office.test.example;


import com.dlw.architecture.office.enums.FileType;
import com.dlw.architecture.office.exception.OfficeException;

import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.util.concurrent.TimeUnit;

/**
 * @author dengliwen
 * @date 2020/6/23
 * @desc
 */
public class ResponseUtil {

    public static void initExcelResponse(HttpServletResponse response, String fileName, FileType.ExcelType excelType) throws OfficeException {
        response.reset();
        response.setCharacterEncoding("UTF-8");
        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        try {
            response.setHeader("Set-Cookie", "fileDownload=true; path=/");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String((fileName + "." + excelType.name().toLowerCase()).getBytes("utf-8"), "iso-8859-1"));
        } catch (UnsupportedEncodingException e) {
            throw new OfficeException("Unsupported encoding exception ",e);
        }
    }

    public static void initWordResponse(HttpServletResponse response, String fileName, FileType.wordType wordType) throws OfficeException {
        try {
            response.reset();
            response.setCharacterEncoding("UTF-8");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String((fileName + "." + wordType.name().toLowerCase()).getBytes(),
                            "iso-8859-1"));
            response.setContentType("application/octet-stream");
        }catch (Exception e) {
            throw new OfficeException("Unsupported encoding exception ",e);
        }
    }

    public static void initPdfResponse(HttpServletResponse response, String fileName) throws OfficeException {
        try {
            response.reset();
            response.setCharacterEncoding("UTF-8");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String((fileName + "." + FileType.PDF.name().toLowerCase()).getBytes(),
                            "iso-8859-1"));
            response.setContentType("application/pdf");
        }catch (Exception e) {
            throw new OfficeException("Unsupported encoding exception ",e);
        }
    }

}
