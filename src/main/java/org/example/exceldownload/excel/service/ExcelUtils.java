package org.example.exceldownload.excel.service;

import lombok.NonNull;

import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;


public class ExcelUtils {

    private static final BaseExcelDownload exclDwnld = new ExcelDownload();

    /**
     * 엑셀 다운로드
     * @param response
     * @param dataList
     * @param pClass
     * @throws IOException
     */
    public static void downloadExcelFile(@NonNull HttpServletResponse response, List<?> dataList, @NonNull Class<?> pClass) throws IOException{
        exclDwnld.downloadExcelFile(response, null, null, dataList, pClass);
    }

    /**
     * 엑셀 다운로드
     * @param response
     * @param fileName
     * @param dataList
     * @param colNms
     * @param pClass
     * @throws IOException
     */

    public static void downloadExcelFile(@NonNull HttpServletResponse response, String fileName, List<String> colNms, List<?> dataList, @NonNull Class<?> pClass) throws IOException{
        exclDwnld.downloadExcelFile(response, fileName, colNms, dataList, pClass);
    }
}
