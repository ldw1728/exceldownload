package org.example.exceldownload.excel.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.web.util.UriUtils;

import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;


@Slf4j
public abstract class BaseExcelDownload {

    /* abstract method */
    abstract protected String getTitleData(Class<?> targetClass);
    abstract protected List<String> getHeadData(Class<?> targetClass);
    abstract protected List<String[]> getBodyData(List<?> dataList);

    /***
     * 엑셀파일 다운로드
     * @param response
     * @param dataList
     * @param colNms
     * @throws IOException
    */
    public void downloadExcelFile(HttpServletResponse response, String fileName, List<String> colNms, List<?> dataList, Class<?> targetClass) throws IOException {

        SXSSFWorkbook workbook = null;

        if(dataList.isEmpty()){
            log.warn("EXCEL_DOWNLOAD : 출력가능한 엑셀 데이터가 존재하지 않습니다.");
        }

        try {
            setResponseHeader(response, fileName, targetClass);

            workbook = createWorkbook(dataList, colNms, fileName, targetClass);

            workbook.write(response.getOutputStream());

        } catch(Exception e) {
            e.printStackTrace();
        } finally {
            workbook.close();
        }

    }

    /**
     *  공통 workbook 생성
     * @param dataList
     * @param colNms
     * @return
     */
    protected SXSSFWorkbook createWorkbook(List<?> dataList, List<String> colNms, String fileName, Class<?> targetClass) {

        SXSSFWorkbook workbook = new SXSSFWorkbook();

        // 필드 갯수
        int fieldCnt    = targetClass.getDeclaredFields().length;

        // sheet 생성
        SXSSFSheet sheet = workbook.createSheet("sheet1");
        sheet.setZoom(85);

        //title row 생성
        String title = getTitleData(targetClass);
        if(title == null || "".equals(title)){
            title = fileName;
        }
        setTitleRow(sheet, title, fieldCnt, getTitleCellStyle(workbook));

        // head row 생성
        // 인자로 넘어온 컬럼명이 없을 경우 get메서드 호출
        List<String> headColNms = colNms != null ? colNms : getHeadData(targetClass);
        setHeadRow(sheet, headColNms, 1, getHeadCellStyle(workbook));

        // body row 생성
        List<String[]> bodyDatas = getBodyData(dataList);
        setBodyRows(sheet, bodyDatas, 2, getBodyCellStyle(workbook));

        // acceptAutoSizeTotCol(sheet, fieldCnt, headColNms);

        log.debug("EXCEL_DOWNLOAD : col = " + fieldCnt + ", row = " + dataList.size());

        return workbook;
    }

    /**
     * 공통 title row setting
     * @param sheet
     * @param title
     * @param fieldCnt
     * @param cellStyle
     */
    protected void setTitleRow(Sheet sheet, String title, int fieldCnt, CellStyle cellStyle){

        Row titleRow = sheet.createRow(0);
        
        titleRow.setHeight((short) 600);

        sheet.setDisplayGridlines(false);
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,fieldCnt-1));

        for (int i = 0 ; i < fieldCnt ; i++ ){
            
            Cell cell = titleRow.createCell(i);
            if(i == 0){
                cell.setCellValue(title);
            }
            cell.setCellStyle(cellStyle);
        }
    }
    
    /**
     * 공통 head row setting
     * @param sheet
     * @param colNms
     * @param rowNum
     * @param cellStyle
     */
    protected void setHeadRow(Sheet sheet, List<String> colNms, int rowNum, CellStyle cellStyle){
        Row headRow = sheet.createRow(rowNum);

        for(int i = 0 ; i < colNms.size() ; i++){
            Cell cell = headRow.createCell(i);
            cell.setCellValue(colNms.get(i));
            cell.setCellStyle(cellStyle);
            if(sheet.getColumnWidth(i) < colNms.get(i).getBytes().length * 250){
                sheet.setColumnWidth(i, (int) (colNms.get(i).getBytes().length * 250 * 1.5));
            }
        }
    }

    /**
     * 공통 body row setting
     * @param sheet
     * @param bodyDatas
     * @param rowNum
     * @param cellStyle
     */
    protected void setBodyRows(Sheet sheet, List<String[]> bodyDatas, int rowNum, CellStyle cellStyle) {
        for(String[] datas : bodyDatas){

            Row row = sheet.createRow(rowNum++);

            for(int i = 0 ; i < datas.length ; i++){
                String cellVal = datas[i].trim();
                Cell cell = row.createCell(i);
                cell.setCellValue(cellVal);
                cell.setCellStyle(cellStyle);

                sheet.setColumnWidth(i, 5000);

                /*
                // java.lang.IllegalArgumentException: The maximum column width for an individual cell is 255 characters.
                if(sheet.getColumnWidth(i) < cellVal.getBytes().length * 250){
                    // sheet.setColumnWidth(i, cellVal.getBytes().length * 260);
                    sheet.setColumnWidth(i, Math.min(cellVal.getBytes().length*256, sheet.getColumnWidth(i)+1024));
                }*/

            }
        }
    }

    /**
     * 공통 response header setting
     * @param response
     * @param fileName
     */
    protected void setResponseHeader(HttpServletResponse response, String fileName, Class<?> targetClass){

        if(fileName == null){
            fileName = "excelDownload";
        }
        if(!fileName.contains(".xls")){
            fileName += ".xlsx";
        }
        
        log.info("EXCEL_DOWNLOAD : " + fileName);
        fileName = UriUtils.encode(fileName, StandardCharsets.UTF_8);

        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName + "\"");
        response.setHeader(HttpHeaders.TRANSFER_ENCODING, "binary");
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }

    /** 셀 너비 자동설정 */
    protected void acceptAutoSizeTotCol(SXSSFSheet sheet, int colCnt, List<String> headColNms){

        for(int i = 0 ; i < colCnt ; i++){
            sheet.trackColumnForAutoSizing(i);
            sheet.autoSizeColumn(i);

            if(sheet.getColumnWidth(i) < headColNms.get(i).getBytes().length * 250){
                sheet.setColumnWidth(i, headColNms.get(i).getBytes().length * 280);
            }
        }
    }

    /**
     *  CellStyle setting
     */
    protected CellStyle getBodyCellStyle(SXSSFWorkbook workbook){

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        drawBorder(cellStyle);
        return cellStyle;
    }

    protected CellStyle getHeadCellStyle(SXSSFWorkbook workbook){

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        drawBorder(cellStyle);

        Font font = workbook.createFont();
        font.setFontHeightInPoints((short)10);
        cellStyle.setFont(font);

        return cellStyle;
    }

    protected CellStyle getTitleCellStyle(SXSSFWorkbook workbook){

        CellStyle cellStyle =  workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        drawBorder(cellStyle);

        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        
        return cellStyle;
    }

    protected void drawBorder(CellStyle cellStyle){
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
    }
    
}
