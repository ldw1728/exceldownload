package org.example.exceldownload.excel.service;


import lombok.extern.slf4j.Slf4j;
import org.example.exceldownload.excel.EnableExcel;
import org.example.exceldownload.excel.ExcelAttributes;
import org.springframework.stereotype.Component;

import jakarta.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@Component
public class ExcelDownload extends BaseExcelDownload {

    @Override
    protected void setResponseHeader(HttpServletResponse response, String fileName, Class<?> targetClass){
        if(targetClass != null){
            EnableExcel enableExcel = getEnableExcelAnno(targetClass);
            if(enableExcel != null){
                fileName = enableExcel.fileName() + "_" + String.valueOf(LocalDate.now()).replaceAll("-", "");
            }
        }
        super.setResponseHeader(response, fileName, null);
    }

    @Override
	protected List<String> getHeadData(Class<?> targetClass) {
		return getColNmsFrmCls(targetClass);
	}

	@Override
	protected List<String[]> getBodyData(List<?> dataList) {
		return getBodyDatasFrmCls(dataList);
	}

	@Override
	protected String getTitleData(Class<?> targetClass) {

        EnableExcel enableExcel = getEnableExcelAnno(targetClass);
        String title = "";

        if(enableExcel != null){
            title = enableExcel.title();
            if("".equals(title)){
                title = enableExcel.fileName();
            }
        }

        return  title;
	}

    /** class의 각 field data 추출 */
    private List<String[]> getBodyDatasFrmCls(List<?> dataList){

        List<String[]> resList = new ArrayList<>();

        if(!dataList.isEmpty()){
            Field[] fields = dataList.get(0).getClass().getDeclaredFields();

            for(Object obj : dataList){

                String[] fieldDatas = new String[fields.length];

                try {
                    for(int i = 0 ; i < fieldDatas.length ; i++){

                        // access to private field
                        fields[i].setAccessible(true);
                        Object fieldObj = fields[i].get(obj);
                        fieldDatas[i] = String.valueOf(fieldObj == null ? "" : fieldObj) ;

                        //log.debug(fieldDatas[i]);
                    }
                } catch (IllegalArgumentException | IllegalAccessException e) {
                    log.error(e.getMessage(), e);
                }

                resList.add(fieldDatas);
            }
        }
        return resList;
    }

    /** class의 각 field name 추출 */
    private List<String> getColNmsFrmCls(Class<?> targetClass){

        List<String> result = new ArrayList<>();

        for(Field field : targetClass.getDeclaredFields()){

            // access to private field
            field.setAccessible(true);
            
            String colNm = "";

            ExcelAttributes excelAttr = getExcelAttrAnno(targetClass, field);

            colNm = (excelAttr != null) ? excelAttr.colName() : field.getName();

            //log.debug(colNm);
            result.add(colNm);
        }

        return result;
    }

    /** get annotations */
    private EnableExcel getEnableExcelAnno(Class<?> pClass){
        return pClass.getAnnotation(EnableExcel.class);
    }

    private ExcelAttributes getExcelAttrAnno(Class<?> pClass, Field field){
        return getEnableExcelAnno(pClass) != null ? field.getAnnotation(ExcelAttributes.class) : null;
    }

	


   
    
}
