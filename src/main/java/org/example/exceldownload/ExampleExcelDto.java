package org.example.exceldownload;

import org.example.exceldownload.excel.EnableExcel;
import org.example.exceldownload.excel.ExcelAttributes;

@EnableExcel(fileName = "회원 리스트", title = "회원 리스트")
public class ExampleExcelDto {

    /*
    * 해당 클래스에서
    *  - 각 컬럼에 표시될 컬럼명 지정
    *  - 표시될 데이터 가공
    */

    /** 사용자ID */
    @ExcelAttributes(colName = "아이디")
    private String userId;

    /** 생년월일 */
    @ExcelAttributes(colName = "생년월일")
    private String brthdy;

    /** 성별 */
    @ExcelAttributes(colName = "성별")
    private String gender;

    /** 회원이름 */
    @ExcelAttributes(colName = "회원이름")
    private String userName;

    public ExampleExcelDto(String userId, String brthdy, String gender, String userName) {
        this.userId = userId;
        this.brthdy = brthdy;
        this.gender = gender;
        this.userName = userName;
    }
}
