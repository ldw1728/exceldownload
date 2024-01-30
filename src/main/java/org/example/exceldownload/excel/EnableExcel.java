package org.example.exceldownload.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface EnableExcel {
    /**
     * 파일명
     * @return
     */
    String fileName();

    /**
     * 최상단에 노출되는 타이틀
     * @return
     */
    String title() default "";
}
