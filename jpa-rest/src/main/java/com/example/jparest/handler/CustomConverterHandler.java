package com.example.jparest.handler;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.util.BooleanUtils;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.util.StringUtils;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

@Slf4j
public class CustomConverterHandler implements Converter<List<String >> {
    @Override
    public Class supportJavaTypeKey() {
        return ArrayList.class;
    }
    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }
    @Override
    public WriteCellData<?> convertToExcelData(List<String> value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
//        List<String > list = (List) value;

//        StringUtils.collectionToDelimitedString(list,",","[",)
        String s = value.toString();
        System.out.println(s);
        return new WriteCellData<>(s);
    }

//    @Override
//    public WriteCellData<?> convertToExcelData(Object value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
//
////        return Converter.super.convertToExcelData(s, contentProperty, globalConfiguration);
//    }
}