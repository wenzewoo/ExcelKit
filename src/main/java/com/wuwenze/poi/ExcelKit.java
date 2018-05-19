/*
 * Copyright (c) 2018, 吴汶泽 (wuwz@live.com).
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.wuwenze.poi;

import com.wuwenze.poi.exception.ExcelKitException;
import com.wuwenze.poi.factory.ExcelMappingFactory;
import com.wuwenze.poi.handler.ExcelReadHandler;
import com.wuwenze.poi.pojo.ExcelMapping;
import com.wuwenze.poi.util.Const;
import com.wuwenze.poi.util.POIUtil;
import com.wuwenze.poi.xlsx.ExcelXlsxReader;
import com.wuwenze.poi.xlsx.ExcelXlsxWriter;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

public class ExcelKit {
    private Class<?> mClass = null;
    private HttpServletResponse mResponse = null;
    private OutputStream mOutputStream = null;
    private Integer mMaxSheetRecords = 50000;
    private String mCurrentOptionMode = MODE_EXPORT;
    private final static String MODE_EXPORT = "$MODE_EXPORT$";
    private final static String MODE_BUILD = "$MODE_BUILD$";
    private final static String MODE_IMPORT = "$MODE_IMPORT$";

    /**
     * 使用此构造器来执行浏览器导出
     *
     * @param clazz    导出实体对象 (需 excel-mapping. xml 或 <@Excel + @ExcelField> 映射)
     * @param response 原生 response 对象, 用于响应浏览器下载
     * @return ExcelKit obj.
     * @see ExcelKit#downXlsx(List, boolean)
     */
    public static ExcelKit $Export(Class<?> clazz, HttpServletResponse response) {
        return new ExcelKit(clazz, response);
    }

    public void downXlsx(List<?> data, boolean isTemplate) {
        if (!mCurrentOptionMode.equals(MODE_EXPORT)) {
            throw new ExcelKitException("请使用com.wuwenze.poi.ExcelKit.$Export(Class<?> clazz, HttpServletResponse response)构造器初始化参数.");
        }
        try {
            ExcelMapping excelMapping = ExcelMappingFactory.get(mClass);
            ExcelXlsxWriter excelXlsxWriter = new ExcelXlsxWriter(excelMapping, mMaxSheetRecords);
            SXSSFWorkbook workbook = excelXlsxWriter.generateXlsxWorkbook(data, isTemplate);
            String fileName = isTemplate ? (excelMapping.getName() + "-导入模板.xlsx") : (excelMapping.getName() + "-导出结果.xlsx");
            POIUtil.download(workbook, mResponse, URLEncoder.encode(fileName, Const.ENCODING));
        } catch (Exception e) {
            throw new ExcelKitException(e);
        }
    }


    /**
     * 使用此构造器来执行构建文件流.
     *
     * @param clazz        导出实体对象 (需 excel-mapping. xml 或 <@Excel + @ExcelField> 映射)
     * @param outputStream 输出流
     * @return ExcelKit obj.
     * @see ExcelKit#writeXlsx(List, boolean)
     */
    public static ExcelKit $Builder(Class<?> clazz, OutputStream outputStream) {
        return new ExcelKit(clazz, outputStream);
    }

    public void writeXlsx(List<?> data, boolean isTemplate) {
        if (!mCurrentOptionMode.equals(MODE_BUILD)) {
            throw new ExcelKitException("请使用com.wuwenze.poi.ExcelKit.$Builder(Class<?> clazz, OutputStream outputStream)构造器初始化参数.");
        }
        ExcelMapping excelMapping = ExcelMappingFactory.get(mClass);
        ExcelXlsxWriter excelXlsxWriter = new ExcelXlsxWriter(excelMapping, mMaxSheetRecords);
        SXSSFWorkbook workbook = excelXlsxWriter.generateXlsxWorkbook(data, isTemplate);
        POIUtil.write(workbook, mOutputStream);
    }

    /**
     * 使用此构造器来执行Excel文件导入.
     *
     * @param clazz 导出实体对象 (需 excel-mapping. xml 或 <@Excel + @ExcelField> 映射)
     * @return ExcelKit obj.
     * @see ExcelKit#readXlsx(File, Integer, ExcelReadHandler)
     * @see ExcelKit#readXlsx(InputStream, Integer, ExcelReadHandler)
     * @see ExcelKit#readXlsx(File, ExcelReadHandler)
     * @see ExcelKit#readXlsx(InputStream, ExcelReadHandler)
     */
    public static ExcelKit $Import(Class<?> clazz) {
        return new ExcelKit(clazz);
    }


    public void readXlsx(File excelFile, ExcelReadHandler<?> excelReadHandler) {
        this.readXlsx(excelFile, -1, excelReadHandler);
    }

    public void readXlsx(File excelFile, Integer sheetIndex, ExcelReadHandler<?> excelReadHandler) {
        try {
            InputStream inputStream = new FileInputStream(excelFile);
            this.readXlsx(inputStream, sheetIndex, excelReadHandler);
        } catch (FileNotFoundException e) {
            throw new ExcelKitException(e);
        }
    }

    public void readXlsx(InputStream inputStream, ExcelReadHandler<?> excelReadHandler) {
        this.readXlsx(inputStream, -1, excelReadHandler);
    }

    public void readXlsx(InputStream inputStream, Integer sheetIndex, ExcelReadHandler<?> excelReadHandler) {
        if (!mCurrentOptionMode.equals(MODE_IMPORT)) {
            throw new ExcelKitException("请使用com.wuwenze.poi.ExcelKit.$Import(Class<?> clazz)构造器初始化参数.");
        }
        ExcelMapping excelMapping = ExcelMappingFactory.get(mClass);
        ExcelXlsxReader excelXlsxReader = new ExcelXlsxReader(mClass, excelMapping, excelReadHandler);
        if (sheetIndex >= 0) {
            excelXlsxReader.process(inputStream, sheetIndex);
            return;
        }
        excelXlsxReader.process(inputStream);
    }

    public ExcelKit setMaxSheetRecords(Integer mMaxSheetRecords) {
        this.mMaxSheetRecords = mMaxSheetRecords;
        return this;
    }

    protected ExcelKit() {
    }

    protected ExcelKit(Class<?> clazz) {
        this(clazz, null, null);
        this.mCurrentOptionMode = MODE_IMPORT;
    }

    protected ExcelKit(Class<?> clazz, OutputStream outputStream) {
        this(clazz, outputStream, null);
        this.mCurrentOptionMode = MODE_BUILD;
    }

    protected ExcelKit(Class<?> clazz, HttpServletResponse response) {
        this(clazz, null, response);
        this.mCurrentOptionMode = MODE_EXPORT;
    }

    protected ExcelKit(Class<?> clazz, OutputStream outputStream, HttpServletResponse response) {
        this.mClass = clazz;
        this.mOutputStream = outputStream;
        this.mResponse = response;
    }
}