/*
 * Copyright (c) 2018, 吴汶泽 (wenzewoo@gmail.com).
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.wuwenze.poi.xlsx;

import com.wuwenze.poi.convert.WriteConverter;
import com.wuwenze.poi.exception.ExcelKitRuntimeException;
import com.wuwenze.poi.pojo.ExcelMapping;
import com.wuwenze.poi.pojo.ExcelProperty;
import com.wuwenze.poi.util.DateUtil;
import com.wuwenze.poi.util.POIUtil;
import com.wuwenze.poi.util.ValidatorUtil;
import java.text.ParseException;
import java.util.Date;
import java.util.List;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFDrawing;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

/**
 * @author wuwenze
 * @date 2018/5/1
 */
public class ExcelXlsxWriter {

  private final ExcelMapping mExcelMapping;
  private final Integer mMaxSheetRecords;

  public ExcelXlsxWriter(ExcelMapping excelMapping, Integer maxSheetRecords) {
    mExcelMapping = excelMapping;
    mMaxSheetRecords = maxSheetRecords;
  }

  /**
   * 构建xlsxWorkbook对象
   *
   * @param data 数据集
   * @param isTemplate 是否是导出模板
   * @return SXSSFWorkbook
   */
  public SXSSFWorkbook generateXlsxWorkbook(List<?> data, boolean isTemplate) {
    SXSSFWorkbook workbook = POIUtil.newSXSSFWorkbook();
    List<ExcelProperty> propertyList = mExcelMapping.getPropertyList();
    double sheetNo = Math.ceil(data.size() / (double) mMaxSheetRecords);
    for (int index = 0; index <= (sheetNo == 0.0 ? sheetNo : sheetNo - 1); index++) {
      String sheetName = mExcelMapping.getName() + (index == 0 ? "" : "_" + index);
      SXSSFSheet sheet = generateXlsxHeader(workbook, propertyList, sheetName, isTemplate);
      if (null != data && data.size() > 0) {
        int startNo = index * mMaxSheetRecords;
        int endNo = Math.min(startNo + mMaxSheetRecords, data.size());
        for (int i = startNo; i < endNo; i++) {
          SXSSFRow bodyRow = POIUtil.newSXSSFRow(sheet, i + 1 - startNo);
          for (int j = 0; j < propertyList.size(); j++) {
            SXSSFCell cell = POIUtil.newSXSSFCell(bodyRow, j);
            ExcelXlsxWriter.buildCellValueByExcelProperty(cell, data.get(i), propertyList.get(j));
          }
        }
      }
    }
    return workbook;
  }

  private SXSSFSheet generateXlsxHeader(SXSSFWorkbook workbook,
      List<ExcelProperty> propertyList,
      String sheetName, boolean isTemplate) {
    SXSSFDrawing sxssfDrawing = null;
    SXSSFSheet sheet = POIUtil.newSXSSFSheet(workbook, sheetName);
    SXSSFRow headerRow = POIUtil.newSXSSFRow(sheet, 0);

    for (int i = 0; i < propertyList.size(); i++) {
      ExcelProperty property = propertyList.get(i);
      SXSSFCell cell = POIUtil.newSXSSFCell(headerRow, i);
      POIUtil.setColumnWidth(sheet, i, property.getWidth(), property.getColumn());
      if (isTemplate) {
        // cell range
        POIUtil.setColumnCellRange(sheet, property.getOptions(), 1, mMaxSheetRecords, i, i);
        // cell comment.
        if (null == sxssfDrawing) {
          sxssfDrawing = sheet.createDrawingPatriarch();
        }
        if (!ValidatorUtil.isEmpty(property.getComment())) {
          // int col1, int row1, int col2, int row2
          Comment cellComment = sxssfDrawing.createCellComment(//
              new XSSFClientAnchor(0, 0, 0, 0, i, 0, i, 0));
          XSSFRichTextString xssfRichTextString = new XSSFRichTextString(
              property.getComment());
          Font commentFormatter = workbook.createFont();
          xssfRichTextString.applyFont(commentFormatter);
          cellComment.setString(xssfRichTextString);
          cell.setCellComment(cellComment);
        }
      }
      cell.setCellStyle(getHeaderCellStyle(workbook));
      String headerColumnValue = property.getColumn();
      if (isTemplate && null != property.getRequired() && property.getRequired()) {
        headerColumnValue = (headerColumnValue + "[*]");
      }
      cell.setCellValue(headerColumnValue);
    }
    return sheet;
  }

  private static void buildCellValueByExcelProperty(SXSSFCell cell, Object entity,
      ExcelProperty property) {
    Object cellValue;
    try {
      cellValue = BeanUtils.getProperty(entity, property.getName());
    } catch (Throwable e) {
      throw new ExcelKitRuntimeException(e);
    }
    if (null != cellValue) {
      String dateFormat = property.getDateFormat();
      if (!ValidatorUtil.isEmpty(dateFormat)) {
        if (cellValue instanceof Date) {
          cell.setCellValue(DateUtil.format(dateFormat, (Date) cellValue));
        } else if (cellValue instanceof String) {
          try {
            Date parse = DateUtil.ENGLISH_LOCAL_DF.parse((String) cellValue);
            cell.setCellValue(DateUtil.format(dateFormat, parse));
          } catch (ParseException e) {
            e.printStackTrace();
          }
          return;
        }
      }
      // writeConverterExp && writeConverter
      String writeConverterExp = property.getWriteConverterExp();
      WriteConverter writeConverter = property.getWriteConverter();
      if (!ValidatorUtil.isEmpty(writeConverterExp)) {
        try {
          cellValue = POIUtil.convertByExp(cellValue, writeConverterExp);
        } catch (Throwable e) {
          throw new ExcelKitRuntimeException(e);
        }
      } else if (null != writeConverter) {
        cell.setCellValue(writeConverter.convert(cellValue));
        return;
      }
      cell.setCellValue(String.valueOf(cellValue));
    }
  }

  private CellStyle mHeaderCellStyle = null;

  public CellStyle getHeaderCellStyle(SXSSFWorkbook wb) {
    if (null == mHeaderCellStyle) {
      mHeaderCellStyle = wb.createCellStyle();
      Font font = wb.createFont();
      mHeaderCellStyle.setFillForegroundColor((short) 12);
      mHeaderCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      mHeaderCellStyle.setBorderTop(BorderStyle.DOTTED);
      mHeaderCellStyle.setBorderRight(BorderStyle.DOTTED);
      mHeaderCellStyle.setBorderBottom(BorderStyle.DOTTED);
      mHeaderCellStyle.setBorderLeft(BorderStyle.DOTTED);
      mHeaderCellStyle.setAlignment(HorizontalAlignment.LEFT);// 对齐
      mHeaderCellStyle.setFillForegroundColor(HSSFColor.GREEN.index);
      mHeaderCellStyle.setFillBackgroundColor(HSSFColor.GREEN.index);
      font.setColor(HSSFColor.WHITE.index);
      // 应用标题字体到标题样式
      mHeaderCellStyle.setFont(font);
      //设置单元格文本形式
      DataFormat dataFormat = wb.createDataFormat();
      mHeaderCellStyle.setDataFormat(dataFormat.getFormat("@"));
    }
    return mHeaderCellStyle;
  }
}
