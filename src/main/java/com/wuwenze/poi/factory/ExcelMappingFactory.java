/*
 * Copyright (c) 2018, 吴汶泽 (wuwz@live.com).
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

package com.wuwenze.poi.factory;

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;
import com.google.common.collect.Lists;

import com.wuwenze.poi.annotation.Excel;
import com.wuwenze.poi.annotation.ExcelField;
import com.wuwenze.poi.config.Options;
import com.wuwenze.poi.convert.ReadConverter;
import com.wuwenze.poi.convert.WriteConverter;
import com.wuwenze.poi.exception.ExcelKitAnnotationAnalyzeException;
import com.wuwenze.poi.exception.ExcelKitConfigAnalyzeFaildException;
import com.wuwenze.poi.exception.ExcelKitConfigFileNotFoundException;
import com.wuwenze.poi.exception.ExcelKitXmlAnalyzeException;
import com.wuwenze.poi.pojo.ExcelMapping;
import com.wuwenze.poi.pojo.ExcelProperty;
import com.wuwenze.poi.util.PathUtil;
import com.wuwenze.poi.util.ReflectionUtil;
import com.wuwenze.poi.util.ValidatorUtil;
import com.wuwenze.poi.validator.Validator;

import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.File;
import java.lang.reflect.Field;
import java.util.Iterator;
import java.util.List;

/**
 * @author wuwenze
 * @date 2018/5/1
 */
public class ExcelMappingFactory {

    private ExcelMappingFactory() {
    }

    private final static LoadingCache<Class<?>, ExcelMapping> mExcelMappingLoadingCache =
            CacheBuilder.newBuilder()
                    .maximumSize(100)
                    .build(new CacheLoader<Class<?>, ExcelMapping>() {
                        @Override
                        public ExcelMapping load(Class<?> key) {
                            return loadExcelMappingByClass(key);
                        }
                    });
    private final static List<String> mClazzFields = Lists
            .newArrayList("options", "writeConverter", "readConverter", "validator");
    private final static List<String> mRequeridAttrs = Lists.newArrayList("name");

    /**
     * 获取指定实体的Excel映射信息
     *
     * @return ExcelMapping
     */
    public static ExcelMapping get(Class<?> clazz) {
        try {
            return mExcelMappingLoadingCache.get(clazz);
        } catch (Exception e) {
            throw new ExcelKitConfigAnalyzeFaildException(e);
        }
    }

    private static ExcelMapping loadExcelMappingByClass(Class<?> clazz) {
        // 1. 从配置文件加载 (classpath:excel-mapping/className.xml)
        ExcelMapping excelMapping = null;
        boolean xmlConfigFileNotFound = false;
        String loadExcelMappingFailedMessage = null;
        try {
            excelMapping = loadExcelMappingByXml(clazz.getName());
        } catch (Exception e) {
            xmlConfigFileNotFound = e instanceof ExcelKitConfigFileNotFoundException;
            loadExcelMappingFailedMessage = e.getMessage();
        }
        // 2. 从注解加载配置信息 (当配置文件未找到时)
        if (null == excelMapping && xmlConfigFileNotFound) {
            try {
                excelMapping = loadExcelMappingByAnnotation(clazz);
            } catch (Exception e) {
                loadExcelMappingFailedMessage = e.getMessage();
            }
        }
        // 3. 加载配置信息失败.
        if (null == excelMapping && null != loadExcelMappingFailedMessage) {
            throw new ExcelKitConfigAnalyzeFaildException(loadExcelMappingFailedMessage);
        }
        return excelMapping;
    }

    private static ExcelMapping loadExcelMappingByAnnotation(Class<?> clazz)
            throws IllegalAccessException, InstantiationException {
        ExcelMapping excelMapping = new ExcelMapping();
        Excel excel = clazz.getAnnotation(Excel.class);
        if (null == excel) {
            throw new ExcelKitAnnotationAnalyzeException(
                    "[" + clazz.getName() + "] @Excel annotations not found.");
        }
        excelMapping.setName(excel.value());
        ExcelProperty excelMappingProperty;
        Field[] fields = clazz.getDeclaredFields();
        List<ExcelProperty> propertyList = Lists.newArrayList();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (null != excelField) {
                Options options = excelField.options().equals(ExcelField.DefaultAnnotation.class) ? null
                        : excelField.options().newInstance();
                WriteConverter writeConverter =
                        excelField.writeConverter().equals(ExcelField.DefaultAnnotation.class) ? null
                                : excelField.writeConverter().newInstance();
                ReadConverter readConverter =
                        excelField.readConverter().equals(ExcelField.DefaultAnnotation.class) ? null
                                : excelField.readConverter().newInstance();
                Validator validator =
                        excelField.validator().equals(ExcelField.DefaultAnnotation.class) ? null
                                : excelField.validator().newInstance();
                excelMappingProperty = ExcelProperty.builder()
                        .name(ValidatorUtil.isEmpty(excelField.name()) ? field.getName() : excelField.name())
                        .required(excelField.required())
                        .column(ValidatorUtil.isEmail(excelField.value()) ? field.getName() : excelField.value())
                        .comment(excelField.comment())
                        .maxLength(excelField.maxLength())
                        .width(excelField.width())
                        .dateFormat(excelField.dateFormat())
                        .options(options)
                        .writeConverterExp(excelField.writeConverterExp())
                        .writeConverter(writeConverter)
                        .readConverterExp(excelField.readConverterExp())
                        .readConverter(readConverter)
                        .regularExp(excelField.regularExp())
                        .regularExpMessage(excelField.regularExpMessage())
                        .validator(validator)
                        .build();
                propertyList.add(excelMappingProperty);
            }
        }
        if (propertyList.isEmpty()) {
            throw new ExcelKitAnnotationAnalyzeException(
                    "[" + clazz.getName() + "] @ExcelField annotations not found.");
        }
        excelMapping.setPropertyList(propertyList);
        return excelMapping;
    }

    private static ExcelMapping loadExcelMappingByXml(String clazzName) throws Exception {
        ExcelMapping excelMapping = new ExcelMapping();
        File config = PathUtil.getFileByClasspath(String.format("excel-mapping/%s.xml", clazzName));
        String configFile = "classpath:excel-mapping/" + config.getName();
        if (!config.exists()) {
            throw new ExcelKitConfigFileNotFoundException(
                    "[" + configFile + "] not found.");
        }
        SAXReader reader = new SAXReader();
        Document document = reader.read(config);
        Element rootElement = document.getRootElement();
        if (!"excel-mapping".equals(rootElement.getName())) {
            throw new ExcelKitXmlAnalyzeException(
                    "[" + configFile + "] <excel-mapping /> not found.");
        }
        Attribute nameAttr = rootElement.attribute("name");
        if (null == nameAttr) {
            throw new ExcelKitXmlAnalyzeException(
                    "[" + configFile + "] <excel-mapping> attribute \"name\"  not found.");
        }
        excelMapping.setName(nameAttr.getValue());
        List<ExcelProperty> propertyList = Lists.newArrayList();
        Iterator<Element> elementIterator = rootElement.elementIterator();
        while (elementIterator.hasNext()) {
            Element element = elementIterator.next();
            if ("property".equals(element.getName())) {
                List<Attribute> attributes = element.attributes();
                checkXmlPropertyRequiredAttr(configFile, attributes);

                ExcelProperty excelMappingProperty = null;
                for (Attribute attribute : attributes) {
                    if (null == excelMappingProperty) {
                        excelMappingProperty = new ExcelProperty();
                    }
                    String name = attribute.getName();
                    String value = attribute.getValue();
                    ReflectionUtil.setProperty(excelMappingProperty, name, validAndGetPropertyValue(configFile, name, value));
                }
                if (null != excelMappingProperty) {
                    propertyList.add(excelMappingProperty);
                }
            }
        }
        if (propertyList.isEmpty()) {
            throw new ExcelKitXmlAnalyzeException(
                    "[" + configFile + "] <property /> not found.");
        }
        excelMapping.setPropertyList(propertyList);
        return excelMapping;
    }

    private static void checkXmlPropertyRequiredAttr(String configFile, List<Attribute> attributes) {
        Integer containsCount = 0;
        for (Attribute attr : attributes) {
            if (mRequeridAttrs.contains(attr.getName())) {
                containsCount++;
            }
        }
        if (containsCount != mRequeridAttrs.size()) {
            throw new ExcelKitXmlAnalyzeException(
                    "[" + configFile + "] <property /> missing required attributes: " + mRequeridAttrs
                            .toString());
        }
    }

    private static Object validAndGetPropertyValue(String configFile, String name, String value) {
        String messageTemplate = String
                .format("[%s] <property %s=\"%s\"/> Analyze failed: ", configFile, name, value);
        if (mClazzFields.contains(name)) {
            try {
                return Class.forName(value).newInstance();
            } catch (Exception e) {
                throw new ExcelKitXmlAnalyzeException(messageTemplate + e.getMessage());
            }
        }
        if ("writeConverterExp".equals(name) || "readConverterExp".equals(name)) {
            // 1=男,2=女 or 男=1,女=2
            for (String item : value.split(",")) {
                if (!item.contains("=")) {
                    throw new ExcelKitXmlAnalyzeException(messageTemplate
                            + "Converter Expression error, Reference:[\"1=男,2=女\" or \"男=1,女=2\"].");
                }
            }
        }
        return value;
    }
}
