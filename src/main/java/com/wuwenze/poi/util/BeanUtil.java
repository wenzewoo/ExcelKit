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

package com.wuwenze.poi.util;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.beanutils.PropertyUtilsBean;
import org.apache.commons.beanutils.converters.DateConverter;

import java.lang.reflect.InvocationTargetException;

/**
 * @author wuwenze
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class BeanUtil extends org.apache.commons.beanutils.BeanUtils {

  public static void setComplexProperty(Object bean, String name, Object value)
      throws InvocationTargetException, IllegalAccessException, NoSuchMethodException, InstantiationException {
    //修复日期为空时bug
    ConvertUtils.register(new DateConverter(null), java.util.Date.class);
    if (!name.contains(".")) {
      BeanUtil.setProperty(bean, name, value);
      return;
    }
    PropertyUtilsBean propertyUtilsBean = new PropertyUtilsBean();
    String[] propertyLevels = name.split("\\.");
    String parentPropertyName = "";
    for (int i = 0; i < propertyLevels.length; i++) {
      String p = propertyLevels[i];
      parentPropertyName = (parentPropertyName.equals("") ? p : parentPropertyName + "." + p);
      if (i < (propertyLevels.length - 1) &&
          propertyUtilsBean.getProperty(bean, parentPropertyName) != null) {
        continue;
      }
      Class<?> parentClass = propertyUtilsBean.getPropertyType(bean, parentPropertyName);
      if (i < (propertyLevels.length - 1)) {
        BeanUtil.setProperty(bean,
            parentPropertyName, parentClass.getConstructor().newInstance());
      } else {
        BeanUtil.setProperty(bean, parentPropertyName,
            parentClass.getConstructor(String.class).newInstance(value));
      }
    }
  }
}
