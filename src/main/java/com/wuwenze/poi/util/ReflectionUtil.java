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

package com.wuwenze.poi.util;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.beanutils.PropertyUtilsBean;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

public class ReflectionUtil {
    private ReflectionUtil() {}

    public static void setProperty(final Object bean, final String name, final Object value)
            throws IllegalAccessException, InvocationTargetException, NoSuchMethodException, InstantiationException {
        if (!name.contains(".")) {
            BeanUtils.setProperty(bean, name, value);
            return;
        }

        PropertyUtilsBean beanUtil = new PropertyUtilsBean();
        String[] propertyLevels = name.split("\\.");
        String propertyNameWithParent = "";
        for (int i = 0; i < propertyLevels.length; i++) {
            String p = propertyLevels[i];
            propertyNameWithParent = (propertyNameWithParent.equals("") ? p
                    : propertyNameWithParent + "." + p);

            if (i < (propertyLevels.length - 1) &&
                    beanUtil.getProperty(bean, propertyNameWithParent) != null) {
                continue;
            }
            Class pType = beanUtil.getPropertyType(bean,
                    propertyNameWithParent);
            if (i < (propertyLevels.length - 1)) {
                BeanUtils.setProperty(bean, propertyNameWithParent, pType
                        .getConstructor().newInstance());
            } else {
                Constructor<String> constructor = pType
                        .getConstructor(String.class);
                BeanUtils.setProperty(bean, propertyNameWithParent,
                        constructor.newInstance(value));
            }
        }
    }
}
