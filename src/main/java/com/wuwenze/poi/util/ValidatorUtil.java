/*
 * Copyright (c) 2018, 吴汶泽 (wenzewoo@gmail.com).
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

package com.wuwenze.poi.util;

import java.util.regex.Pattern;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;

/**
 * @author wuwenze
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class ValidatorUtil {

  private static final Pattern PATTERN_NUMERIC = Pattern.compile("^(-?\\d+)(\\.\\d+)?$");
  private static final Pattern PATTERN_EMAIL = Pattern
      .compile("^[\\w-]+(\\.[\\w-]+)*@[\\w-]+(\\.[\\w-]+)+$");
  private static final Pattern PATTERN_MOBILE_PHONE = Pattern.compile("^(1)\\d{10}$");

  public static boolean match(String value, String regularExp) {
    return ValidatorUtil.isEmpty(value) ? false
        : Pattern.compile(regularExp).matcher(value).matches();
  }

  /**
   * 判断是否为浮点数或者整数
   *
   * @param value 字符串
   * @return true Or false
   */
  public static boolean isNumeric(String value) {
    return ValidatorUtil.isEmpty(value) ? false
        : ValidatorUtil.PATTERN_NUMERIC.matcher(value).matches();
  }

  /**
   * 判断是否为正确的邮件格式
   *
   * @param value 字符串
   * @return true Or false
   */
  public static boolean isEmail(String value) {
    return ValidatorUtil.isEmpty(value) ? false
        : ValidatorUtil.PATTERN_EMAIL.matcher(value).matches();
  }

  /**
   * 判断字符串是否为合法手机号
   *
   * @param value 字符串
   * @return true Or false
   */
  public static boolean isMobile(String value) {
    return ValidatorUtil.isEmpty(value) ? false
        : ValidatorUtil.PATTERN_MOBILE_PHONE.matcher(value).matches();
  }

  /**
   * 判断是否为数字
   *
   * @param value 字符串
   * @return true Or false
   */
  public static boolean isNumber(String value) {
    try {
      Integer.parseInt(value);
    } catch (Exception ex) {
      return false;
    }
    return true;
  }

  /**
   * 判断字符串是否为空
   *
   * @param value 字符串
   * @return true Or false
   */
  public static boolean isEmpty(String value) {
    return null == value || "".equals(value.trim()) || "null".equalsIgnoreCase(value.trim());
  }
}
