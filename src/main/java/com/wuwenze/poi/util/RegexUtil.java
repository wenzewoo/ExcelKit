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

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;
import java.util.concurrent.ExecutionException;
import java.util.regex.Pattern;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;

/**
 * @author wuwenze
 * @date 2018/5/1
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
public class RegexUtil {

  private final static LoadingCache<String, Pattern> mRegexPatternLoadingCache =
      CacheBuilder.newBuilder()
          .maximumSize(5)
          .build(new CacheLoader<String, Pattern>() {

            @Override
            public Pattern load(String pattern) {
              return Pattern.compile(pattern);
            }
          });

  public static Boolean isMatches(String pattern, Object value) {
    try {
      String valueString = (String) value;
      return RegexUtil.mRegexPatternLoadingCache.get(pattern).matcher(valueString).matches();
    } catch (ExecutionException e) {
      e.printStackTrace();
    }
    return false;
  }
}
