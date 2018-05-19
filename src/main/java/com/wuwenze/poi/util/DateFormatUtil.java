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

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.concurrent.ExecutionException;

public class DateFormatUtil {
    private DateFormatUtil() {
    }

    public final static SimpleDateFormat ENGLISH_LOCAL_DF = new SimpleDateFormat("EEE MMM dd HH:mm:ss z yyyy", Locale.ENGLISH);
    private final static LoadingCache<String, SimpleDateFormat> mDateFormatLoadingCache =
            CacheBuilder.newBuilder()
                    .maximumSize(5)
                    .build(new CacheLoader<String, SimpleDateFormat>() {

                        @Override
                        public SimpleDateFormat load(String pattern) {
                            return new SimpleDateFormat(pattern);
                        }
                    });

    public static Date parse(String pattern, Object value) throws Exception {
        String valueString = (String) value;
        return mDateFormatLoadingCache.get(pattern).parse(valueString);
    }

    public static String format(String pattern, Date value) {
        try {
            return mDateFormatLoadingCache.get(pattern).format(value);
        } catch (ExecutionException e) {
            e.printStackTrace();
        }
        return value.toString();
    }
}
