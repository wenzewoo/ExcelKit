/**

 * Copyright (c) 2017, 吴汶泽 (wuwz@live.com).

 *

 * Licensed under the Apache License, Version 2.0 (the "License");

 * you may not use this file except in compliance with the License.

 * You may obtain a copy of the License at

 *

 *      http://www.apache.org/licenses/LICENSE-2.0

 *

 * Unless required by applicable law or agreed to in writing, software

 * distributed under the License is distributed on an "AS IS" BASIS,

 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.

 * See the License for the specific language governing permissions and

 * limitations under the License.

 */
package com.wuwenze.poi.pojo;

public class ExportItem {

	private String field; // 属性名
	private String display; // 显示名
	private short width; // 宽度
	private String convert;
//	private short color;
	private String replace;
	private String  range;//数据有效性 下拉框

	public String getField() {
		return field;
	}

	public ExportItem setField(String field) {
		this.field = field;
		return this;
	}

	public String getDisplay() {
		return display;
	}

	public ExportItem setDisplay(String display) {
		this.display = display;
		return this;
	}

	public short getWidth() {
		return width;
	}

	public ExportItem setWidth(short width) {
		this.width = width;
		return this;
	}

	public String getConvert() {
		return convert;
	}

	public ExportItem setConvert(String convert) {
		this.convert = convert;
		return this;
	}

//	public short getColor() {
//		return color;
//	}
//
//	public ExportItem setColor(short color) {
//		this.color = color;
//		return this;
//	}

	public String getReplace() {
		return replace;
	}

	public ExportItem setReplace(String replace) {
		this.replace = replace;
		return this;
	}

	public String getRange() {
		return range;
	}

	public ExportItem setRange(String range) {
		this.range = range;
		return this;
	}
}
