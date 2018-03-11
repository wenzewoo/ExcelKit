package com.wuwenze.poi.test.entity;

import com.wuwenze.poi.annotation.ExportConfig;

public class User {

    @ExportConfig("UID")
    private Integer uid;

    @ExportConfig("用户名")
    private String username;

    @ExportConfig(value = "密码", replace = "******")
    private String password;

    @ExportConfig(value = "性别", width = 50, convert = "s:1=男,2=女")
    private Integer sex;

    @ExportConfig(value = "年级", convert = "c:com.wuwenze.poi.test.convert.GradeIdConvert")
    private Integer gradeId;

    @ExportConfig(value = "下拉框", range="c:com.wuwenze.poi.test.convert.RangeConvert")
    private String gendex;


    public Integer getUid() {
        return uid;
    }

    public User setUid(Integer uid) {
        this.uid = uid;
        return this;
    }

    public String getUsername() {
        return username;
    }

    public User setUsername(String username) {
        this.username = username;
        return this;
    }

    public String getPassword() {
        return password;
    }

    public User setPassword(String password) {
        this.password = password;
        return this;
    }

    public Integer getSex() {
        return sex;
    }

    public User setSex(Integer sex) {
        this.sex = sex;
        return this;
    }

    public Integer getGradeId() {
        return gradeId;
    }

    public User setGradeId(Integer gradeId) {
        this.gradeId = gradeId;
        return this;
    }

    public String getGendex() {
        return gendex;
    }

    public User setGendex(String gendex) {
        this.gendex = gendex;
        return this;
    }

    @Override
    public String toString() {
        return "User{" +
                "uid=" + uid +
                ", username='" + username + '\'' +
                ", password='" + password + '\'' +
                ", sex=" + sex +
                ", gradeId=" + gradeId +
                ", gendex='" + gendex + '\'' +
                '}';
    }
}
