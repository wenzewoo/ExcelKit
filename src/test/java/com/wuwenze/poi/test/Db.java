package com.wuwenze.poi.test;

import com.wuwenze.poi.test.entity.User;

import java.util.*;

/**
 * 模拟数据库
 */
public class Db {
    private final static Map<Integer, String> grades = new HashMap<Integer, String>();
    private final static List<User> users = new ArrayList<User>();
    static {
        // 默认数据库字典查询 select * from tb_grades
        grades.put(1, "一年级学生");
        grades.put(2, "二年级学生");

        // 用户信息
        for (int i = 0; i < 10000; i++) {
            User user = new User()
                    .setUid(i)
                    .setUsername("Username:" + i)
                    .setPassword("123123123")
                    .setSex(i % 2 == 0 ? 1 : 2)
                    .setGradeId(i % 3 == 0 ? 1 : 2)
                    .setGendex("下拉框1");
            users.add(user);
        }

    }

    public static Map<Integer, String> getGrades() {
        return grades;
    }

    public static Integer getGradeIdByName(String name) {
        for (Integer key : grades.keySet()) {
            String val = grades.get(key);
            if(name.equals(val)) {
                return key;
            }
        }
        return -1;
    }

    public static List<User> getUsers() {
        return users;
    }

    public static void addUser(User user) {
        users.add(user);
    }


}
