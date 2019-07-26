package com.beiming.doexcel.dao;


import org.springframework.stereotype.Repository;

import java.sql.*;
import java.util.ArrayList;

@Repository
public class DoExcelDao {

    public String saveExcel(String tableName, ArrayList dataList, int column, Statement stat, int f, ArrayList notesList) throws SQLException, ClassNotFoundException{

        StringBuffer sb = new StringBuffer();
        //创建表
        if (f == 0) {
            sb = new StringBuffer("create table " + tableName + "(id int NOT NULL AUTO_INCREMENT,\n");
            for (int i = 0; i < column; i++) {
                sb.append("message" + String.valueOf(i) + " varchar(80) comment'"+ notesList.get(i) +"',\n");
            }
            sb.append("PRIMARY KEY (id)\n)");
            try{
                stat.executeUpdate(sb.toString());
            }catch (Exception e){
                return e.getMessage();
            }
        }

        //添加数据
        int flag = 0, plan = 0;
        sb.delete(0, sb.length());
        sb.append("insert into " + tableName + " (");
        for (int i = 0; i < column - 1; i++) {
            sb.append("message" + String.valueOf(i) + ", ");
        }
        sb.append("message" + String.valueOf(column - 1) + ")\nVALUES ");
        StringBuffer head = sb;
        int row = dataList.size()/column;
        for (int z= 0; z < row; z++) {
            sb.append("(");
            for (int i = 0; i < column - 1; i++) {
                sb.append(" '" + dataList.get(i+flag*column) + "',");
            }
            if (z == row - 1) {
                sb.append(" '" + dataList.get(column + flag*column - 1) + "') ");
                break;
            }
            sb.append(" '" + dataList.get(column + flag*column - 1) + "'), ");
            flag++;
        }
        plan = stat.executeUpdate(sb.toString());
        //判断是否插入成功
        if (plan == 0){
            return "数据库插入数据失败";
        }

        return "数据库插入数据成功";
    }
}
