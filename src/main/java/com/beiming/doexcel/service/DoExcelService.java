package com.beiming.doexcel.service;

import com.beiming.doexcel.dao.DoExcelDao;
import com.beiming.doexcel.datasource.DatasourceCon;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

@Service
public class DoExcelService {

    @Autowired
    DoExcelDao excelDao;

    @Autowired
    DatasourceCon datasourceCon;

    /*public static void main(String[] args) {
            doSave("D:/ceshi/","newOne", 4, 5);
    }*/

    public StringBuffer doSave(String fileAddress, String tableName,
                        int startRow, int endRow) throws ClassNotFoundException, SQLException {
        Class.forName("com.mysql.jdbc.Driver");

        //创建数据库连接
        String url = datasourceCon.getUrl();
        Connection conn = DriverManager.getConnection(url, datasourceCon.getUsername(), datasourceCon.getPassword());
        Statement stat = conn.createStatement();

        ArrayList addressList = new ArrayList();
        ArrayList titleList = new ArrayList();
        isDirectory(new File(fileAddress), addressList);
        StringBuffer message = new StringBuffer();
        int f = 0;
        //遍历所有文件
        A : for (int add = 0; add < addressList.size(); add++){
            message = message.append(addressList.get(add).toString() + ":");
            String address = addressList.get(add).toString();
            File file = new File(address.replaceAll("\\\\", "/"));
            try {
                InputStream is = new FileInputStream(file.getAbsolutePath());
                // jxl提供的Workbook类
                Workbook wb = Workbook.getWorkbook(is);
                Sheet sheet = wb.getSheet(0);
                int myEndRow = sheet.getRows() - endRow;
                ArrayList thingList = new ArrayList();
                StringBuffer sb = new StringBuffer();
                int col = sheet.getColumns();
                
                //获得标题的List
                if (f == 0) {
                    for (int j = 0; j < col; j++) {
                        if (startRow == 1) {
                            titleList.add(sheet.getCell(j, 0).getContents());
                            continue;
                        } else {
                            int z = startRow - 2;
                            for (; z >= 1; z--) {
                                if (sheet.getCell(j, z).getContents().equals(sheet.getCell(j, z - 1).getContents())) {
                                    continue;
                                } else {
                                    sb = sb.insert(0, sheet.getCell(j, z).getContents());
                                }
                            }
                            sb = sb.insert(0, sheet.getCell(j, 0).getContents());
                        }
                        titleList.add(sb.toString());
                        sb.delete(0, sb.length());
                    }
                }
                
                //拼接数据，生成最终的List
                for (int i = startRow; i <= myEndRow; i++) {
                    for (int j = 0; j < col; j++) {
                        thingList.add(sheet.getCell(j, (i-1)).getContents());
                    }
                }

                //执行sql语句
                String getMes = excelDao.saveExcel(tableName, thingList, col, stat, f, titleList);
                message.append(getMes + "<br>");
                if ("数据库插入数据成功".equals(getMes)){
                    f++;
                }else {
                    for (; add < addressList.size(); add++){
                        message.append("上数据插入失败 停止执行。" + "<br>");
                    }
                    break A;
                }
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (BiffException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (SQLException e) {
                e.printStackTrace();
            } catch (ClassNotFoundException e) {
                e.printStackTrace();
            }
        }
        return message;

    }

    public void isDirectory(File file, ArrayList addressList) {
        if(file.exists()){
            if (file.isFile()) {
                addressList.add(file.getAbsolutePath());
            } else {

                File[] list = file.listFiles();
                if (list.length == 0) {
                    System.out.println(file.getAbsolutePath() + " is null");
                } else {
                    for (int i = 0; i < list.length; i++) {
                        isDirectory(list[i], addressList);
                    }
                }
            }
        }else{
            System.out.println("The directory is not exist!");
        }
    }
}