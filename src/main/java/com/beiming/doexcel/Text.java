package com.beiming.doexcel;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.*;
import java.util.LinkedList;

public class Text {

        public static void isDirectory(File file, LinkedList addressList) {
            if(file.exists()){
                if (file.isFile()) {
//                    System.out.println("file is ==>>" + file.getAbsolutePath());
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

        public static int ceshi(int i){
            try {
                i = 10;
                System.out.print("try:" + i + "\n");
                return i;
            }catch (Exception e){
                i = 20; 
                System.out.println("catch:" + i + "\n");
            }finally {
                i = 30;
                System.out.println("finally" + i + "\n");
                return i;
            }
        }
        
        public static void main(String[] args) {
                /*File file = new File("D:/1月份(1).xls");
                try {
                    InputStream is = new FileInputStream(file.getAbsolutePath());
                    // jxl提供的Workbook类
                    Workbook wb = Workbook.getWorkbook(is);
                    Sheet sheet = wb.getSheet(0);
                    System.out.println(sheet.getCell(0, 550).getContents());
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (BiffException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                } */
                //System.out.println("main:" + ceshi(1) + "\n");
                System.out.println(0%250);
        }
    }
