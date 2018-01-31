package com.douma;

import com.douma.entity.MingxiEntity;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class excel {

    public static void main(String[] args) throws IOException {
        File file = new File("D:\\workspace\\doumaexcel\\src\\main\\resources\\file\\回款明细.xlsx");
        //读取原始数据
        List<MingxiEntity> yuanshiData = getData(file, 1);
        //处理重复
        Map<String, List<MingxiEntity>> lastDataMap = handleData(yuanshiData);
        //向整理中填写数据
        file = new File("D:\\workspace\\doumaexcel\\src\\main\\resources\\file\\整理.xlsx");
        Map<String, List<MingxiEntity>> misData = writeData(file, lastDataMap, 1);
        //标注在整理中未找到的明细
        file = new File("D:\\workspace\\doumaexcel\\src\\main\\resources\\file\\回款明细.xlsx");
        markNotFound(file, misData);
        System.out.println("结束");
    }

    /**
     * 读取原始数据
     * @param file
     * @param ignoreRows
     * @return
     * @throws IOException
     */
    public static List<MingxiEntity> getData(File file, int ignoreRows) throws IOException {
        Workbook xssfWorkbook = null;
        BufferedInputStream is = new BufferedInputStream(new FileInputStream(file));
        if(file.toString().endsWith("xls")){
            xssfWorkbook = new HSSFWorkbook(is);
        } else if(file.toString().endsWith("xlsx")){
            xssfWorkbook = new XSSFWorkbook(is);
        }
        Sheet sheet = xssfWorkbook.getSheetAt(0);
        if (sheet == null) {
            return null;
        }
        List<MingxiEntity> mingxiEntityList = new LinkedList<MingxiEntity>();
        //循环行
        for(int i=ignoreRows; i<=sheet.getLastRowNum(); i++){
            Row row = sheet.getRow(i);
            if(row == null){
                continue;
            }
            MingxiEntity mingxiEntity = new MingxiEntity();
            mingxiEntity.setIndex(i);
            mingxiEntityList.add(mingxiEntity);
            //循环列
            for(int j=0; j<row.getLastCellNum(); j++){
                Cell cell = row.getCell(j);
                if(cell == null){
                    continue;
                }
                String value = getCellValue(cell);
                if(j == 1){
                    mingxiEntity.setRiqi(value);
                } else if(j == 6){
                    mingxiEntity.setBenji(value);
                } else if(j == 3){
                    mingxiEntity.setKehu(value);
                }
            }
        }
        return mingxiEntityList;
    }

    /**
     * 获取不重复的数据
     * @param oldList
     * @return
     */
    public static Map<String, List<MingxiEntity>> handleData(List<MingxiEntity> oldList){
        Map<String, List<MingxiEntity>> newMap = new HashMap<String, List<MingxiEntity>>();
        List<MingxiEntity> newList = new LinkedList<MingxiEntity>();
        String lastKehu = null;
        for (MingxiEntity mingxiEntity : oldList) {
            if(!mingxiEntity.getKehu().equals(lastKehu)){
                lastKehu = mingxiEntity.getKehu();
                MingxiEntity newMingxiEntity = new MingxiEntity();
                newMingxiEntity.setIndex(mingxiEntity.getIndex());
                newMingxiEntity.setOtherIndex(new ArrayList<Integer>());
                newMingxiEntity.setRiqi(mingxiEntity.getRiqi());
                newMingxiEntity.setBenji(mingxiEntity.getBenji());
                newMingxiEntity.setKehu(mingxiEntity.getKehu());
                newList.add(newMingxiEntity);
            } else {
                MingxiEntity newMingxiEntity = newList.get(newList.size()-1);
                List<Integer> otherIndex = newMingxiEntity.getOtherIndex();
                otherIndex.add(mingxiEntity.getIndex());
                String benjin = newMingxiEntity.getBenji();
                Long benjinLong = Long.valueOf(benjin);
                benjinLong = benjinLong + Long.valueOf(mingxiEntity.getBenji());
                newMingxiEntity.setBenji(String.valueOf(benjinLong));
            }
        }
        for (MingxiEntity mingxiEntity : newList) {
            if(newMap.get(mingxiEntity.getRiqi()) == null){
                List<MingxiEntity> mingxiEntityList = new LinkedList<MingxiEntity>();
                mingxiEntityList.add(mingxiEntity);
                newMap.put(mingxiEntity.getRiqi(),mingxiEntityList);
            } else {
                List<MingxiEntity> mingxiEntityList = newMap.get(mingxiEntity.getRiqi());
                mingxiEntityList.add(mingxiEntity);
                newMap.put(mingxiEntity.getRiqi(),mingxiEntityList);
            }
        }
        return newMap;
    }

    /**
     * 整理信息回填到整理表格中
     * @param file
     * @param lastMap
     * @param ignoreRows
     * @return
     * @throws IOException
     */
    public static Map<String, List<MingxiEntity>> writeData(File file, Map<String, List<MingxiEntity>> lastMap, int ignoreRows) throws IOException {
        Workbook workbook = null;
        BufferedInputStream is = new BufferedInputStream(new FileInputStream(file));
        if(file.toString().endsWith("xls")){
            workbook = new HSSFWorkbook(is);
        } else if(file.toString().endsWith("xlsx")){
            workbook = new XSSFWorkbook(is);
        }
        Sheet sheet = workbook.getSheetAt(0);
        if (sheet == null) {
            return null;
        }
        for(int i=ignoreRows; i<=sheet.getLastRowNum(); i++){
            Row row = sheet.getRow(i);
            if(row == null){
                continue;
            }
            MingxiEntity thisMingxiEntity = new MingxiEntity();
            for(int j=0; j<row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }
                String value = getCellValue(cell);
                if(j == 5){
                    thisMingxiEntity.setRiqi(value);
                } else if(j == 3){
                    thisMingxiEntity.setBenji(value);
                } else if(j == 0){
                    thisMingxiEntity.setKehu(value);
                }
            }
            List<MingxiEntity> mingxiEntityList = lastMap.get(thisMingxiEntity.getRiqi());
            if(mingxiEntityList == null){
                System.out.println("明细中不存在 " + thisMingxiEntity.getRiqi() + "|" + thisMingxiEntity.getBenji() + "|" + thisMingxiEntity.getKehu());
            } else {
                for (MingxiEntity mingxiEntity : mingxiEntityList) {
                    if(mingxiEntity.getRiqi().equals(thisMingxiEntity.getRiqi())){
                        Long valueLong1 = Long.valueOf(mingxiEntity.getBenji());
                        Long valueLong2 = Long.valueOf(thisMingxiEntity.getBenji());
                        if((valueLong1 - valueLong2)>-100L && (valueLong1 - valueLong2)<100L){
                            row.getCell(4).setCellValue(mingxiEntity.getBenji());
                            if(!mingxiEntity.getKehu().equals(thisMingxiEntity.getKehu())){
                                row.getCell(13).setCellValue(mingxiEntity.getKehu());
                            }
                            //命中的项目从list中删除
                            mingxiEntityList.remove(mingxiEntity);
                            break;
                        }
                    }
                }
            }
        }
        FileOutputStream fo = new FileOutputStream(file); // 输出到文件
        workbook.write(fo);
        return lastMap;
    }

    /**
     * 在明细中标记在整理中未找到的
     * @param file
     * @param misData
     * @throws IOException
     */
    public static void markNotFound(File file, Map<String, List<MingxiEntity>> misData) throws IOException {
        Workbook workbook = null;
        BufferedInputStream is = new BufferedInputStream(new FileInputStream(file));
        if (file.toString().endsWith("xls")) {
            workbook = new HSSFWorkbook(is);
        } else if (file.toString().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(is);
        }
        Sheet sheet = workbook.getSheetAt(0);
        if (sheet == null) {
            return;
        }
        for (String s : misData.keySet()) {
            for (MingxiEntity mingxiEntity : misData.get(s)) {
                sheet.getRow(mingxiEntity.getIndex()).getCell(9).setCellValue("没找到");
                List<Integer> otherIndex = mingxiEntity.getOtherIndex();
                if(otherIndex != null && otherIndex.size()>0){
                    for (Integer index : otherIndex) {
                        sheet.getRow(index).getCell(9).setCellValue("没找到");
                    }
                }
            }
        }
        FileOutputStream fo = new FileOutputStream(file); // 输出到文件
        workbook.write(fo);
    }

    /**
     * 获取cell的值
     * @param cell
     * @return
     */
    static public String getCellValue(Cell cell){
        String value = null;
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                //如果为时间格式的内容
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    //注：format格式 yyyy-MM-dd hh:mm:ss 中小时为12小时制，若要24小时制，则把小h变为H即可，yyyy-MM-dd HH:mm:ss
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
                    value=sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue())).toString();
                    break;
                } else {
                    value = new DecimalFormat("0").format(cell.getNumericCellValue());
                }
                break;
            case HSSFCell.CELL_TYPE_STRING: // 字符串
                value = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                value = cell.getBooleanCellValue() + "";
                break;
            case HSSFCell.CELL_TYPE_FORMULA: // 公式
                value = cell.getCellFormula() + "";
                break;
            case HSSFCell.CELL_TYPE_BLANK: // 空值
                value = "";
                break;
            case HSSFCell.CELL_TYPE_ERROR: // 故障
                value = "非法字符";
                break;
            default:
                value = "未知类型";
                break;
        }
        return value;
    }
}
