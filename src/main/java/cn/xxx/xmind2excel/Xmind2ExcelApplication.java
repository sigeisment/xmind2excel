package cn.xxx.xmind2excel;

import cn.xxx.xmind2excel.biz.XMind2Excel;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;
import java.util.stream.Collectors;

public class Xmind2ExcelApplication {
    private static final Logger logger = LoggerFactory.getLogger(Xmind2ExcelApplication.class);

    public static void main(String[] args) {
        String currentDir = System.getProperty("user.dir");
        String xMindDir = currentDir+"\\xmind";
        String xlsxDir = currentDir+"\\xlsx_file";
        File xMinFiledir = new File(xMindDir);
        File xlsxFiledir = new File(xlsxDir);
        File[] xlsxFiles = xlsxFiledir.listFiles((dir, name) -> name.endsWith(".xlsx"));

        Set<String> existsSet;
        if (xlsxFiles != null) {
            existsSet = Arrays.stream(xlsxFiles).map(item -> item.getName().replace(".xlsx", "")).collect(Collectors.toSet());
        } else {
            existsSet = new HashSet<>();
        }
        File[] xMindFiles = xMinFiledir.listFiles((dir, name) -> name.endsWith(".xmind") && !existsSet.contains(name.replace(".xmind","")));

        if (xMindFiles == null) {
            return;
        }
        for (File xMindFile : xMindFiles) {
            XMind2Excel.setXMindFile(xMindFile);
            XMind2Excel.setExcelFilePath(xlsxDir+"\\"+xMindFile.getName().replace(".xmind",".xlsx"));
            try {
                XMind2Excel.xMind2Excel();
            } catch (Exception e) {
                logger.error("xMind2Excel error",e);
            }
        }
//        String xMindFile = "D:\\tool\\xmind_to_xlsx\\xmind\\资金管控-230619.xmind";
//        XMind2Excel.setXMindFile(new File(xMindFile));
//        String excelFilePath = "D:\\tool\\xmind_to_xlsx\\xlsx_file\\资金管控-230619.xlsx";
//        XMind2Excel.setExcelFilePath(excelFilePath);
//        XMind2Excel.xMind2Excel();
    }

}
