package cn.xxx.xmind2excel.biz;


import cn.xxx.xmind2excel.model.*;
import cn.xxx.xmind2excel.util.ExcelUtil;
import cn.xxx.xmind2excel.util.FileExtension;
import cn.xxx.xmind2excel.util.FileUtil;
import com.google.gson.Gson;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

/**
 * @author xiongchenghui
 * @date 2020-08-12
 * &Desc XMind to Excel tool
 */
public class XMind2Excel {
    private static Logger logger = LoggerFactory.getLogger(XMind2Excel.class);

    /** XMind 原始文件 */
    static File xMindFile;

    /** 输出的测试用例Excel文件 */
    static String excelFilePath;

    /** 用例集合 */
    static List<TestCasePO> testCases = new ArrayList<>();

    /** 用例集合 */
    public static TestCaseInfo testCaseInfo = new TestCaseInfo();

    /** 导出Excel Style */
    static CellStyle cellStyle = null;

    public static void setXMindFile(File xMindFile) {
        XMind2Excel.xMindFile = xMindFile;
    }

    public static void setExcelFilePath(String excelFilePath) {
        XMind2Excel.excelFilePath = excelFilePath;
    }

    /***
     * &Desc: xMind 转换 Excel文件
     * @param
     * @return void
     */
    public static void xMind2Excel() {
        // 将xMind转换为zip文件
        File xMindZipFile = FileUtil.transferXMind2Zip(xMindFile);
        // 获取xMind的zip文件路径
        String zipPath = xMindZipFile.getAbsolutePath();
        // 获取ZIP文件解压目录descDir
        String descDir = zipPath.replaceAll("\\" + FileExtension.ZIP, "");

        // 解压ZIP文件 到descDir目录
        try {
            FileUtil.unZipFiles(zipPath, descDir);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error(e.getMessage());
        }

        try (FileReader fileReader = new FileReader(descDir + "/content.json")){
            // 获取json
            Gson gson = new Gson();
            // 获取xMind根节点
            XmindRoot[] xmindRoot = gson.fromJson(fileReader, XmindRoot[].class);
            // 从根节点开始遍历解析所有节点, 生成测试用例list
            testCases.clear();
            logger.info("******************解析用例********************");
            parseXMind(xmindRoot);
            logger.info("=====================>用例解析完毕");
        }catch (Exception de){
            de.printStackTrace();
            logger.error(de.getMessage());
        }

        //删除临时文件夹
        logger.info("******************清理临时文件********************");
        File temp = new File(xMindZipFile.getParent());
        FileUtil.deleteDir(temp);

        // list用例写入Excel 统计用例总数、步骤数、验证点数
        logger.info("******************用例写入Excel，统计用例基本信息********************");
        testCaseWrite2Excel();

        logger.info("******************用例转换完毕********************");

    }

    /***
     * &Desc: test case从list中写入Excel
     * @param
     * @return void
     */
    private static void testCaseWrite2Excel(){
        ExcelUtil.createExcel(excelFilePath);
        // 读取Excel文件;
        Workbook workbook = ExcelUtil.readExcel(excelFilePath);

        // 创建一个样式
        setCellStyle(workbook);
        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        // 获取解析用例的表格
        Sheet caseSheet = workbook.getSheetAt(0);
        // 创建表头
        setSheetColumnHeader(caseSheet, cellStyle);

        setCellStyle(workbook);
        // 逐条写入用例 并统计测试步骤验证点数量
        int steps = 0;
        int checkPointers = 0;
        int testCaseRowNum = 1;
        Row testcaseRow = caseSheet.createRow(testCaseRowNum);
        for (TestCasePO po: testCases) {
            steps += po.getActions().size();
            checkPointers += po.getResults().size();
            insertTestCase2Excel(testcaseRow, po);
            if(testCaseRowNum < testCases.size()){
                testCaseRowNum ++;
                testcaseRow = caseSheet.createRow(testCaseRowNum);
            }
        }
        // 统计测试用例、测试步骤、验证点数量
        testCaseInfo.setTestCaseNo(testCases.size());
        testCaseInfo.setTestCaseSteps(steps);
        testCaseInfo.setTestCaseCheckPointers(checkPointers);

        // 关闭文件流
        OutputStream stream = null;
        try {
            stream = new FileOutputStream(excelFilePath);
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
            logger.error(e.getMessage());
        } finally {
            try {
                stream.close();
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    /***
     * &Desc: 设置单元格格式
     * @param workbook 工作簿
     * @return void
     */
    private static void setCellStyle(Workbook workbook){
        /** 创建一个样式 */
        cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
    }

    /***
     * &Desc: 设置用例列头
     * @param sheet 工作sheet
     * @param cellStyle 单元格格式
     * @return void
     */
    private static void setSheetColumnHeader(Sheet sheet, CellStyle cellStyle){
        // 创建表头
        Row testcaseTitle = sheet.createRow(0);

        Cell testCatalogCell = testcaseTitle.createCell(TestCaseTemplate.TESTCASECATALOG);
        testCatalogCell.setCellValue("用例目录");
        testCatalogCell.setCellStyle(cellStyle);

        Cell testNameCell = testcaseTitle.createCell(TestCaseTemplate.TESTCASENAME);
        testNameCell.setCellValue("用例名称");
        testNameCell.setCellStyle(cellStyle);

        Cell demandId = testcaseTitle.createCell(TestCaseTemplate.DEMAND_ID);
        demandId.setCellValue("需求ID");
        demandId.setCellStyle(cellStyle);

        Cell predicationCell = testcaseTitle.createCell(TestCaseTemplate.PREDICATION);
        predicationCell.setCellValue("前置条件");
        predicationCell.setCellStyle(cellStyle);

        Cell stepCell = testcaseTitle.createCell(TestCaseTemplate.ACTIONS);
        stepCell.setCellValue("用例步骤");
        stepCell.setCellStyle(cellStyle);

        Cell expectCell = testcaseTitle.createCell(TestCaseTemplate.RESULTS);
        expectCell.setCellValue("预期结果");
        expectCell.setCellStyle(cellStyle);

        Cell testCaseTypeCell = testcaseTitle.createCell(TestCaseTemplate.TESTCASE_TYPE);
        testCaseTypeCell.setCellValue("用例类型");
        testCaseTypeCell.setCellStyle(cellStyle);

        Cell testcaseStatus = testcaseTitle.createCell(TestCaseTemplate.TESTCASE_STATUS);
        testcaseStatus.setCellValue("用例状态");
        testcaseStatus.setCellStyle(cellStyle);

        Cell importanceCell = testcaseTitle.createCell(TestCaseTemplate.PRIORITY);
        importanceCell.setCellValue("用例等级");
        importanceCell.setCellStyle(cellStyle);

        Cell createdBy = testcaseTitle.createCell(TestCaseTemplate.CREATED_BY);
        createdBy.setCellValue("创建人");
        createdBy.setCellStyle(cellStyle);

    }

    /***
     * &Desc: 测试用例对象 写入到Excel的指定行
     * @param testcaseRow 写入的Excel sheet 表行
     * @param testCasePO 测试用例对象
     * @return void
     */
    private static void insertTestCase2Excel(Row testcaseRow, TestCasePO testCasePO){
        Cell testCaseCatalogCell = testcaseRow.createCell(TestCaseTemplate.TESTCASECATALOG);
        testCaseCatalogCell.setCellStyle(cellStyle);
        testCaseCatalogCell.setCellValue(testCasePO.getTestCaseCatalog());

        Cell demandIdCell = testcaseRow.createCell(TestCaseTemplate.DEMAND_ID);
        demandIdCell.setCellStyle(cellStyle);
        demandIdCell.setCellValue(testCasePO.getDemandId());

        Cell testCaseNameCell = testcaseRow.createCell(TestCaseTemplate.TESTCASENAME);
        testCaseNameCell.setCellStyle(cellStyle);
        testCaseNameCell.setCellValue(testCasePO.getTestCaseName());

        Cell predicationCell = testcaseRow.createCell(TestCaseTemplate.PREDICATION);
        predicationCell.setCellStyle(cellStyle);
        predicationCell.setCellValue(testCasePO.getPredication());

        Cell actionsCell = testcaseRow.createCell(TestCaseTemplate.ACTIONS);
        actionsCell.setCellStyle(cellStyle);
        List<String> actions = testCasePO.getActions();
        StringBuilder sb = new StringBuilder();
        for(int item=0; item<actions.size(); item++){
            sb.append(actions.get(item));
            if(item < actions.size()-1){
                sb.append("\n");
            }
        }
        actionsCell.setCellValue(sb.toString());

        Cell resultsCell = testcaseRow.createCell(TestCaseTemplate.RESULTS);
        resultsCell.setCellStyle(cellStyle);
        List<String> results = testCasePO.getResults();
        sb.setLength(0);
        for(int item=0; item<results.size(); item++){
            sb.append(results.get(item));
            if(item < results.size()-1){
                sb.append("\n");
            }
        }
        resultsCell.setCellValue(sb.toString());

        Cell testCaseTypeCell = testcaseRow.createCell(TestCaseTemplate.TESTCASE_TYPE);
        testCaseTypeCell.setCellStyle(cellStyle);
        testCaseTypeCell.setCellValue(testCasePO.getTestCaseType());

        Cell testCaseStatusCell = testcaseRow.createCell(TestCaseTemplate.TESTCASE_STATUS);
        testCaseStatusCell.setCellStyle(cellStyle);
        testCaseStatusCell.setCellValue(testCasePO.getTestCaseStatus());

        Cell priorityCell = testcaseRow.createCell(TestCaseTemplate.PRIORITY);
        priorityCell.setCellStyle(cellStyle);
        priorityCell.setCellValue(testCasePO.getPriority());

        Cell createdByCell = testcaseRow.createCell(TestCaseTemplate.CREATED_BY);
        createdByCell.setCellStyle(cellStyle);
        createdByCell.setCellValue(testCasePO.getCreatedBy());
    }

    /***
     * &Desc: 解析XMind
     * @param nodes 指定节点开始解析
     * @return void
     */
    private static void parseXMind(XmindRoot... nodes) throws FileNotFoundException {
        if (nodes.length == 0) {
            return;
        }
        for (XmindRoot node : nodes) {
            RootTopic rootTopic = node.getRootTopic();
            if (Objects.isNull(rootTopic)) {
                continue;
            }
            RootTopicChildren children = rootTopic.getChildren();
            if (Objects.isNull(children)) {
                continue;
            }
            List<PurpleAttached> attached = children.getAttached();
            if (CollectionUtils.isEmpty(attached)) {
                continue;
            }
            for (PurpleAttached purpleAttached : attached) {
                // 测试用例名称
                String parentName = purpleAttached.getTitle();
                String demandId = purpleAttached.getNotes().getPlain().getContent();
                RootTopicChildren subTestCase = purpleAttached.getChildren();
                List<PurpleAttached> subAttachedList = subTestCase.getAttached();
                for (PurpleAttached subAttached : subAttachedList) {
                    TestCasePO testCasePO = new TestCasePO();
                    testCasePO.setTestCaseType("功能用例");
                    testCasePO.setTestCaseStatus("正常");
                    testCases.add(testCasePO);
                    // 需求id
                    testCasePO.setDemandId(demandId);
                    // 测试用例子名称
                    String subName = subAttached.getTitle();
                    testCasePO.setTestCaseName(parentName+"-"+subName);
                    // 前置条件
                    String note = subAttached.getNotes().getPlain().getContent();
                    testCasePO.setPredication(note);
                    // 优先级
                    String priority = subAttached.getMarkers().get(0).getMarkerId().replaceAll("priority-", "");
                    testCasePO.setPriority("P"+priority);
                    PurpleAttached operate = subAttached.getChildren().getAttached().get(0);
                    // 用例步骤
                    String operateTitle = operate.getTitle();
                    testCasePO.setActions(Collections.singletonList(operateTitle));
                    // 预期结果
                    String resultTitle = operate.getChildren().getAttached().get(0).getTitle();
                    testCasePO.setResults(Collections.singletonList(resultTitle));
                }
            }
        }
    }
}