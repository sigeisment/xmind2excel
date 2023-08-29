package cn.xxx.xmind2excel.biz;

import lombok.Data;

import java.util.List;

/**
 * @author xiongchenghui
 * @date 2020-09-14
 * &Desc 测试用例对象
 */
@Data
public class TestCasePO {
    /** 用例目录 */
    private String testCaseCatalog;
    private String demandId;
    private String createdBy;

    /** 用例名称 */
    private String testCaseName;

    /** 前置条件 */
    private String predication;

    /** 用例步骤 */
    private List<String> actions;

    /** 预期结果 */
    private List<String> results;

    /** 用例类型 */
    private String testCaseType;

    /**
     * 用例状态
     */
    private String testCaseStatus;

    /** 用例等级 */
    private String priority;

}