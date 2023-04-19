package cn.xxx.xmind2excel.model;


@lombok.Data
public class XmindRoot {
    private String id;
    private String title;
    private Theme theme;
    private RootTopic rootTopic;
    private String topicPositioning;
    private String coreVersion;
}
