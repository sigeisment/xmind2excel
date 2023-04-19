package cn.xxx.xmind2excel.model;

@lombok.Data
public class Theme {
    private SubTopic subTopic;
    private Summary summary;
    private Boundary boundary;
    private CalloutTopic calloutTopic;
    private Topic centralTopic;
    private Topic mainTopic;
    private SummaryTopic summaryTopic;
    private FloatingTopic floatingTopic;
    private Relationship relationship;
    private MapClass map;
}
