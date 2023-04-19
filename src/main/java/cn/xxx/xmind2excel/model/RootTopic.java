package cn.xxx.xmind2excel.model;

import java.util.List;

@lombok.Data
public class RootTopic {
    private String id;
    private String structureClass;
    private String title;
    private Long width;
    private List<Extension> extensions;
    private RootTopicChildren children;
}
