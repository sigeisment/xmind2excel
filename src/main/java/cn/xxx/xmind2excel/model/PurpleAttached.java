package cn.xxx.xmind2excel.model;

import java.util.List;

@lombok.Data
public class PurpleAttached {
    private String id;
    private String title;
    private Long width;
    private Notes notes;
    private List<String> labels;
    private List<Marker> markers;
    private RootTopicChildren children;
}
