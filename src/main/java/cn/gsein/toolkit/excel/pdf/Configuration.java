package cn.gsein.toolkit.excel.pdf;

import lombok.Builder;

@Builder
public class Configuration {

    @Builder.Default
    public PageSize pageSize = PageSize.A4;

    @Builder.Default
    public Mode mode = Mode.PORTRAIT;


}
