package com.maxexplode;

import lombok.Builder;
import lombok.Getter;

@Getter
@Builder
public class ReadOptions {
    /**
     * Below indexes should be physical row indexes
     * Only set below if it is different from default values
     */
    @Builder.Default
    private String headerRowIdx = "1";
    @Builder.Default
    private String dataRowIdx = "2";
    private int sheetIdx;
    @Builder.Default
    private boolean skipNull = true;
}