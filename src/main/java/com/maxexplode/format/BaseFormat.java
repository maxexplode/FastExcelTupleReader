package com.maxexplode.format;

import java.util.Collections;
import java.util.Set;

public abstract class BaseFormat {

    protected String currentFormat;

    public abstract String format(Integer formatId, String value);

    public abstract boolean supports(int formatId, String value);

    protected void setCurrentFormat(String currentFormat) {
        this.currentFormat = currentFormat;
    }

    public static BaseFormat any() {
        return anyFormat;
    }

    public Set<Integer> supportedFormats(){
        return Collections.emptySet();
    }

    private static final BaseFormat anyFormat = new BaseFormat() {
        @Override
        public String format(Integer formatId, String value) {
            return value;
        }

        @Override
        public boolean supports(int formatId, String value) {
            return false;
        }
    };
}