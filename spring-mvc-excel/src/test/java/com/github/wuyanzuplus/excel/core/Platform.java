package com.github.wuyanzuplus.excel.core;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @author daniel.hu
 */
@Getter
@AllArgsConstructor
public enum Platform {
    系统("00"),
    运营("01"),
    租户("02"),
    UNKNOWN("99");

    private String value;

}
