package com.github.wuyanzuplus.excel.core;

/**
 * @author daniel.hu
 */
public interface ValueImportResolver<T> {
    T importResolve(String val);
}
