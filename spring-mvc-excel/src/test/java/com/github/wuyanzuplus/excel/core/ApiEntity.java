package com.github.wuyanzuplus.excel.core;


import lombok.Getter;
import lombok.Setter;

import javax.persistence.Entity;
import javax.persistence.Table;

/**
 * @author daniel.hu
 */
@Getter
@Setter
@Entity
@Table(name = ApiEntity.TABLE)
public class ApiEntity {

    public static final String TABLE = "t_sys_api";

    private Long id;

    private String project;

    private String apiCode;

    private String apiName;

    private Platform apiPlatform;

    private String apiUrl;

    private String memo;

    private Boolean valid;
}