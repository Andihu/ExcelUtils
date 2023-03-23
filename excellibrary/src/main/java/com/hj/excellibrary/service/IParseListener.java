package com.hj.excellibrary.service;

import org.json.JSONArray;

public interface IParseListener<T> {

    void onStartParse();

    /**
     * @param t 实体类
     * @param jsonObject 没有注解的字段
     */
    void onParse(T t, JSONArray jsonObject);

    /**
     * @param e 异常
     */
    void onParseError(Exception e);


    void onEndParse();

}