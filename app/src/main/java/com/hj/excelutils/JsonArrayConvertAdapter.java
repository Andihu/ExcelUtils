package com.hj.excelutils;


import com.hj.excellibrary.service.IConvertParserAdapter;
import com.hj.excellibrary.service.ISheet;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class JsonArrayConvertAdapter implements IConvertParserAdapter {

    @Override
    public void convert(ISheet sheet, Object o) {
        JSONArray jsonArray = null;
        try {
            jsonArray = new JSONArray((String) o);
        } catch (JSONException e) {
            e.printStackTrace();
        }
        for (int i = 0; i < jsonArray.length(); i++) {
            JSONObject json = (JSONObject) jsonArray.opt(i);
            String name = (String) json.opt("name");
            Object value = json.opt("value");
            int index = (int) json.opt("index");
            sheet.onCreateCell(name, value, index);
        }
    }
}
