package com.hj.excelutils;

import static android.os.Environment.DIRECTORY_DOCUMENTS;

import androidx.appcompat.app.AppCompatActivity;

import android.content.res.AssetManager;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;

import com.hj.excellibrary.Excel;
import com.hj.excellibrary.service.IParseListener;
import com.hj.excellibrary.service.IWriteListener;

import org.json.JSONArray;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    private static final String TAG = "MainActivity";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        AssetManager am = getResources().getAssets();
        try {
            InputStream is = am.open("export_20230320_1600.xlsx");

            List<Table> data = new ArrayList<>();

            Excel.get().readWith(is).doReadXLSX(new IParseListener<Table>() {
                @Override
                public void onStartParse() {
                    Log.d(TAG, "onStartParse >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
                }

                @Override
                public void onParse(Table test, JSONArray jsonArray) {
                    Log.d(TAG, "onParse: " + test.toString() + "json:" + jsonArray.toString());
                    data.add(test);
                }

                @Override
                public void onParseError(Exception e) {
                    Log.d(TAG, "onParseError: " + e);
                }

                @Override
                public void onEndParse() {
                    Log.d(TAG, "onEndParse>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");


                    File externalStoragePublicDirectory = Environment.getExternalStoragePublicDirectory(DIRECTORY_DOCUMENTS);
                    File file = new File(externalStoragePublicDirectory, "temp_excel.xlsx");
                    if (!file.exists()) {
                        try {
                            file.createNewFile();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    Excel.get()
                            .writeWith(file)
                            .doWrite(
                                    new IWriteListener() {
                                        @Override
                                        public void onStartWrite() {
                                            Log.d(TAG, "onStartWrite: ");
                                        }

                                        @Override
                                        public void onWriteError(Exception e) {
                                            Log.d(TAG, "onWriteError: " + e);
                                        }

                                        @Override
                                        public void onEndWrite() {
                                            Log.d(TAG, "onEndWrite: ");
                                        }
                                    }, data);


                }
            }, Table.class);


        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}

