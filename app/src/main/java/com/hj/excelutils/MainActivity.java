package com.hj.excelutils;

import static android.os.Environment.DIRECTORY_DOCUMENTS;

import androidx.activity.result.contract.ActivityResultContracts;
import androidx.annotation.NonNull;
import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.ContextCompat;

import android.Manifest;
import android.content.pm.PackageManager;
import android.content.res.AssetManager;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.widget.Toast;

import com.hj.excellibrary.Excel;
import com.hj.excellibrary.service.IParseListener;
import com.hj.excellibrary.service.IWriteListener;

import org.json.JSONArray;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.security.Permission;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    private static final String TAG = "MainActivity";

    @RequiresApi(api = Build.VERSION_CODES.R)
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED &&
                ContextCompat.checkSelfPermission(this, Manifest.permission.READ_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED) {
            test();
        } else {
            requestPermissions(new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE, Manifest.permission.READ_EXTERNAL_STORAGE,Manifest.permission.MANAGE_EXTERNAL_STORAGE}, 1000);
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (requestCode == 1000) {
            if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED && grantResults[1] == PackageManager.PERMISSION_GRANTED) {
                Log.d(TAG, "onRequestPermissionsResult: ");
                test();
            } else {
                Toast.makeText(this, "权限申请失败", Toast.LENGTH_SHORT).show();
            }
        }
    }

    private void test() {
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

                    File file = new File(getExternalCacheDir(), "temp_excel.xlsx");
                    Log.d(TAG, "onEndParse: +"+file.getPath());
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

