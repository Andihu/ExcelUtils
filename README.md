# Excel 

| 注解                | 描述                                 |
| ------------------- | ------------------------------------ |
| @ExcelReadCell      | Name  标记表头名称                   |
| @ExcelTable         | 使用类上用来指定表名                 |
| @ExcelWriteAdapter  | 展开数据集合适配器                   |
| @ExcelWriteCell     | 输出文件编辑列名称，列号信息。       |
| @ExcelReadAggregate | 标记类成员变量用来保存没有标记的数据 |

### 读取excel文件：

数据源(表名称：测试表1)：

| 物品编码 | 物品名称        | 存放位置 | 备注                       | 日期     |
| -------- | --------------- | -------- | -------------------------- | -------- |
| TY122635 | 厨房-面团分割机 | 0        | SDS-30S                    | 2021.2.1 |
| TY122654 | 黑白激光打印机  | 0        | 兄弟 HL-5590DN             | 2021.2.2 |
| TY122652 | 黑白激光打印机  | 0        | 兄弟 HL-5590DN             | 2021.2.3 |
| TY122634 | 台式计算机      | 0        | 联想ThinkCentre M710t-D749 | 2021.2.4 |

创建实体对象：

```java
@ExcelTable(sheetName = "测试表1")
public class Table {

    @ExcelReadCell(name = "存放位置")
    public String storageLocation;

    @ExcelReadCell(name = "物品名称")
    public String name;

    @ExcelReadCell(name = "物品编码")
    public String code;

    //指定此变量保存其他数据，也可以不处理
    @ExcelReadAggregate
    public String extend;
}
```

这里只指定了三列数据，其他没有指定的数据列（备注、日期），将被聚合保存到被**@ExcelReadAggregate**标注extend变量中，当然如果不需要这些数据也可以不用声明变量使用**@ExcelReadAggregate**标注。

> 被@ExcelReadAggregate标注的对象接收的是一个JsonArray String 对象。

```
Table
{ 
storageLocation='0',
note='SDS-30S', 
name='厨房-面团分割机', 
code='TY2023122635', 
extend=
'[{"name":"日期","value":"2021.2.1","index":7},{"name":"备注","value":"SDS30S","index":8}]'
```

#### Use:

```java
Excel.get().readWith(is).doReadXLSX(new IParseListener<Table>() {
    @Override
    public void onStartParse() {
    }

    @Override
    public void onParse(Table test, JSONArray jsonArray) {
    }

    @Override
    public void onParseError(Exception e) {
    }

    @Override
    public void onEndParse() {
       
    }
}, Table.class);
```





## 输出excel文件：



```java
@ExcelTable(sheetName = "测试表1")
public class Table {

    @ExcelWriteCell(writeIndex = 2, writeName = "存放位置")
    public String storageLocation;

    @ExcelWriteCell(writeIndex = 1, writeName = "物品名称")
    public String name;

    @ExcelWriteCell(writeIndex = 0, writeName = "物品编码")
    public String code;
    
    //如果你将多个数据聚合在某一个变量中，可以通过实现IConvertParserAdapter接口来处理数据以便正确写入文件
    @ExcelWriteAdapter(adapter = JsonArrayConvertAdapter.class)
    public String extend;
}
```
##### @ExcelWriteCell

ExcelWriteCell注解有两个属性，writeIndex指定数据所属列，writeName指定列名称

##### @ExcelWriteAdapter

ExcelWriteAdapter用来辅助工具正确写入用户自定义的聚合数据。

这里extend 的数据如下：

```json
[
    {
        "name":"日期",
        "value":"2021.2.9",
        "index":3
    },
    {
        "name":"备注",
        "value":"1.0",
        "index":4
    }
]
```

Name 表示列名称，value表示值，index表示列号，这里的数据结构可以自行定义。

##### IConvertParserAdapter 接口

使用了聚合数据，就需要实现IConvertParserAdapter接口用来解析你的聚合数据并通过**ISheet**接口回调数据的列名称，值，列号等信息。

针对上面的聚合数据：

```java
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
```

##### @ExcelWriteAdapter使用方法：

```java
 @ExcelWriteAdapter(adapter = JsonArrayConvertAdapter.class)
 public String extend;
```

#### Use:

```java
 Excel.get().writeWith(file).doWrite(new IWriteListener() {
                        @Override
                        public void onStartWrite() {
                            Log.d(TAG, "onStartWrite: ");
                        }

                        @Override
                        public void onWriteError(Exception e) {
                            Log.d(TAG, "onWriteError: "+e);
                        }

                        @Override
                        public void onEndWrite() {
                            Log.d(TAG, "onEndWrite: ");
                        }
                    },data);
```

