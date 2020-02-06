package excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;

/**
 * 简单的导出测试，不包括表头
 */
@Slf4j
public class ExcelWriteTest extends BaseTest {
    /**
     * 每行数据是List<String>无表头
     *
     * @throws IOException
     */
    @Test
    public void writeWithoutHead() throws IOException {
        Date start = new Date();
        try (OutputStream out = new FileOutputStream(getMethodName())) {
            ExcelWriterBuilder builder=new ExcelWriterBuilder();
            ExcelWriter writer = new ExcelWriter(out, ExcelTypeEnum.XLSX, false);
            Sheet sheet1 = new Sheet(1, 0);
            sheet1.setSheetName("sheet1");
            List<List<String>> data = new ArrayList<>();
            for (int i = 0; i < 100000; i++) {
                List<String> item = new ArrayList<>();
                item.add("item0" + i);
                item.add("item1" + i);
                item.add("item2" + i);
                data.add(item);
            }
            writer.write0(data, sheet1);
            writer.finish();
        }
        Date end = new Date();
        String result = "总计耗时:" + ((end.getTime() - start.getTime())) + "ms";
        System.out.printf(result);
        log.info(result);
    }

}
