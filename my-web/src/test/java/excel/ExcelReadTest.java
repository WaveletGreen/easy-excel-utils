package excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import excel.listener.PropertyIndexListener;
import excel.model.ExcelPropertyIndexModel;
import lombok.extern.slf4j.Slf4j;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.InputStream;

/**
 * 读取文件操作
 */
@Slf4j
public class ExcelReadTest extends BaseTest {
    @Test
    public void read() throws Exception {
        //保证有下面的文件，没有的话先运行ExcelWriteMultiHead#writeWithMultiHead单元测试方法
        try (InputStream in = new FileInputStream("target/writeWithMultiHead.xls");) {
            ExcelReaderBuilder excelReader = EasyExcel.read(in, ExcelPropertyIndexModel.class, new PropertyIndexListener());
            excelReader.sheet().doRead();
        }
    }
}
