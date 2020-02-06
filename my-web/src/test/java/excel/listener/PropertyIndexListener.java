package excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Slf4j
public class PropertyIndexListener<ExcelPropertyIndexModel> extends AnalysisEventListener<ExcelPropertyIndexModel> {
    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 100;

    List<ExcelPropertyIndexModel> list = new ArrayList<>();


    @Override
    public void onException(Exception e, AnalysisContext analysisContext) throws Exception {
        e.printStackTrace();
    }

    @Override
    public void invokeHead(Map map, AnalysisContext analysisContext) {
        log.info(map.toString());
    }

    @Override
    public void invoke(ExcelPropertyIndexModel o, AnalysisContext analysisContext) {
        log.info(o.toString());
        list.add(o);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (list.size() >= BATCH_COUNT) {
            saveData();
            // 存储完成清理 list
            list.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        saveData();
        log.info("所有数据解析完成！");
    }

    /**
     * 加上存储数据库
     */
    private void saveData() {
        log.info("{}条数据，开始存储数据库！", list.size());
        log.info("存储数据库成功！");
    }
}
