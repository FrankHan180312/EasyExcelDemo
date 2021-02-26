import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.CellExtra;
import model.Data;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class DataListener extends AnalysisEventListener<Data> {

    List<Data> list = new ArrayList<Data>();
    public void invoke(Data data, AnalysisContext analysisContext) {
        list.add(data);
        System.out.println(data);
    }

    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println(list);
    }
}
