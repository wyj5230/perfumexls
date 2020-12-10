import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.*;

public class Main {

    public static void main(String[] args) throws Exception {
        Map<String, String> sampleMap = ExcelService.getSampleXls("./key.xls");
        List<Cloth> cloths = new ArrayList<>();
        Title title = ExcelService.getTitle("./1.xls");
        ExcelService.getActualXls("./1.xls", 2, sampleMap, cloths);
        ExcelService.getActualXls("./2.xls", 2, sampleMap, cloths);
        Instant now = Instant.now();
        Instant yesterday = now.minus(1, ChronoUnit.DAYS);
        DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
        String yest = dateFormat.format(Date.from(yesterday));
        ExcelService.printMergedXls("TMALL_SALES_"+ yest, title, cloths);
        Scanner input = new Scanner(System.in);
        input.nextLine();
    }
}
