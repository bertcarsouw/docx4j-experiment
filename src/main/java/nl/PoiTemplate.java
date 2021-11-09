package nl;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class PoiTemplate {

    private final XWPFDocument doc;
    private final String startDelimiter;
    private final String endDelimiter;

    public PoiTemplate(String startDelimiter, String endDelimiter) throws Exception {
        this.startDelimiter = startDelimiter;
        this.endDelimiter = endDelimiter;
        this.doc = new XWPFDocument(OPCPackage.open(Paths.get("").toAbsolutePath() + "/src/main/resources/testbrief-poi.docx"));
    }

    public void fillTemplate(String voornaam, String achternaam, Date datum) throws Exception {
        Map<String, String> variables = new HashMap<>();
        variables.put("voornaam", voornaam);
        variables.put("achternaam", achternaam);
        variables.put("datum", new SimpleDateFormat("dd-MM-yyyy").format(datum));

        doc.getParagraphs().forEach(paragraph -> {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs == null) {
                return;
            }

            boolean replaceCanidate = false;
            for (XWPFRun run : runs) {
                String text = run.getText(0);

                if (text != null) {
                    if (replaceCanidate) {
                        run.setText(variables.get(run.getText(0)), 0);
                        replaceCanidate = false;
                    } else if (text.equals(startDelimiter)) {
                        replaceCanidate = true;
                        run.setText("", 0); // remove start delimiter
                    } else if (text.equals(endDelimiter)) {
                        replaceCanidate = false;
                        run.setText("", 0); // remove end delimiter
                    }
                }
            }
        });

    }

    public void writeToFile(String path) throws Exception {
        this.doc.write(new FileOutputStream(path));
    }

}
