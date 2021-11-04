package nl;

import org.docx4j.Docx4J;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * benodigde dependencies in pom:
 *
 * <dependency>
 *     <groupId>org.docx4j</groupId>
 *     <artifactId>docx4j</artifactId>
 *     <version>6.1.2</version>
 * </dependency>
 * <dependency>
 *     <groupId>javax.xml.bind</groupId>
 *     <artifactId>jaxb-api</artifactId>
 *     <version>2.3.1</version>
 * </dependency>
 * <dependency>
 *     <groupId>com.sun.xml.bind</groupId>
 *     <artifactId>jaxb-impl</artifactId>
 *     <version>2.3.4</version>
 * </dependency>
 */
public class Template {

    private WordprocessingMLPackage wordprocessingMLPackage;

    public Template(String templateName) throws Docx4JException {
        this.loadTemplate(templateName);
    }

    /**
     * voorbeeld word template:
     *
     * Beste ${voornaam} ${achternaam},
     *
     * Hier is uw brief.
     *
     * Doei.
     *
     * de computert.
     * ${datum}
     */
    public void fillTemplate(String voornaam, String achternaam, Date datum) throws Exception {
        VariablePrepare.prepare(wordprocessingMLPackage);

        Map<String, String> variables = new HashMap<>();
        variables.put("voornaam", voornaam);
        variables.put("achternaam", achternaam);
        variables.put("datum", new SimpleDateFormat("dd-MM-yyyy").format(datum));

        wordprocessingMLPackage.getMainDocumentPart().variableReplace(variables);
    }

    public void writeToFile(String targetFile) throws Docx4JException {
        File file = new File(targetFile);
        this.wordprocessingMLPackage.save(file);
    }

    private void loadTemplate(String templateName) throws Docx4JException {
        File file = new File(Thread.currentThread().getContextClassLoader().getResource(templateName).getPath());
        this.wordprocessingMLPackage = Docx4J.load(file);
    }

}
