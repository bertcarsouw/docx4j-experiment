package nl;

import java.util.Date;


public class App {

    public static void main(String[] args) throws Exception {
        Template template = new Template("testbrief.docx");
        template.fillTemplate("John", "Doe", new Date());
        template.writeToFile("testbrief-2.docx");
    }

}
