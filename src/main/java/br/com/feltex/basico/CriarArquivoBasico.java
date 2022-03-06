package br.com.feltex.basico;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class CriarArquivoBasico {

    private static String logo = "avatar.png";
    private static String paragraph1 = "arquivo1.txt";
    private static String paragraph2 = "arquivo2.txt";
    private static String paragraph3 = "arquivo3.txt";
    private static String output = "relatórioPedidos.docx";
    private static final String FONT_COURIER = "Courier";

    public static void main(String[] args) throws Exception {
        try (var document = new XWPFDocument()) {

            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Confirmacão de compras");
            titleRun.setColor("6495ED");
            titleRun.setBold(true);
            titleRun.setFontFamily(FONT_COURIER);
            titleRun.setFontSize(20);

            XWPFParagraph subTitle = document.createParagraph();
            subTitle.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun subTitleRun = subTitle.createRun();
            subTitleRun.setText("Este email é confidencial e direcionado ao cliente.");
            subTitleRun.setColor("00CC44");
            subTitleRun.setFontFamily(FONT_COURIER);
            subTitleRun.setFontSize(16);
            subTitleRun.setTextPosition(20);
            subTitleRun.setUnderline(UnderlinePatterns.DOT_DOT_DASH);

            XWPFParagraph image = document.createParagraph();
            image.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun imageRun = image.createRun();
            imageRun.setTextPosition(20);
            var imagePath = Paths.get(ClassLoader.getSystemResource(logo).toURI());
            imageRun.addPicture(Files.newInputStream(imagePath), Document.PICTURE_TYPE_PNG, imagePath.getFileName().toString(), Units.toEMU(50), Units.toEMU(50));

            XWPFParagraph sectionTitle = document.createParagraph();
            XWPFRun sectionTRun = sectionTitle.createRun();
            sectionTRun.setText("Detalhes do Pedido.");
            sectionTRun.setColor("00CC44");
            sectionTRun.setBold(true);
            sectionTRun.setFontFamily(FONT_COURIER);

            XWPFParagraph para1 = document.createParagraph();
            para1.setAlignment(ParagraphAlignment.BOTH);
            var string1 = convertTextFileToString(paragraph1);
            XWPFRun para1Run = para1.createRun();
            para1Run.setText(string1);

            XWPFParagraph para2 = document.createParagraph();
            para2.setAlignment(ParagraphAlignment.RIGHT);
            var string2 = convertTextFileToString(paragraph2);
            XWPFRun para2Run = para2.createRun();
            para2Run.setText(string2);
            para2Run.setItalic(true);

            XWPFParagraph para3 = document.createParagraph();
            para3.setAlignment(ParagraphAlignment.LEFT);
            var string3 = convertTextFileToString(paragraph3);
            XWPFRun para3Run = para3.createRun();
            para3Run.setText(string3);

            try (var out = new FileOutputStream(output)) {
                document.write(out);
            }
        }

    }

    public static String convertTextFileToString(String fileName) throws URISyntaxException, IOException {
        try (Stream<String> stream = Files.lines(Paths.get(ClassLoader.getSystemResource(fileName).toURI()))) {
            return stream.collect(Collectors.joining(" "));
        }
    }

}
