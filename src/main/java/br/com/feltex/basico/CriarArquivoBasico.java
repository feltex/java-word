package br.com.feltex.basico;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

public class CriarArquivoBasico {

    private static String output = "HistoriaDoDia.docx";
    private static final String FONT_COURIER = "Courier";

    public static void main(String[] args) throws Exception {
        try (var document = new XWPFDocument()) {

            XWPFParagraph titulo = document.createParagraph();
            titulo.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun tituloRun = titulo.createRun();
            tituloRun.setText("Batatinha Quando Nasce");
            tituloRun.setColor("6495ED");
            tituloRun.setBold(true);
            tituloRun.setFontFamily(FONT_COURIER);
            tituloRun.setFontSize(20);

            XWPFParagraph subTitulo = document.createParagraph();
            subTitulo.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun subTituloRun = subTitulo.createRun();
            subTituloRun.setText("Esta é uma história infantil");
            subTituloRun.setColor("00CC44");
            subTituloRun.setFontFamily(FONT_COURIER);
            subTituloRun.setFontSize(16);
            subTituloRun.setTextPosition(20);
            subTituloRun.setUnderline(UnderlinePatterns.DOT_DOT_DASH);

            XWPFParagraph imagem = document.createParagraph();
            imagem.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun imagemRun = imagem.createRun();
            imagemRun.setTextPosition(20);
            var imagePath = Paths.get(ClassLoader.getSystemResource("batatinha.jpg").toURI());
            imagemRun.addPicture(Files.newInputStream(imagePath), Document.PICTURE_TYPE_PNG, imagePath.getFileName().toString(), Units.toEMU(50), Units.toEMU(50));

            XWPFParagraph secao = document.createParagraph();
            XWPFRun secaoRun = secao.createRun();
            secaoRun.setText("Segue a história.");
            secaoRun.setColor("00CC44");
            secaoRun.setBold(true);
            secaoRun.setFontFamily(FONT_COURIER);

            XWPFParagraph paragrafo1 = document.createParagraph();
            paragrafo1.setAlignment(ParagraphAlignment.BOTH);
            XWPFRun para1Run = paragrafo1.createRun();
            para1Run.setText("Batatinha quando nasce \n Espalha a rama pelo chão");

            XWPFParagraph paragrafo2 = document.createParagraph();
            paragrafo2.setAlignment(ParagraphAlignment.RIGHT);
            XWPFRun para2Run = paragrafo2.createRun();
            para2Run.setText("O brotinho quando ama\n Poe a mão, no coração");
            para2Run.setItalic(true);

            XWPFParagraph paragrafo3 = document.createParagraph();
            paragrafo3.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun para3Run = paragrafo3.createRun();
            para3Run.setText("Menininha quando dorme\nPõe a mão no coração");

            try (var out = new FileOutputStream(output)) {
                document.write(out);
            }
        }
    }



}
