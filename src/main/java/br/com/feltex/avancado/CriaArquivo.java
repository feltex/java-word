package br.com.feltex.avancado;

import br.com.feltex.avancado.modelo.DadosArquivo;
import br.com.feltex.avancado.modelo.Produto;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class CriaArquivo {

    private static final String FONT_COURIER = "Courier";

    public void gerarArquivo(final DadosArquivo dados, final List<Produto> produtos) throws IOException, URISyntaxException, InvalidFormatException {

        try (var documento = new XWPFDocument()) {
            adicionarTitulo(documento, dados.getTitulo());
            adicionarSubTitulo(documento, dados.getSubTitulo());
            adicionarImagem(documento, dados.getNomeImagemAvatar());
            adicionarSaudacao(documento, dados.getSaudacao());
            dados.getParagrafos().forEach(p -> adicionarParagrafo(documento, p));

            adicionarTabela(documento, produtos);

            try (var out = new FileOutputStream(dados.getNomeArquivo())) {
                documento.write(out);
            }
        }
    }

    private void adicionarTitulo(final XWPFDocument documento, final String titulo) {
        XWPFParagraph title = documento.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText(titulo);
        titleRun.setColor("6495ED");
        titleRun.setBold(true);
        titleRun.setFontFamily(FONT_COURIER);
        titleRun.setFontSize(20);
    }

    private void adicionarSubTitulo(final XWPFDocument documento, final String subtitulo) {
        XWPFParagraph subTitle = documento.createParagraph();
        subTitle.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun subTitleRun = subTitle.createRun();
        subTitleRun.setText(subtitulo);
        subTitleRun.setColor("00CC44"); // Verde
        subTitleRun.setFontFamily(FONT_COURIER);
        subTitleRun.setFontSize(16);
        subTitleRun.setTextPosition(20);
        subTitleRun.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
    }

    private void adicionarImagem(final XWPFDocument documento, final String imagem) throws IOException, URISyntaxException, InvalidFormatException {
        XWPFParagraph paragrafoComImagem = documento.createParagraph();
        paragrafoComImagem.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun imageRun = paragrafoComImagem.createRun();
        imageRun.setTextPosition(20);
        var imagemPath = Paths.get(ClassLoader.getSystemResource(imagem).toURI());
        imageRun.addPicture(Files.newInputStream(imagemPath), Document.PICTURE_TYPE_PNG,
                imagemPath.getFileName().toString(), Units.toEMU(50), Units.toEMU(50));

    }

    private void adicionarParagrafo(final XWPFDocument documento, final String conteudo) {
        XWPFParagraph para1 = documento.createParagraph();
        para1.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun para1Run = para1.createRun();
        para1Run.setText(conteudo);
    }

    private void adicionarSaudacao(final XWPFDocument documento, final String saudacao) {

        XWPFParagraph sectionTitle = documento.createParagraph();
        XWPFRun sectionTRun = sectionTitle.createRun();
        sectionTRun.setText(saudacao);
        sectionTRun.setColor("00CC44");
        sectionTRun.setBold(true);
        sectionTRun.setFontFamily(FONT_COURIER);
    }

    private void adicionarTabela(final XWPFDocument documento, final List<Produto> produtos) {

        var tabela = documento.createTable();
        XWPFTableRow tableRowOne = tabela.getRow(0);
        tableRowOne.getCell(0).setText("Codigo");
        tableRowOne.addNewTableCell().setText("Nome");
        tableRowOne.addNewTableCell().setText("Preco");
        tableRowOne.addNewTableCell().setText("Quantidade");
        tableRowOne.addNewTableCell().setText("Total");

        produtos.forEach(p -> criarLinha(tabela, p));

    }

    private void criarLinha(XWPFTable tabela, Produto p) {
        XWPFTableRow tableRowTwo = tabela.createRow();
        tableRowTwo.getCell(0).setText("" + p.getCodigo());
        tableRowTwo.getCell(1).setText(p.getNome());
        tableRowTwo.getCell(2).setText("" + p.getPreco());
        tableRowTwo.getCell(3).setText("" + p.getQuantidade());
        tableRowTwo.getCell(4).setText("" + p.getPreco().multiply(BigDecimal.valueOf(p.getQuantidade())));
    }
}
