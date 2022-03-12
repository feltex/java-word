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
        XWPFParagraph paragrafoTitulo = documento.createParagraph();
        paragrafoTitulo.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun tituloRun = paragrafoTitulo.createRun();
        tituloRun.setText(titulo);
        tituloRun.setColor("6495ED");
        tituloRun.setBold(true);
        tituloRun.setFontFamily(FONT_COURIER);
        tituloRun.setFontSize(20);
    }

    private void adicionarSubTitulo(final XWPFDocument documento, final String subtitulo) {
        XWPFParagraph paragrafoSubTitulo = documento.createParagraph();
        paragrafoSubTitulo.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun subtituloRun = paragrafoSubTitulo.createRun();
        subtituloRun.setText(subtitulo);
        subtituloRun.setColor("00CC44"); // Verde
        subtituloRun.setFontFamily(FONT_COURIER);
        subtituloRun.setFontSize(16);
        subtituloRun.setTextPosition(20);
        subtituloRun.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
    }

    private void adicionarImagem(final XWPFDocument documento, final String imagem) throws IOException, URISyntaxException, InvalidFormatException {
        XWPFParagraph paragrafoComImagem = documento.createParagraph();
        paragrafoComImagem.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun imagemRun = paragrafoComImagem.createRun();
        imagemRun.setTextPosition(20);
        var imagemPath = Paths.get(ClassLoader.getSystemResource(imagem).toURI());
        imagemRun.addPicture(Files.newInputStream(imagemPath), Document.PICTURE_TYPE_PNG,
                imagemPath.getFileName().toString(), Units.toEMU(50), Units.toEMU(50));

    }

    private void adicionarParagrafo(final XWPFDocument documento, final String conteudo) {
        XWPFParagraph paragrafo = documento.createParagraph();
        paragrafo.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun para1Run = paragrafo.createRun();
        para1Run.setText(conteudo);
    }

    private void adicionarSaudacao(final XWPFDocument documento, final String saudacao) {

        XWPFParagraph paragrafoSecao = documento.createParagraph();
        XWPFRun secaoRun = paragrafoSecao.createRun();
        secaoRun.setText(saudacao);
        secaoRun.setColor("00CC44");
        secaoRun.setBold(true);
        secaoRun.setFontFamily(FONT_COURIER);
    }

    private void adicionarTabela(final XWPFDocument documento, final List<Produto> produtos) {

        var tabela = documento.createTable();
        XWPFTableRow linha = tabela.getRow(0);
        linha.getCell(0).setText("Código");
        linha.addNewTableCell().setText("Nome");
        linha.addNewTableCell().setText("Preço");
        linha.addNewTableCell().setText("Quantidade");
        linha.addNewTableCell().setText("Total");

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
