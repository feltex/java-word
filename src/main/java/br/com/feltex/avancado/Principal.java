package br.com.feltex.avancado;

import br.com.feltex.avancado.modelo.DadosArquivo;
import br.com.feltex.avancado.modelo.Produto;

import java.math.BigDecimal;
import java.util.List;

public class Principal {

    public static void main(String[] args) throws Exception {

        var listaProdutos = gerarLista();
        var totalDoPedido = listaProdutos.stream()
                .map(p -> p.getPreco().multiply(BigDecimal.valueOf(p.getQuantidade())))
                .reduce(BigDecimal.ZERO, BigDecimal::add);
        var dados = DadosArquivo.builder()
                .titulo("Pedido de Compra realizado com sucesso.")
                .subTitulo("Obrigado por comprar no site feltex.com.br.")
                .saudacao("Olá Cliente, veja os detalhes dos itens comprados.")
                .nomeImagemAvatar("avatar.png")
                .paragrafos(
                        List.of("Veja abaixo toda a lista dos produtos. Valor total R$ " + totalDoPedido,
                                "Caso precise fazer uma troca ou avaliar os produtos comprados acesse compras.feltex.com.br",
                                "Obrigado!",
                                ""
                        ))
                .nomeArquivo("relatorioCompra.docx")
                .build();


        var criaArquivo = new CriaArquivo();
        criaArquivo.gerarArquivo(dados, listaProdutos);
    }

    private static List<Produto> gerarLista() {
        var produto1 = Produto.builder()
                .codigo(1)
                .nome("Mouse")
                .preco(BigDecimal.valueOf(10))
                .quantidade(1)
                .build();

        var produto2 = Produto.builder()
                .codigo(2)
                .nome("Teclado")
                .preco(BigDecimal.valueOf(22.51))
                .quantidade(3)
                .build();

        var produto3 = Produto.builder()
                .codigo(3)
                .nome("Monitor")
                .preco(BigDecimal.valueOf(395.51))
                .quantidade(2)
                .build();
        return List.of(produto1, produto2, produto3);
    }
}
