package br.com.feltex.avancado.modelo;

import lombok.Builder;
import lombok.Getter;

import java.util.List;

@Builder
@Getter
public class DadosArquivo {

    private String titulo;
    private String subTitulo;
    private String saudacao;
    private List<String> paragrafos;
    private String nomeArquivo;
    private String nomeImagemAvatar;
}
