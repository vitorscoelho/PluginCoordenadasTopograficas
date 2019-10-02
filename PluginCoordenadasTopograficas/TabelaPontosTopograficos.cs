using System;
using System.Collections.Generic;
using System.Globalization;

namespace PluginCoordenadasTopograficas
{
    public enum TipoRepresentacaoPonto
    {
        SemRepresentacao,
        Ponto,
        BlocoPadrao,
        Bloco
    }

    public class PontoTopografico
    {
        public readonly string nome;
        public readonly double norte;
        public readonly double leste;
        public readonly double altitude;

        public PontoTopografico(string nome, double norte, double leste, double altitude)
        {
            this.nome = nome;
            this.norte = norte;
            this.leste = leste;
            this.altitude = altitude;
        }
    }

    public class TabelaPontosTopograficos
    {
        private readonly TabelaExcel tabelaExcel;
        public readonly IEnumerable<PontoTopografico> pontosTopograficos;
        public readonly string separadorDecimal;
        public readonly string separadorMilhar;
        public readonly int casasDecimais;
        public readonly double origemNorte;
        public readonly double origemLeste;
        public readonly double multiplicadorDistancia;
        public readonly string leaderDimStyle;
        public readonly string leaderLayer;
        public readonly double leaderDeltaX;
        public readonly double leaderDeltaY;
        public readonly string mTextStyle;
        public readonly string mTextLayer;
        public readonly double mTextHeight;
        public readonly double mTextLineSpaceFactor;
        public readonly TipoRepresentacaoPonto representacaoPontoTipo;
        public readonly string representacaoPontoLayer;
        public readonly string representacaoPontoNomeBloco;
        public readonly double? representacaoPontoEscalaBloco;
        public readonly string padraoTextoDescritivo;
        private readonly CultureInfo cultureInfo;

        private static readonly int linhaInicialDadosPontosTopograficos = 2;

        public TabelaPontosTopograficos(string caminhoArquivo)
        {
            this.tabelaExcel = new TabelaExcel(caminhoArquivo);
            this.pontosTopograficos = criarListaPontos();
            this.separadorDecimal = tabelaExcel.getConfiguracaoString(2, 2);
            this.separadorMilhar = tabelaExcel.getConfiguracaoString(3, 2, valorPadrao: "");
            this.casasDecimais = tabelaExcel.getConfiguracaoInt(4, 2);
            this.origemNorte = tabelaExcel.getConfiguracaoDouble(5, 2);
            this.origemLeste = tabelaExcel.getConfiguracaoDouble(6, 2);
            this.multiplicadorDistancia = tabelaExcel.getConfiguracaoDouble(7, 2);
            this.leaderDimStyle = tabelaExcel.getConfiguracaoString(10, 2);
            this.leaderLayer = tabelaExcel.getConfiguracaoString(11, 2); ;
            this.leaderDeltaX = tabelaExcel.getConfiguracaoDouble(12, 2);
            this.leaderDeltaY = tabelaExcel.getConfiguracaoDouble(13, 2);
            this.mTextStyle = tabelaExcel.getConfiguracaoString(16, 2);
            this.mTextLayer = tabelaExcel.getConfiguracaoString(17, 2);
            this.mTextHeight = tabelaExcel.getConfiguracaoDouble(18, 2);
            this.mTextLineSpaceFactor = tabelaExcel.getConfiguracaoDouble(19, 2);
            this.representacaoPontoTipo = parseRepresentacaoPonto(tabelaExcel.getConfiguracaoString(22, 2));
            if (this.representacaoPontoTipo != TipoRepresentacaoPonto.SemRepresentacao)
            {
                this.representacaoPontoLayer = tabelaExcel.getConfiguracaoString(23, 2);
                if (this.representacaoPontoTipo != TipoRepresentacaoPonto.Ponto)
                {
                    this.representacaoPontoEscalaBloco = tabelaExcel.getConfiguracaoDouble(25, 2);
                    if (this.representacaoPontoTipo == TipoRepresentacaoPonto.Bloco)
                    {
                        this.representacaoPontoNomeBloco = tabelaExcel.getConfiguracaoString(24, 2);
                    }
                }
            }
            this.padraoTextoDescritivo = tabelaExcel.getConfiguracaoString(26, 2, valorPadrao: "");

            if (this.casasDecimais < 0) throw new ConversaoDadoExcelException($"O valor da célula L4C2, na planilha 'Configurações', deveria ser um número inteiro maior ou igual a zero.");
            if (this.multiplicadorDistancia <= 0) throw new ConversaoDadoExcelException($"O valor da célula L7C2, na planilha 'Configurações', deveria ser um número real maior que zero.");
            if (this.mTextHeight <= 0) throw new ConversaoDadoExcelException($"O valor da célula L18C2, na planilha 'Configurações', deveria ser um número real maior que zero.");
            if (this.mTextLineSpaceFactor < 0.25 || this.mTextLineSpaceFactor > 4.0) throw new ConversaoDadoExcelException($"O valor da célula L19C2, na planilha 'Configurações', deveria ser um número real maior ou igual a 0.25 e menor ou igual a 4.0.");

            if (this.representacaoPontoTipo != TipoRepresentacaoPonto.SemRepresentacao)
            {
                if (this.representacaoPontoLayer == null) throw new ConversaoDadoExcelException($"O valor na célula L23C2, na planilha 'Configurações', é nulo, mas não poderia ser.");
                if (this.representacaoPontoTipo != TipoRepresentacaoPonto.Ponto)
                {
                    if (this.representacaoPontoTipo == TipoRepresentacaoPonto.Bloco && this.representacaoPontoNomeBloco == null) throw new ConversaoDadoExcelException($"O valor na célula L24C2, na planilha 'Configurações', é nulo, mas não poderia ser.");
                    if (this.representacaoPontoEscalaBloco == null) throw new ConversaoDadoExcelException($"O valor na célula L25C2, na planilha 'Configurações', é nulo, mas não poderia ser.");
                }
            }

            this.cultureInfo = new CultureInfo("pt-BR");
            cultureInfo.NumberFormat.NumberDecimalDigits = this.casasDecimais;
            cultureInfo.NumberFormat.NumberDecimalSeparator = this.separadorDecimal;
            cultureInfo.NumberFormat.NumberGroupSeparator = this.separadorMilhar;
        }

        private List<PontoTopografico> criarListaPontos()
        {
            List<PontoTopografico> lista = new List<PontoTopografico>();
            int linha = linhaInicialDadosPontosTopograficos;
            while (true)
            {
                string colunaNaoDesenhar = tabelaExcel.getDadoString(linha, 5, valorPadrao: "");
                bool naoDesenhar = colunaNaoDesenhar.Equals("x", StringComparison.InvariantCultureIgnoreCase);
                if (naoDesenhar)
                {
                    linha++;
                    continue;
                }
                if (tabelaExcel.dadoEhNull(linha, 2) && tabelaExcel.dadoEhNull(linha, 3)) break;
                PontoTopografico ponto = new PontoTopografico(
                    nome: tabelaExcel.getDadoString(linha, 1, valorPadrao: ""),
                    norte: tabelaExcel.getDadoDouble(linha, 2),
                    leste: tabelaExcel.getDadoDouble(linha, 3),
                    altitude: tabelaExcel.getDadoDouble(linha, 4, valorPadrao: 0.0)
                );
                lista.Add(ponto);
                linha++;
            }
            return lista;
        }

        private string formatar(double valor) => string.Format(cultureInfo, "{0:N}", valor);

        public string textoDescritivo(PontoTopografico pontoTopografico)
        {
            string nome = pontoTopografico.nome;
            string norte = formatar(pontoTopografico.norte);
            string leste = formatar(pontoTopografico.leste);
            string altitude = formatar(pontoTopografico.altitude);
            return this.padraoTextoDescritivo.Replace("{nome}", nome).Replace("{norte}", norte).Replace("{leste}", leste).Replace("{altitude}", altitude);
        }

        private static TipoRepresentacaoPonto parseRepresentacaoPonto(string valor)
        {
            switch (valor)
            {
                case "Sem representação":
                    return TipoRepresentacaoPonto.SemRepresentacao;
                case "Ponto":
                    return TipoRepresentacaoPonto.Ponto;
                case "Bloco padrão":
                    return TipoRepresentacaoPonto.BlocoPadrao;
                case "Bloco":
                    return TipoRepresentacaoPonto.Bloco;
                default:
                    throw new ConversaoDadoExcelException($"O tipo de representação do ponto topográfico escolhido, '{valor}', na célula L22C2, é inválido.");
            }
        }
    }
}
