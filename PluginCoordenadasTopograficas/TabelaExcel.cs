using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PluginCoordenadasTopograficas
{
    public class TabelaExcel
    {
        private readonly ExcelWorksheet planilhaDados;
        private readonly ExcelWorksheet planilhaConfiguracoes;

        public TabelaExcel(string caminhoArquivo)
        {
            ExcelPackage excelPackage = new ExcelPackage(new FileInfo(caminhoArquivo));
            this.planilhaDados = excelPackage.Workbook.Worksheets["Dados"];
            this.planilhaConfiguracoes = excelPackage.Workbook.Worksheets["Configurações"];
        }

        public string getConfiguracaoString(int linha, int coluna, string valorPadrao = null)
        {
            object configuracao = getConfiguracao(linha, coluna);
            if (configuracao == null) return valorPadrao;
            return configuracao.ToString();
        }

        public int getConfiguracaoInt(int linha, int coluna) => int.Parse(getConfiguracao(linha, coluna).ToString());
        public double getConfiguracaoDouble(int linha, int coluna) => double.Parse(getConfiguracao(linha, coluna).ToString());

        public string getDadoString(int linha, int coluna, string valorPadrao = null)
        {
            object dado = getDado(linha, coluna);
            if (dado == null) return valorPadrao;
            return dado.ToString();
        }

        public double getDadoDouble(int linha, int coluna, double valorPadrao)
        {
            object dado = getDado(linha, coluna);
            if (dado == null) return valorPadrao;
            return getDadoDouble(linha, coluna);
        }

        public int getDadoInt(int linha, int coluna) => int.Parse(getDado(linha, coluna).ToString());
        public double getDadoDouble(int linha, int coluna) => int.Parse(getDado(linha, coluna).ToString());

        private object getConfiguracao(int linha, int coluna) => planilhaConfiguracoes.GetValue(linha, coluna);
        private object getDado(int linha, int coluna) => planilhaDados.GetValue(linha, coluna);
    }
}
