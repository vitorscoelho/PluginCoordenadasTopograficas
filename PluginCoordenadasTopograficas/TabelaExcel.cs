using OfficeOpenXml;
using System.IO;

namespace PluginCoordenadasTopograficas
{
    class ConversaoDadoExcelException : System.Exception
    {
        public ConversaoDadoExcelException(string mensagem) : base(mensagem) { }
    }

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

        public string getDadoString(int linha, int coluna) => getString(linha, coluna, planilhaDados);
        public string getDadoString(int linha, int coluna, string valorPadrao) => getString(linha, coluna, valorPadrao, planilhaDados);
        public int getDadoInt(int linha, int coluna) => getInt(linha, coluna, planilhaDados);
        public int getDadoInt(int linha, int coluna, int valorPadrao) => getInt(linha, coluna, valorPadrao, planilhaDados);
        public double getDadoDouble(int linha, int coluna) => getDouble(linha, coluna, planilhaDados);
        public double getDadoDouble(int linha, int coluna, double valorPadrao) => getDouble(linha, coluna, valorPadrao, planilhaDados);

        public string getConfiguracaoString(int linha, int coluna) => getString(linha, coluna, planilhaConfiguracoes);
        public string getConfiguracaoString(int linha, int coluna, string valorPadrao) => getString(linha, coluna, valorPadrao, planilhaConfiguracoes);
        public int getConfiguracaoInt(int linha, int coluna) => getInt(linha, coluna, planilhaConfiguracoes);
        public int getConfiguracaoInt(int linha, int coluna, int valorPadrao) => getInt(linha, coluna, valorPadrao, planilhaConfiguracoes);
        public double getConfiguracaoDouble(int linha, int coluna) => getDouble(linha, coluna, planilhaConfiguracoes);
        public double getConfiguracaoDouble(int linha, int coluna, double valorPadrao) => getDouble(linha, coluna, valorPadrao, planilhaConfiguracoes);

        private string getString(int linha, int coluna, string valorPadrao, ExcelWorksheet worksheet) => getValor(linha, coluna, valorPadrao, worksheet).ToString();
        private string getString(int linha, int coluna, ExcelWorksheet worksheet) => getValor(linha, coluna, worksheet).ToString();
        private int getInt(int linha, int coluna, int valorPadrao, ExcelWorksheet worksheet) => parseInt(getString(linha, coluna, valorPadrao.ToString(), worksheet), linha, coluna, worksheet);
        private int getInt(int linha, int coluna, ExcelWorksheet worksheet) => parseInt(getString(linha, coluna, worksheet), linha, coluna, worksheet);
        private double getDouble(int linha, int coluna, double valorPadrao, ExcelWorksheet worksheet) => parseDouble(getString(linha, coluna, valorPadrao.ToString(), worksheet), linha, coluna, worksheet);
        private double getDouble(int linha, int coluna, ExcelWorksheet worksheet) => parseDouble(getString(linha, coluna, worksheet), linha, coluna, worksheet);

        private int parseInt(string valor, int linha, int coluna, ExcelWorksheet worksheet)
        {
            int valorConvertido;
            bool conversaoBemSucedida = int.TryParse(valor, out valorConvertido);
            if (conversaoBemSucedida) return valorConvertido;
            throw new ConversaoDadoExcelException($"O valor da célula L{linha}C{coluna}, na planilha '{worksheet.Name}', é igual a '{valor}', mas deveria ser um número inteiro.");
        }

        private double parseDouble(string valor, int linha, int coluna, ExcelWorksheet worksheet)
        {
            double valorConvertido;
            bool conversaoBemSucedida = double.TryParse(valor, out valorConvertido);
            if (conversaoBemSucedida) return valorConvertido;
            throw new ConversaoDadoExcelException($"O valor da célula L{linha}C{coluna}, na planilha '{worksheet.Name}', é igual a '{valor}', mas deveria ser um número real.");
        }

        /// <summary>
        /// Recupera o valor de uma determinada célula no arquivo xlsx.
        /// </summary>
        /// <param name="linha">índice da linha da célula</param>
        /// <param name="coluna">índice da coluna da célula</param>
        /// <param name="valorPadrao">valor que será o retorno do método caso o valor da célula seja nulo</param>
        /// <param name="worksheet">planilha da célula</param>
        /// <returns>o valor da célula, caso não seja nulo. Retorna 'valorPadrao', caso contrário</returns>
        private object getValor(int linha, int coluna, object valorPadrao, ExcelWorksheet worksheet)
        {
            object valor = worksheet.GetValue(linha, coluna);
            if (valor == null) return valorPadrao;
            return valor;
        }

        /// <summary>
        /// Recupera o valor de uma determinada célula no arquivo xlsx.
        /// Lança uma <see cref="ConversaoDadoExcelException"/> se o valor da célula é nulo.
        /// </summary>
        /// <param name="linha">índice da linha da célula</param>
        /// <param name="coluna">índice da coluna da célula</param>
        /// <param name="worksheet">planilha da célula</param>
        /// <returns>o valor da célula</returns>
        private object getValor(int linha, int coluna, ExcelWorksheet worksheet)
        {
            object valor = worksheet.GetValue(linha, coluna);
            if (valor == null) throw new ConversaoDadoExcelException($"O valor na célula L{linha}C{coluna}, na planilha '{worksheet.Name}', é nulo, mas não poderia ser.");
            return valor;
        }

        public bool dadoEhNull(int linha, int coluna) => (planilhaDados.GetValue(linha, coluna) == null);
    }
}
