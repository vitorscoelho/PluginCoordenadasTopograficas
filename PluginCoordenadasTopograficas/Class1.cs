using Autodesk.AutoCAD.Runtime;

namespace PluginCoordenadasTopograficas
{
    public class Class1
    {
        [CommandMethod("DPT")]
        [CommandMethod("DesenharPontosTopograficos")]
        public void DesenharPontosTopograficos()
        {
            string caminhoArquivo = abrirJanelaSelecaoArquivo();
            if (caminhoArquivo == null) return;
            CriadorDesenho criadorDesenho = new CriadorDesenho();
            criadorDesenho.desenhar(caminhoArquivo);
        }

        /// <summary>
        /// Abre uma janela do tipo OpenFileDialog para que o usuário selecione o arquivo excel
        /// </summary>
        /// <returns> o caminho do arquivo selecionado. Retorna nulo se nenhum arquivo for selecionado
        public static string abrirJanelaSelecaoArquivo()
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog
            {
                Title = "Selecione o arquivo xlsx com as coordenadas topográficas",
                DefaultExt = "xlsx",
                Filter = "Arquivo xlsx (*.xlsx)|*.xlsx"
            };
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) return openFileDialog.FileName;
            return null;
        }
    }
}
