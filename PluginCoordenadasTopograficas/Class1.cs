using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Interop.Common;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.PlottingServices;
using System.IO;

namespace PluginCoordenadasTopograficas
{
    public class Class1
    {
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
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Title = "Selecione o arquivo xlsx com as coordenadas topográficas";
            openFileDialog.DefaultExt = "xlsx";
            openFileDialog.Filter = "Arquivo xlsx (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
            else
            {
                return null;
            }
        }
    }
}
