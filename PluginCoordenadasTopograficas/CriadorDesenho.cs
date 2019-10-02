using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace PluginCoordenadasTopograficas
{
    class CriadorDesenhoException : System.Exception
    {
        public CriadorDesenhoException(string mensagem) : base(mensagem) { }
    }

    class CriadorDesenho
    {
        private static readonly string NOME_BLOCO_PADRAO = "PONTO_TOPOGRAFICO_PADRAO";

        private TabelaPontosTopograficos tabela = null;

        private Document document = null;
        private Database database = null;
        private Transaction transaction = null;
        private BlockTable blockTable = null;
        private LayerTable layerTable = null;
        private DimStyleTable dimStyleTable = null;
        private TextStyleTable textStyleTable = null;
        private BlockTableRecord blockTableRecordModel = null;

        public void desenhar(string caminhoArquivo)
        {
            this.document = Application.DocumentManager.MdiActiveDocument;
            this.database = document.Database;
            this.transaction = database.TransactionManager.StartTransaction();
            this.blockTable = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
            this.layerTable = transaction.GetObject(database.LayerTableId, OpenMode.ForRead) as LayerTable;
            this.dimStyleTable = transaction.GetObject(database.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
            this.textStyleTable = transaction.GetObject(database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
            this.blockTableRecordModel = transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            try
            {
                this.tabela = new TabelaPontosTopograficos(caminhoArquivo);
                verificarValidadeDosDados(tabela);
                desenharPontosTopograficos(tabela);
            }
            catch (CriadorDesenhoException exception)
            {
                Application.ShowAlertDialog("Não foi possível desenhar os pontos. Motivo:\r\n" + exception.Message);
            }
            catch (ConversaoDadoExcelException exception)
            {
                Application.ShowAlertDialog("Não foi possível desenhar os pontos. Motivo:\r\n" + exception.Message);
            }
            catch (System.Exception exception)
            {
                System.Windows.Forms.ThreadExceptionDialog dialog = new System.Windows.Forms.ThreadExceptionDialog(exception);
                dialog.ShowDialog();
            }
            finally
            {
                this.transaction.Commit();
                this.document = null;
                this.database = null;
                this.transaction = null;
                this.blockTable = null;
                this.layerTable = null;
                this.dimStyleTable = null;
                this.textStyleTable = null;
                this.blockTableRecordModel = null;
                this.tabela = null;
            }
        }

        private void desenharPontosTopograficos(TabelaPontosTopograficos tabela)
        {
            if (tabela.representacaoPontoTipo == TipoRepresentacaoPonto.BlocoPadrao) criarBlocoPadraoSeNaoExistir(tabela);
            foreach (PontoTopografico ponto in tabela.pontosTopograficos)
            {
                double x1 = (ponto.leste - tabela.origemLeste) * tabela.multiplicadorDistancia;
                double y1 = (ponto.norte - tabela.origemNorte) * tabela.multiplicadorDistancia;
                Point3d ponto1 = new Point3d(x1, y1, 0.0).TransformBy(document.Editor.CurrentUserCoordinateSystem);
                double x2 = ponto1.X + tabela.leaderDeltaX;
                double y2 = ponto1.Y + tabela.leaderDeltaY;
                Point3d ponto2 = new Point3d(x2, y2, 0.0);
                string descricao = tabela.textoDescritivo(ponto);
                MText mText = criarMText(ponto2, descricao, tabela);
                Leader leader = criarLeader(ponto1, ponto2, mText, tabela);
                Entity representacaoPonto = criarRepresentacaoPonto(ponto1, tabela);
            }
        }

        private MText criarMText(Point3d ponto, string conteudo, TabelaPontosTopograficos tabela)
        {
            if (String.IsNullOrWhiteSpace(conteudo)) return null;
            MText mText = new MText();
            mText.Contents = conteudo;
            mText.TextStyleId = textStyleTable[tabela.mTextStyle];
            mText.Layer = tabela.mTextLayer;
            mText.Height = tabela.mTextHeight;
            mText.LineSpacingFactor = tabela.mTextLineSpaceFactor;
            if (tabela.leaderDeltaX < 0) mText.Attachment = AttachmentPoint.BottomRight; else mText.Attachment = AttachmentPoint.BottomLeft;
            mText.Location = ponto;
            blockTableRecordModel.AppendEntity(mText);
            transaction.AddNewlyCreatedDBObject(mText, true);
            return mText;
        }

        private Leader criarLeader(Point3d ponto1, Point3d ponto2, MText mText, TabelaPontosTopograficos tabela)
        {
            Leader leader = new Leader();
            leader.Layer = tabela.leaderLayer;
            leader.AppendVertex(ponto1);
            leader.AppendVertex(ponto2);
            leader.HasArrowHead = true;
            blockTableRecordModel.AppendEntity(leader);
            transaction.AddNewlyCreatedDBObject(leader, true);
            leader.Annotation = mText.ObjectId;
            leader.EvaluateLeader();
            return leader;
        }

        private Entity criarRepresentacaoPonto(Point3d ponto, TabelaPontosTopograficos tabela)
        {
            if (tabela.representacaoPontoTipo == TipoRepresentacaoPonto.Ponto) return criarDBPoint(ponto, tabela);
            if (tabela.representacaoPontoTipo == TipoRepresentacaoPonto.Bloco || tabela.representacaoPontoTipo == TipoRepresentacaoPonto.BlocoPadrao)
            {
                string nomeBloco = NOME_BLOCO_PADRAO;
                if (tabela.representacaoPontoTipo == TipoRepresentacaoPonto.Bloco) nomeBloco = tabela.representacaoPontoNomeBloco;
                return criarInsert(ponto, nomeBloco, tabela);
            }
            return null;
        }

        private DBPoint criarDBPoint(Point3d ponto, TabelaPontosTopograficos tabela)
        {
            DBPoint dbPoint = new DBPoint(ponto);
            dbPoint.Layer = tabela.representacaoPontoLayer;
            blockTableRecordModel.AppendEntity(dbPoint);
            transaction.AddNewlyCreatedDBObject(dbPoint, true);
            return dbPoint;
        }

        private BlockReference criarInsert(Point3d ponto, string nomeBloco, TabelaPontosTopograficos tabela)
        {
            BlockReference insert = new BlockReference(ponto, blockTable[nomeBloco]);
            insert.Layer = tabela.representacaoPontoLayer;
            double escala = tabela.representacaoPontoEscalaBloco.GetValueOrDefault();
            insert.ScaleFactors = new Scale3d(escala, escala, 1.0);
            blockTableRecordModel.AppendEntity(insert);
            transaction.AddNewlyCreatedDBObject(insert, true);
            return insert;
        }

        private void criarBlocoPadraoSeNaoExistir(TabelaPontosTopograficos tabela)
        {
            if (blockTable.Has(NOME_BLOCO_PADRAO)) return;

            BlockTableRecord blockTableRecord = new BlockTableRecord();
            blockTableRecord.Name = NOME_BLOCO_PADRAO;
            blockTableRecord.Origin = new Point3d(0, 0, 0);
            double raio = 0.5;
            double meiaLinha = 1.0;

            Circle circulo = new Circle();
            circulo.Center = new Point3d(0, 0, 0);
            circulo.Radius = raio;
            circulo.Layer = tabela.representacaoPontoLayer;
            blockTableRecord.AppendEntity(circulo);

            Line linhaHorizontal = new Line();
            linhaHorizontal.StartPoint = new Point3d(-meiaLinha, 0.0, 0.0);
            linhaHorizontal.EndPoint = new Point3d(meiaLinha, 0.0, 0.0);
            linhaHorizontal.Layer = tabela.representacaoPontoLayer;
            blockTableRecord.AppendEntity(linhaHorizontal);

            Line linhaVertical = new Line();
            linhaVertical.StartPoint = new Point3d(0.0, -meiaLinha, 0.0);
            linhaVertical.EndPoint = new Point3d(0.0, meiaLinha, 0.0);
            linhaVertical.Layer = tabela.representacaoPontoLayer;
            blockTableRecord.AppendEntity(linhaVertical);

            blockTable.UpgradeOpen();
            blockTable.Add(blockTableRecord);
            transaction.AddNewlyCreatedDBObject(blockTableRecord, true);
        }

        private void verificarValidadeDosDados(TabelaPontosTopograficos tabela)
        {
            verificarDimStyle(tabela.leaderDimStyle, "Leader");
            verificarLayer(tabela.leaderLayer, "Leader");
            verificarTextStyle(tabela.mTextStyle, "MText");
            verificarLayer(tabela.mTextLayer, "MText");
            if (tabela.representacaoPontoTipo != TipoRepresentacaoPonto.SemRepresentacao)
            {
                verificarLayer(tabela.representacaoPontoLayer, "Representação dos Pontos Topográficos");
            }
            if (tabela.representacaoPontoTipo == TipoRepresentacaoPonto.Bloco)
            {
                verificarBlock(tabela.representacaoPontoNomeBloco, "Representação dos Pontos Topográficos");
            }
        }

        private void verificarLayer(string nome, string nomeProprietario) => verificarExistenciaItemTable(nome, nomeProprietario, "Layer", layerTable);
        private void verificarDimStyle(string nome, string nomeProprietario) => verificarExistenciaItemTable(nome, nomeProprietario, "DimStyle", dimStyleTable);
        private void verificarTextStyle(string nome, string nomeProprietario) => verificarExistenciaItemTable(nome, nomeProprietario, "Style", textStyleTable);
        private void verificarBlock(string nome, string nomeProprietario) => verificarExistenciaItemTable(nome, nomeProprietario, "Bloco", blockTable);

        private void verificarExistenciaItemTable(string nomeItem, string nomeProprietario, string nomeTable, SymbolTable table)
        {
            if (!table.Has(nomeItem))
            {
                throw new CriadorDesenhoException($"O {nomeTable} '{nomeItem}', escolhido para {nomeProprietario}, não existe no desenho do AutoCAD.");
            }
        }
    }
}
