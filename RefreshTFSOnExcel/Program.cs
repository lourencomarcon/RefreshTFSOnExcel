using NLog;
using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace RefreshTFSOnExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            LogConfiguration();

            var logger = LogManager.GetCurrentClassLogger();

            if (args.Length < 2)
            {
                logger.Error($"Não foram informados os parâmetros");
                return;
            }

            Excel.Application xlApp = new Excel.Application();

            xlApp.DisplayAlerts = false;

            Excel.Workbook xlWorkBook = null;

            string fileName = args[0];

            string[] worksheets = args[1].Split(';').Select(p => p.Trim()).ToArray();

            logger.Info("Iniciando a rotina de atualização automática do excel com query do TFS");

            try
            {
                logger.Info($"Abrindo a planilha '{fileName}'");

                xlWorkBook = xlApp.Workbooks.Open(fileName);

                foreach (string worksheet in worksheets)
                {
                    logger.Info($"Atualizando a aba '{worksheet}'");

                    string message = xlApp.Run("RefreshTeamQueryOnWorksheet", worksheet);

                    if (message != "Sucess")
                        logger.Error($"Ocorreu um erro ao atualizar a aba '{worksheet}'. Erro: {message}");
                }

                logger.Info($"Salvando a planilha '{fileName}'");

                xlWorkBook.Save();

                logger.Info($"Fechando a planilha '{fileName}'");

                xlWorkBook.Close(false);
            }
            catch (Exception e)
            {
                logger.Error($"Ocorreu um erro ao atualizar a planilha '{fileName}'. Erro: {e}");
            }
            finally
            {
                xlApp.Quit();
                releaseObject(xlApp);
                releaseObject(xlWorkBook);
            }

            logger.Info("Finalizando a rotina de atualização automática do excel com query do TFS");
        }

        private static void LogConfiguration()
        {
            var config = new NLog.Config.LoggingConfiguration();

            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = "Log.txt" };

            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);

            LogManager.Configuration = config;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
