using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Interop;

[assembly: CommandClass(typeof(MyFirstProject.Class1))]

namespace MyFirstProject
{
    public class Class1
    {
        private Document _acDoc = Application.DocumentManager.MdiActiveDocument;

        [CommandMethod("OPENALL")]
        public async void OpenAll()
        {
            OpenFile();
            await Task.Delay(1000);
            AdskGreeting();
        }

       
        public async void AdskGreeting()
        {
            string ip = string.Empty;
            Editor ed = _acDoc.Editor;
           

            try
            {
                ip = await GetExternalIpAsync();

                using (DocumentLock acLckDoc = _acDoc.LockDocument())
                {
                    Database acCurDb = _acDoc.Database;

                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        BlockTable acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                        using (MText objText = new MText())
                        {
                            objText.Location = new Point3d(2, 2, 0);
                            objText.Contents = $"IP: {ip}";
                            objText.TextStyleId = acCurDb.Textstyle;

                            acBlkTblRec.AppendEntity(objText);
                            acTrans.AddNewlyCreatedDBObject(objText, true);

                            Point3d basePoint = objText.Location;
                            double angle = -90.5 * (Math.PI / 180);
                            Matrix3d rotationMatrix = Matrix3d.Rotation(angle, Vector3d.ZAxis, basePoint);
                            objText.TransformBy(rotationMatrix);
                        }
                        acTrans.Commit();
                    }
                    ed.WriteMessage($"\nIP-адрес '{ip}' добавлен на чертеж.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nОшибка выполнения команды ADSKGREETING: {ex.Message}");
                System.Diagnostics.Trace.WriteLine($"ADSKGREETING Error: {ex.ToString()}");
            }
        }

        private async Task<string> GetExternalIpAsync()
        {
            try
            {
                using HttpClient client = new HttpClient();
                string ip = await client.GetStringAsync("https://ipv4.icanhazip.com");
                return ip.Trim();
            }
            catch (System.Exception ex)
            {
                return $"Ошибка получения IP: {ex.Message}";
            }
        }

        public void OpenFile()
        {
            Editor ed = _acDoc.Editor;

            try
            {
                string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                string localDwgPath = Path.Combine(downloadsPath, "cold-rolled-steel-production.dwg");

                if (!File.Exists(localDwgPath))
                {
                    ed.WriteMessage($"\nФайл не найден: {localDwgPath}");
                    return;
                }

                if (System.Windows.Application.Current == null)
                {
                    new System.Windows.Application();
                }

                var loadingWindow = new LoadingProgressWindow();
                var acadHandle = Application.MainWindow.Handle;
                new WindowInteropHelper(loadingWindow).Owner = acadHandle;
                loadingWindow.Show();

                Application.DocumentManager.Open(localDwgPath, false);

                loadingWindow.Close();
                ed.WriteMessage($"\nФайл открыт: {localDwgPath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nОшибка открытия файла: {ex.Message}");
            }
        }

    }
}
