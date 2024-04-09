using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Common;
using DAL;
using LSExtensionControlLib;
using LSSERVICEPROVIDERLib;
using LSEXT;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

namespace RightClickCOA
{
    [ComVisible(true)]
    [ProgId("RightClickCOA.RightClickPDF-WordCOA")]
    public class RightClickPDF_WordCOA : IEntityExtension
    {
        private INautilusDBConnection _ntlsCon;
        INautilusServiceProvider sp;
        private double coaId;
        string _pdfPath = "";
        string _wordPath = "";
        private IDataLayer dal;
        private COA_Report CurrentCOA;



        public void Execute(ref LSExtensionParameters Parameters)
        {
            sp = Parameters["SERVICE_PROVIDER"];
            var records = Parameters["RECORDS"];

            _ntlsCon = Utils.GetNtlsCon(sp);
            Utils.CreateConstring(_ntlsCon);

            dal = new DataLayer();
            dal.Connect();

            List<string> Ids = new List<string>();
            while (!records.EOF)
            {
                var id = records.Fields["U_COA_REPORT_ID"].Value;
                Ids.Add(id);
                records.MoveNext();
            }

            foreach (var id in Ids)
            {
                try
                {
                    CurrentCOA = dal.GetCoaReportById(Convert.ToInt64(id));
                    if (CurrentCOA != null)
                    {
                        _pdfPath = CurrentCOA.PdfPath;
                    }
                    if (_pdfPath != null && File.Exists(_pdfPath))
                    {
                        Process p = Process.Start(_pdfPath);
                        System.Threading.Thread.Sleep(1000);
                        p.WaitForExit();
                    }
                    else
                    {
                        _wordPath = CurrentCOA.DocPath;


                        if (!openFile(_wordPath))
                        {
                            if (CurrentCOA.Name.Substring(0, 2).Equals("17") && CurrentCOA.Sdg.LabInfo.LabLetter.Equals("C"))
                            {
                                List<string> files = new List<string>();
                                string fileName;
                                if(_pdfPath != null)
                                {
                                    fileName = Path.Combine(@"\\micro-lims\coa_documents-h\Cosmetics\cosmetics 2016-2017", _pdfPath.Substring(_pdfPath.LastIndexOf(@"\") + 1));

                                    if (File.Exists(fileName) && openFile(fileName))
                                    {
                                        return;
                                    }
                                }
                                else if (_wordPath != null)
                                {
                                    fileName = Path.Combine(@"\\micro-lims\coa_documents-h\Cosmetics\cosmetics 2016-2017", _wordPath.Substring(_pdfPath.LastIndexOf(@"\") + 1));

                                    if (File.Exists(fileName) && openFile(fileName))
                                    {
                                        return;
                                    }
                                }
                            }

                            MessageBox.Show("המסמכים לא קיימים");
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("ERROR ON COA Report id: " + id + ". The error is: " + e.Message);
                }
            }
        }


        private bool openFile(string filePath)
        {
            if (filePath != null && File.Exists(filePath))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                var lastMod = fileInfo.LastWriteTime;
                var createdStatus = CurrentCOA.Status == "C";

                Process p = Process.Start(filePath);
                System.Threading.Thread.Sleep(1000);
                p.WaitForExit();

                if (createdStatus)
                {
                    FileInfo newfileInfo = new FileInfo(CurrentCOA.DocPath);

                    //אם נעשו שינויים במסמך
                    if (!DateTime.Equals(lastMod, newfileInfo.LastWriteTime))
                    {
                        CurrentCOA.Status = "E";
                        dal.SaveChanges();
                    }
                }

                return true;
            }

            return false;
        }
        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exEnabled;
        }
    }
}
