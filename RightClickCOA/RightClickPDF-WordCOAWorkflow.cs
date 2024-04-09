using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RightClickCOA
{
    [ComVisible(true)]
    [ProgId("RightClickCOA.RightClickPDF-WordCOAWorkflow")]
    public class RightClickPDF_WordCOAWorkflow : IWorkflowExtension
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

            //List<string> Ids = new List<string>();
            records.MoveFirst();
            //   while (!records.EOF)
            // {
            var id = records.Fields["U_COA_REPORT_ID"].Value;
            string sID = id.ToString();
            //Ids.Add(sID);
            //records.MoveNext();
            //  }

            // foreach (var id in Ids)
            //{
            try
            {
                CurrentCOA = dal.GetCoaReportById(Convert.ToInt64(sID));
                if (CurrentCOA != null)
                {
                    _pdfPath = CurrentCOA.PdfPath;
                }
                if (_pdfPath != null)
                {
                    Process p = Process.Start(_pdfPath);
                    System.Threading.Thread.Sleep(1000);
                    p.WaitForExit();
                }
                else
                {
                    _wordPath = CurrentCOA.DocPath;
                    if (_wordPath != null)
                    {
                        FileInfo fileInfo = new FileInfo(_wordPath);
                        var lastMod = fileInfo.LastWriteTime;
                        var createdStatus = CurrentCOA.Status == "C";

                        Process p = Process.Start(_wordPath);
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

                    }
                    else
                    {
                        MessageBox.Show("המסמכים לא קיימים");
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("ERROR ON COA Report id: " + id + ". The error is: " + e.Message);
            }
            // }
        }
    }
}
