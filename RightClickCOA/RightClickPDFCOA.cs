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
    [ProgId("RightClickCOA.RightClickPDFCOA")]
    public class RightClickPDFCOA : IEntityExtension 
    {
        private INautilusDBConnection _ntlsCon;
        INautilusServiceProvider sp;
        private double coaId;
        string _pdfPath = "";
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

                    if (CurrentCOA != null && _pdfPath != null)
                    {
                        Process p = Process.Start(_pdfPath);
                        System.Threading.Thread.Sleep(1000);
                        p.WaitForExit();
                    }
                    else
                    {
                        MessageBox.Show("מסמך אינו קיים");
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("ERROR ON COA Report id: " + id + ". The error is: " + e.Message);
                }
                
            }            
        }

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exInvisible;
        }

    }
}
