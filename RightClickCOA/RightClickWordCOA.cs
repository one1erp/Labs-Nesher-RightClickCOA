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
    [ProgId("RightClickCOA.RightClickWordCOA")]
    public class RightClickWordCOA : IEntityExtension
    {
        private INautilusDBConnection _ntlsCon;
        INautilusServiceProvider sp;
        private double coaId;
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
                    _wordPath = CurrentCOA.DocPath;

                    if (CurrentCOA != null && _wordPath != null)
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
                        MessageBox.Show("מסמך אינו קיים");
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("ERROR ON COA Report id: "+ id +". The error is: "+ e.Message);
                }
                
            }
            
        }

        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            return ExecuteExtension.exInvisible;

        }
    }
}
