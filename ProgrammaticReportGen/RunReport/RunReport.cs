using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using System.IO;
using SharePointDataSourceDriver;
using Kailua.net.windward.utils.ado.SPList.custom;
using net.windward.api.csharp;
using WindwardInterfaces.net.windward.api.csharp;
using System.Collections.Generic;
using Microsoft.SharePoint.Utilities;
using System.Collections;

namespace ProgrammaticReportGen.RunReport
{
    [ToolboxItemAttribute(false)]
    public class RunReport : WebPart
    {
        Label lblZipCode = new Label();
        TextBox txtZipCode = new TextBox();
        Button btnRunReport = new Button();

        protected override void OnInit(EventArgs e)
        {
            txtZipCode.ID = "txtZipCode";
            lblZipCode.ID = "lblZipCode";
            btnRunReport.ID = "btnRunReport";

            base.OnInit(e);
        }
        protected override void CreateChildControls()
        {

            this.Controls.Add(lblZipCode);
            this.Controls.Add(txtZipCode);
            this.Controls.Add(btnRunReport);
            btnRunReport.Text = @"Run Report";
            btnRunReport.Click += new EventHandler(btnRunReport_Click);

            base.CreateChildControls();

            // Set Windward license
            java.util.Properties props = java.lang.System.getProperties();
            props.put(@"wr.license", @"bHtzLjSa8TwnWysGZbHlvj2B5T53uzsdJzqxHD1Zu61kYWW9LXD9j2wyRxZ8vGk2dpjpPHYSRQQ9x2W0J7v5PCfI7z6nyjs3ZonxjCa76QQ8yLOtZ4M1jTnl9z08UHEUdKB/nnxSJQ52ElU1PcXlNH27ORY3yzk/9LszJ6a5aYy1mykFN8s5P3UwuSdm+eE9JjJFPj3F5TR9uzkWd9mpPnX68w5neauUNw67liZZu7cvGms8bbs7LHfZqT71YB8OPcH1FH/Iu64nKikFLIMtLn0iqbY2mXsfdjJHPT3F5TR9O/kWNVm5pWaedZx+m+GWNdPzHWTpsaR22Kk9ZbHxDncK+Y9ngfcVN2D5nybKsTw3Wm+/fwrplGbIpTxuW3UUZ/rzPWWxoYcnCvEMN9g3l2f6dwR3eGO9PchplH/KuT1nSvGXbyr5PWWxsZ8nCvEMN8i3Fz6bMRZni/mevZvzLTeB9xdnSPmMJtrxPD3Ley01wzM853j5DTdrKRUmQL8fPcMzBeep2Y0mRL8PPcMzBOb4uY02m3+fPYHlFv1gDz89wfUUL/h7rC6r2RR2QL8fNtr7Dj3IaZQ+mzE/Z4v5nj3KOy0nOuEdPchplD9Y8T93CLuMZaPxP3ca4T810/MMZXo5JGabfZw90+MUZfrxJCbKaRxlejkMZkC9HDzJoZR22Kk9ZaMxDvab8R9mq/OEd7x7HWeB5TxnSPmMJtrxPD3Ley01wzM85rvxDWar84RloaWHPYH1D2daST4nWXMVZpttvC2B5RZnWIk+Yp99hn2B5RZnSsm+N9j7NiZAvx49wzMM5vi5jTabf589geUW/WAPPzxSNRQnu/k8J8ovvjh2Nz0nKuENJovhDbnBxSw=");
            java.lang.System.setProperties(props);
            Report.Init();
        }
        void btnRunReport_Click(object sender, EventArgs e)
        {
            // Get template
            string templateUrl = @"http://wwr-test/bpa-test/documents/ContactReportTemplate.docx";
            Stream templateStream;
            using (SPSite siteCollection = new SPSite(templateUrl))
            {
                using (SPWeb web = siteCollection.OpenWeb())
                {
                    templateStream = new MemoryStream(web.GetFile(templateUrl).OpenBinary(), false);
                }
            }

            // ProcessSetup()
            Stream reportStream = new MemoryStream();
            Report reportGen = new ReportDocx(templateStream, reportStream);

            // ProcessData()
            IReportDataSource irds = new SharePointDataSource(typeof(SPListConnection), @"URL=http://wwr-test/bpa-test/;USER=rovisys;PASSWORD=1455Danner;");

            // Process Setup
            reportGen.ProcessSetup();

            //This is where we pass in the parameters    
                Dictionary<string, object> map = new Dictionary<string, object>();
                map.Add("ZipCode", txtZipCode.Text);
                //This is the function where we actually tell our report the parameter values
                irds.Map = map;
          


            // The second parameter is the name of the data source
            reportGen.ProcessData(irds, "");

            // ProcessComplete()
            reportGen.ProcessComplete();
            reportStream.Flush();
            templateStream.Flush();
            reportGen.Dispose();

            // Save the report
            string reportUrl = @"http://wwr-test/bpa-test/documents/ContactReport.docx";
            using (SPSite siteCollection = new SPSite(reportUrl))
            {
                using (SPWeb web = siteCollection.OpenWeb())
                {
                    web.AllowUnsafeUpdates = true;
                    web.Files.Add(reportUrl, reportStream, true);
                    web.AllowUnsafeUpdates = false;
                }
            }

            lblZipCode.Text = "Zip Code";
        }
    }
}
