using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel.Security;
using System.Windows.Forms;
using System.Xml;
using Sanama.SanamaService;

namespace Sanama
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

        string sFileName, format;
        int iRow, iCol = 2;

        // OPEN FILE DIALOG AND SELECT AN EXCEL FILE.
        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            { 
                OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
                OpenFileDialog1.Title = "لطفا فایل اطلاعات خود را انتخاب کنید.";
                OpenFileDialog1.FileName = "";
                OpenFileDialog1.Filter = "File|*.xlsx;*.xls;*.xml";

                OpenFileDialog OpenFileDialog2 = new OpenFileDialog();
                OpenFileDialog2.Title = "لطفا زوج کلید خود را انتخاب کنید.";
                OpenFileDialog2.FileName = "";
                OpenFileDialog2.Filter = "key pair|*.p12;*.pfx;";

                if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if(OpenFileDialog2.ShowDialog() == DialogResult.OK)
                    {
                        var keypairPassword = ShowDialog("پسورد کلید خود را وارد کنید", "");
                        sFileName = OpenFileDialog1.FileName;
                        string ext = Path.GetExtension(OpenFileDialog1.FileName);
                        var sanamaInof = new sanamaInfo();
                        if(ext == ".xls" || ext == ".xlsx")
                        {
                            throw new Exception("not implement yet...");
                        }
                        else if (ext == ".xml")
                        {
                            sanamaInof = readXml(sFileName);
                        }
                        else
                        {
                            throw new Exception("فرمت فایل شما اشتباه است.");
                        }
                        var client = new PushSanamaClient();
                        client.ClientCredentials.ServiceCertificate.DefaultCertificate = new X509Certificate2("bitaServicebusCert.cer");
                        client.ClientCredentials.ClientCertificate.Certificate = new X509Certificate2(OpenFileDialog2.FileName, keypairPassword, X509KeyStorageFlags.PersistKeySet);
                        client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = X509CertificateValidationMode.None;
                        var result = client.send(sanamaInof);
                        MessageBox.Show(result, "اطلاعات با موفقیت ارسال شد.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static string ShowDialog(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
            System.Windows.Forms.Label textLabel = new System.Windows.Forms.Label() { Left = 50, Top = 20, Text = text };
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox() { Left = 50, Top = 50, Width = 400 };
            System.Windows.Forms.Button confirmation = new System.Windows.Forms.Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }

        private sanamaInfo readXml(string sFile)
        {
            var sanamaInfo = createSanamaInfo(sFile);
            return sanamaInfo;
        }

        private sanamaInfo createSanamaInfo(string sFile)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(sFile);
            XmlNodeList sanamaInfoList = xmlDocument.SelectNodes("SanamaInfo");
            if (sanamaInfoList.Count != 1)
            {
                throw new Exception("فایل مورد نظر باید تنها شامل 1 SanamaInfo باشد.");
            }
            var sanamaInfoXml = sanamaInfoList.Item(0);
            var sanamaInfo = new sanamaInfo();
            sanamaInfo.mainOrgCode = sanamaInfoXml.Attributes["MainOrgCode"] == null ? null : sanamaInfoXml.Attributes["MainOrgCode"].Value;
            sanamaInfo.mainOrgID = sanamaInfoXml.Attributes["MainOrgID"] == null ? null : sanamaInfoXml.Attributes["MainOrgID"].Value;
            sanamaInfo.month = getXmlElementAttributeValueAsInt32(sanamaInfoXml.Attributes["Month"]);
            sanamaInfo.protocolName = sanamaInfoXml.Attributes["ProtocolName"] == null ? null : sanamaInfoXml.Attributes["ProtocolName"].Value;
            sanamaInfo.protocolType = sanamaInfoXml.Attributes["ProtocolType"] == null ? null : sanamaInfoXml.Attributes["ProtocolType"].Value;
            sanamaInfo.protocolVer = sanamaInfoXml.Attributes["ProtocolVer"] == null ? null : sanamaInfoXml.Attributes["ProtocolVer"].Value;
            sanamaInfo.year = getXmlElementAttributeValueAsInt32(sanamaInfoXml.Attributes["Year"]);
            var attachments = new List<attachment>();
            var contrastAccounts = new List<contrastAccount>();
            var reports = new List<report>();
            foreach (XmlElement xmlElement in sanamaInfoXml.SelectNodes("Report_List"))
            {
                var report = new report();
                report.accCode = xmlElement.Attributes["AccCode"] == null ? null : xmlElement.Attributes["AccCode"].Value;
                report.summaryProgressDeptor = getXmlElementAttributeValueAsLong(xmlElement.Attributes["SummaryProgressDeptor"]);
                report.summaryProgressCreditor = getXmlElementAttributeValueAsLong(xmlElement.Attributes["SummaryProgressCreditor"]);
                report.sourceType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["SourceType"]);
                report.sourceEssence = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["SourceEssence"]);
                report.otherSourceType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["OtherSourceType"]);
                report.creditType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["CreditType"]);
                report.transferalType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["TransferalType"]);
                report.creditInfo = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["CreditInfo"]);
                report.expenseArticle = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["ExpenseArticle"]);
                report.constructArticle = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["ConstructArticle"]);
                report.expenseDetailArticle = xmlElement.Attributes["ExpenseDetailArticle"] == null ? null : xmlElement.Attributes["ExpenseDetailArticle"].Value;
                report.incomeCode = xmlElement.Attributes["IncomeCode"] == null ? null : xmlElement.Attributes["IncomeCode"].Value;
                report.incomeSubject = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["IncomeSubject"]);
                report.governmental = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["Governmental"]);
                report.taxSeason = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["TaxSeason"]);
                report.debentureSenderRank = xmlElement.Attributes["DebentureSenderRank"] == null ? null : xmlElement.Attributes["DebentureSenderRank"].Value;
                report.debentureReceiverRank = xmlElement.Attributes["DebentureReceiverRank"] == null ? null : xmlElement.Attributes["DebentureReceiverRank"].Value;
                report.costCenter = xmlElement.Attributes["CostCenter"] == null ? null : xmlElement.Attributes["CostCenter"].Value;
                report.awardArticle = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["AwardArticle"]);
                report.securitiesType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["SecuritiesType"]);
                report.guaranteeEssence = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["GuaranteeEssence"]);
                report.year = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["Year"]);
                report.nominee = xmlElement.Attributes["Nominee"] == null ? null : xmlElement.Attributes["Nominee"].Value;
                report.demandStatus = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["DemandStatus"]);
                report.tempPaymentType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["TempPaymentType"]);
                report.leakageSubject = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["LeakageSubject"]);
                report.assuranceType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["AssuranceType"]);
                report.assuranceSubject = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["AssuranceSubject"]);
                report.currencyType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["CurrencyType"]);
                report.accountNumber = xmlElement.Attributes["AccountNumber"] == null ? null : xmlElement.Attributes["AccountNumber"].Value;
                report.debitSubject = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["DebitSubject"]);
                report.fixedAssetType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["FixedAssetType"]);
                report.inventoryType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["InventoryType"]);
                report.quantity = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["Quantity"]);
                report.dueDate = xmlElement.Attributes["DueDate"] == null ? null : xmlElement.Attributes["DueDate"].Value;
                report.securitiesProperties = xmlElement.Attributes["SecuritiesProperties"] == null ? null : xmlElement.Attributes["SecuritiesProperties"].Value;
                report.investmentType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["InvestmentType"]);
                report.annualAdjustmentsSubject = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["AnnualAdjustmentsSubject"]);
                reports.Add(report);
            }
            foreach (XmlElement xmlElement in sanamaInfoXml.SelectNodes("ContrastAccount_list"))
            {
                var contrastAccount = new contrastAccount();
                contrastAccount.accountNumber = xmlElement.Attributes["AccountNumber"] == null ? null : xmlElement.Attributes["AccountNumber"].Value;
                contrastAccount.accountDscp = xmlElement.Attributes["AccountDscp"] == null ? null : xmlElement.Attributes["AccountDscp"].Value;
                contrastAccount.accountType = getXmlElementAttributeValueAsInt32(xmlElement.Attributes["AccountType"]);
                contrastAccount.mojoodiTebgheDaftar = getXmlElementAttributeValueAsLong(xmlElement.Attributes["MojoodiTebgheDaftar"]);
                contrastAccount.mojoodiTebgheBank = getXmlElementAttributeValueAsLong(xmlElement.Attributes["MojoodiTebgheBank"]);

                contrastAccount.diffType1 = createDiffType(xmlElement, "DiffType1");
                contrastAccount.diffType2 = createDiffType(xmlElement, "DiffType2");
                contrastAccount.diffType3 = createDiffType(xmlElement, "DiffType3");
                contrastAccount.diffType4 = createDiffType(xmlElement, "DiffType4");
                contrastAccount.diffType5 = createDiffType(xmlElement, "DiffType5");
                contrastAccount.diffType6 = createDiffType(xmlElement, "DiffType6");
                contrastAccount.diffType7 = createDiffType(xmlElement, "DiffType7");
                contrastAccount.diffType8 = createDiffType(xmlElement, "DiffType8");
                contrastAccount.diffType9 = createDiffType(xmlElement, "DiffType9");
                contrastAccount.diffType10 = createDiffType(xmlElement, "DiffType10");
                contrastAccount.diffType11 = createDiffType(xmlElement, "DiffType11");
                contrastAccount.diffType12 = createDiffType(xmlElement, "DiffType12");
                contrastAccounts.Add(contrastAccount);
            }
            sanamaInfo.contrastAccounts = contrastAccounts.ToArray();
            sanamaInfo.reports = reports.ToArray();
            return sanamaInfo;
        }

        private diffType createDiffType(XmlElement xmlElement, string v)
        {
            var diffType = new diffType();
            XmlNode DiffTypexmlElement = xmlElement.SelectSingleNode(v);
            diffType.value = getXmlElementAttributeValueAsLong(DiffTypexmlElement.Attributes["Value"]);
            var details = new List<detail>();
            foreach (XmlElement xmlElement1 in DiffTypexmlElement.SelectNodes("Detail_List"))
            {
                var detail = new detail();
                detail.date = xmlElement.Attributes["Date"] == null ? null : xmlElement.Attributes["Date"].Value;
                detail.description = xmlElement.Attributes["Description"] == null ? null : xmlElement.Attributes["Description"].Value;
                detail.expense = getXmlElementAttributeValueAsLong(xmlElement.Attributes["Expense"]);
                details.Add(detail);
            }
            diffType.details = details.ToArray();
            return diffType;
        }

        // GET DATA FROM EXCEL AND POPULATE COMB0 BOX.
        private void readExcel(string sFile)
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sFile);           // WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets["Sheet1"];      // NAME OF THE SHEET.

            for (iRow = 2; iRow <= xlWorkSheet.Rows.Count; iRow++)  // START FROM THE SECOND ROW.
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {               // POPULATE COMBO BOX.
                    Console.WriteLine(xlWorkSheet.Cells[iRow, 1].value);
                }
            }

            xlWorkBook.Close();
            xlApp.Quit();
        }

        private Int32 getXmlElementAttributeValueAsInt32(XmlAttribute xmlElement)
        {
            if (xmlElement == null)
            {
                return 0;
            }
            else if (xmlElement.Value == "")
            {
                return 0;
            }
            else
            {
                return Int32.Parse(xmlElement.Value);
            }
        }

        private long getXmlElementAttributeValueAsLong(XmlAttribute xmlElement)
        {
            if (xmlElement == null)
            {
                return 0;
            }
            else if (xmlElement.Value == "")
            {
                return 0;
            }
            else
            {
                return long.Parse(xmlElement.Value);
            }
        }
    }
}
