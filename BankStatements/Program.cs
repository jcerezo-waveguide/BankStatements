using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using Microsoft.Office.Interop.Excel;


namespace BankStatements
{
    class Program
    {
        static void Main(string[] args)
        {
            string FileLocation = @"C:\Users\jcerezo\Documents\Visual Studio 2015\Projects\BankStatements\BankStatements\TestFiles";

            Application Excel = null; 
            Workbook book = null;
            Worksheet sheet = null;

            try
            {
                List<Transaction> TransData = new List<Transaction>(ParseTransactionFiles(FileLocation));
                System.Data.DataTable dt = new System.Data.DataTable("Bank Transactions");
                
                dt.Columns.Add("Bank",typeof(string));
                dt.Columns.Add("Holder", typeof(string));
                dt.Columns.Add("Number", typeof(string));
                dt.Columns.Add("ID", typeof(string));
                dt.Columns.Add("Type", typeof(string));
                dt.Columns.Add("Amount", typeof(double));
                dt.Columns.Add("Date", typeof(DateTime));

                
                foreach (Transaction tran in TransData) {

                    System.Data.DataRow row = dt.NewRow();
                    row["Bank"] = tran.Bank;
                    row["Holder"] = tran.AccountHolder;
                    row["Number"] = tran.AccountNumber;
                    row["ID"] = tran.ID;
                    row["Type"] = tran.Type.ToString();
                    row["Amount"] = tran.Amount;
                    row["Date"] = tran.Date;
                    dt.Rows.Add(row);

                }
               
                if (TransData.Count > 0)
                {
                    Excel = new Application();
                    Excel.Visible = false;
                    

                    book = Excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    sheet = (Worksheet)book.Worksheets.get_Item(1);
                    
                    sheet.Name = dt.TableName;
                    

                    for (int i=1 ,j=0; i<= dt.Columns.Count ; i++,j++) {
                        sheet.Cells[1,i].Value2 = dt.Columns[j].ColumnName;
                    }

                    Range Header = sheet.Rows[1];
                    Header.Font.Bold = 1;
                    Header.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    Header.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    

                    for (int i=2 , r=0; i < dt.Rows.Count; i++,r++) {
                        for (int j=1,x=0; j<= dt.Columns.Count; j++,x++) {
                            sheet.Cells[i,j]= dt.Rows[r].ItemArray[x];
                        }
                    }
                    
                    for (int i = 1; i <= dt.Columns.Count; i++)
                    {
                        Range col = sheet.Columns[i];
                        col.EntireColumn.AutoFit();
                    }

                    string fout = Path.Combine(FileLocation, "ExcelTransactionsOut.xlsx");
                    if(File.Exists(fout)) File.Delete(fout);

                    book.SaveAs(fout);                    
                    book.Close(true, fout);
                    Excel.Quit();

                }
                else {
                    throw new ApplicationException("No files were found in:\n " + FileLocation);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
            finally {
            
                if (sheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                if (book != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
                if (Excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel);
            }           

        }

        private static List<Transaction> ParseTransactionFiles(string Path) {
            List<string> rf = new List<string>();
            List<Transaction> Trans = new List<Transaction>();

            string Extensions = "*.xml|*.csv|*.txt";

            if (Directory.Exists(Path)) {
                string[] exts = Extensions.Split('|');
                foreach (string fe in exts) {
                   string[] files = Directory.GetFiles(Path,fe.Trim());
                    if(files!=null) rf.AddRange(files);
                }
            }

            foreach (string file in rf){
                switch (System.IO.Path.GetExtension(file).ToLower())
                {
                    case ".csv":
                    case ".txt":
                        {
                            List<Transaction> csv = new List<Transaction>(ParseTransactionCSVFile(file));
                            if (csv.Count > 0) Trans.AddRange(csv);
                            break;
                        }
                    case ".xml":
                        {
                            List<Transaction> xml = new List<Transaction>(ParseTransactionXMLFile(file));
                            if (xml.Count > 0) Trans.AddRange(xml);
                            break;
                        }

                }
            }

            return Trans;
        }

        enum CVSP { Bank = 0, Holder = 1, Number = 2, ID = 3, Transation = 4, Amount = 6, Date = 7 };

        private static List<Transaction> ParseTransactionCSVFile(string file) {
            List<Transaction> Trans = new List<Transaction>();
            if ((file != null) && (File.Exists(file))) {

                try {
                    using (StreamReader fs = new StreamReader(file)){
                        string ln;
                        while ((ln = fs.ReadLine()) != null) {
                            string[] param = ln.Split(',');
                            for (int i = 0; i < param.Length; i++) { param[i].Trim(); }
                            Trans.Add(new Transaction(
                                param[(int)CVSP.Bank], param[(int)CVSP.Holder], param[(int)CVSP.Number],
                                param[(int)CVSP.ID], param[(int)CVSP.Transation], param[(int)CVSP.Amount],
                                param[(int)CVSP.Date]));

                        }
                    }
                    
                }
                catch (Exception ex)    {
                    throw new Exception("Error parsing csv file.", ex);
                }
                
            }
            return Trans;
        }

        private static List<Transaction> ParseTransactionXMLFile(string file)
        {
            List<Transaction> Trans = new List<Transaction>();
            if (file != null)
            {
                try {
                    XmlDocument xd = new XmlDocument();
                    xd.Load(file);
                    XmlNodeList xnl = xd.GetElementsByTagName("transaction");
                    foreach (XmlNode node in xnl) {
                        Trans.Add(new Transaction(
                            "Suntrust", node["AcctHolder"].InnerText, node["AcctNum"].InnerText, 
                            node["FedRefNum"].InnerText,node["DorW"].InnerText, node["Amount"].InnerText, 
                            node["DateTime"].InnerText
                            ));
                    }
                }
                catch (Exception ex) {
                    throw new Exception("Error parsing xml file.", ex);
                }
            }
            return Trans;
        }
    }

    public class Transaction
    {
        public enum TransactionType { Deposit, Withdraw };

        private string _Bank;
        private string _AcctHolder;
        private string _AcctNumber;
        private string _ID;
        private TransactionType _TransType;
        private double _Amount;
        private DateTime _Date;

        public string Bank { get { return _Bank; } }
        public string AccountHolder { get { return _AcctHolder; } }
        public string AccountNumber { get { return _AcctNumber; } }
        public string ID { get { return _ID; } }
        public TransactionType Type { get { return _TransType; } }
        public double Amount { get { return _Amount; } }
        public DateTime Date { get { return _Date; } }

        public Transaction(string bank, string accountHolder, string accountNumber, string id,
                            string transaction, string amount, string date)
        {
            if (bank == null) throw new ArgumentNullException("Bank");
            _Bank = bank;
            if (accountHolder == null) throw new ArgumentNullException("AccountHolder");
            _AcctHolder = accountHolder;
            if (accountNumber == null) throw new ArgumentNullException("AccountNumber");
            _AcctNumber = accountNumber;
            if (id == null) throw new ArgumentNullException("ID");
            _ID = id;
            if (double.TryParse(amount, out _Amount) == false)
            {
                throw new ArgumentException("Error trying to parse amount. Please verify that amount is a valid number", "Amount");
            }
            if (transaction == null)
            {
                _TransType = (_Amount >= 0) ? TransactionType.Deposit : TransactionType.Withdraw;

            }
            else {
                if (transaction.ToLower().IndexOf("d",0) >= 0 ||
                    transaction.IndexOf("+") >= 0
                )
                {
                    _TransType = TransactionType.Deposit;
                }
                else if (transaction.ToLower().IndexOf("w",0) >= 0 ||
                         transaction.IndexOf("-") >= 0
                    )
                {
                    _TransType = TransactionType.Withdraw;
                }
                else {
                    throw new ArgumentException("Error trying to parse transaction type", "Transaction");
                }
            }
            if (DateTime.TryParse(date, out _Date) == false) throw new ArgumentException("Error trying to parse transaction date. Please verify that date is a valid format", "Date");

            if ((_Amount > 0) && (_TransType == TransactionType.Withdraw))
            {
                _Amount *= -1;
            }

            if ((_Amount < 0) && (_TransType == TransactionType.Deposit))
            {
                throw new ArgumentException("Transation data is corrupted. A deposit transaction can't be a negative amount", "Transaction");
            }
        }


    }
}
