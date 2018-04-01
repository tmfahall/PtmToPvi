using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Data.OleDb;
using iTextSharp.text.pdf;
using System.IO;

namespace pviDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public class UnitClass
        {
            public string UnitNumber { get; set; }
            public string UnitMake { get; set; }
            public string UnitVin { get; set; }
            public string UnitYear { get; set; }
            public string UnitLicenseNumber { get; set; }
        }

        public class CustomerClass
        {
            public string CustomerNumber { get; set; }
            public string CustomerName { get; set; }
            public string CustomerStreetAddress { get; set; }
            public string CustomerStreetAddress2 { get; set; }
            public string CustomerCityStateZip { get; set; }
        }

        public static List<UnitClass> unitList(DataTable units)
        {
            var convertedList = (from rw in units.AsEnumerable()
                                 select new UnitClass()
                                 {
                                     UnitNumber = rw["UNIT"].ToString().Trim(),
                                     UnitMake = rw["MAKE"].ToString().Trim(),
                                     UnitVin = rw["VIN"].ToString().Trim(),
                                     UnitYear = rw["YEAR"].ToString().Trim(),
                                     UnitLicenseNumber = rw["LICENSE"].ToString().Trim()
                                 }).ToList();

            return convertedList;
        }

        public static List<CustomerClass> customerList(DataTable customer, string accountNumber)
        {
            var convertedList = (from rw in customer.AsEnumerable()
                                 select new CustomerClass()
                                 {
                                     CustomerNumber = accountNumber,
                                     CustomerName = rw["FLD1"].ToString().Trim(),
                                     CustomerCityStateZip = String.Format("{0}, {1}, {2}", rw["FLD4"].ToString().Trim(), rw["FLD5"].ToString().Trim(), rw["FLD6"].ToString().Trim()),
                                     CustomerStreetAddress = rw["FLD3"].ToString().Trim(),
                                     CustomerStreetAddress2 = rw["FLD3A"].ToString().Trim()
                                 }).ToList();

            return convertedList;
        }

        public DataTable getUnitInfoFromDb(string accountNumber)
        {
            OleDbConnection connection = new OleDbConnection(
            "Provider=VFPOLEDB.1;Data Source=Z:\\");

            DataTable Result = new DataTable();
            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();

                string mySQL = String.Format("SELECT UNIT, MAKE, VIN, YEAR, LICENSE FROM CUSTUNIT WHERE ACCTNO == '{0}'", accountNumber);

                OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

                DA.SelectCommand = MyQuery;

                DA.Fill(Result);

                connection.Close();
            }

            return Result;
        }

        public DataTable getAccountInfoFromDb(string accountNumber)
        {
            OleDbConnection connection = new OleDbConnection(
"Provider=VFPOLEDB.1;Data Source=Z:\\");

            DataTable Result = new DataTable();
            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();

                string mySQL = String.Format("SELECT FLD1, FLD3, FLD3A, FLD4, FLD5, FLD6 FROM CUST WHERE FLD2 == '{0}'", accountNumber);

                OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

                DA.SelectCommand = MyQuery;

                DA.Fill(Result);

                connection.Close();
            }

            return Result;
        }

        public DataTable getSpecificUnitInfo(string accountNumber, string unitNumber)
        {
            OleDbConnection connection = new OleDbConnection(
"Provider=VFPOLEDB.1;Data Source=Z:\\");

            DataTable Result = new DataTable();
            connection.Open();

            if (connection.State == ConnectionState.Open)
            {
                OleDbDataAdapter DA = new OleDbDataAdapter();

                string mySQL = String.Format("SELECT MAKE, VIN, YEAR, LICENSE FROM CUSTUNIT WHERE ACCTNO == '{0}' AND UNIT == '{1}'", accountNumber, unitNumber);

                OleDbCommand MyQuery = new OleDbCommand(mySQL, connection);

                DA.SelectCommand = MyQuery;

                DA.Fill(Result);

                connection.Close();
            }

            return Result;
        }

        public void writePdf(string date, string inspectionLocation, string inspectionCityStateZip, string timeIn, string amOrPm, string timeOut, string timeOutAmOrPm, string vehMake, string year, string vin, string unitNum, string lic, string ownerName, string ownerStreetAddress, string ownerCityStateZip, string carrierName, string carrierStreetAddress, string carrierCityStateZip, string ownerUsdotNum, string carrierUsDotNum, string inspectorName, string inspectorNumber)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string P_InputStream = String.Format("{0}\\PVI.pdf", desktop);
            string P_OutputStream = String.Format("{0}\\PVInew.pdf", desktop);

            PdfReader reader = new PdfReader((P_InputStream));
            using (PdfStamper stamper = new PdfStamper(reader, new FileStream(System.IO.Path.GetFullPath(P_OutputStream), FileMode.Create)))
            {
                AcroFields form = stamper.AcroFields;
                var fieldKeys = form.Fields.Keys;

                foreach (string fieldKey in fieldKeys)
                {
                    //Replace Address Form field with my custom data
                    if (fieldKey.Contains("1 Date"))
                    {
                        form.SetField(fieldKey, date);
                    }
                    if (fieldKey.Contains("2Insp"))
                    {
                        form.SetField(fieldKey, inspectionLocation);
                    }
                    if (fieldKey.Contains("3 City"))
                    {
                        form.SetField(fieldKey, inspectionCityStateZip);
                    }
                    if (fieldKey.Contains("4 Time in"))
                    {
                        form.SetField(fieldKey, timeIn);
                    }
                    if (amOrPm == "AM")
                    {
                        if (fieldKey.Contains("am1"))
                        {
                            form.SetField(fieldKey, "am");
                        }
                    }
                    if (fieldKey.Contains("6Veh Make"))
                    {
                        form.SetField(fieldKey, vehMake);
                    }
                    if (fieldKey.Contains("7Year"))
                    {
                        form.SetField(fieldKey, year);
                    }
                    if (fieldKey.Contains("8VIN"))
                    {
                        form.SetField(fieldKey, vin);
                    }
                    if (fieldKey.Contains("9 Unit"))
                    {
                        form.SetField(fieldKey, unitNum);
                    }
                    if (fieldKey.Contains("11Lic"))
                    {
                        form.SetField(fieldKey, lic);
                    }
                    if (fieldKey.Contains("14Owner Name"))
                    {
                        form.SetField(fieldKey, ownerName);
                    }
                    if (fieldKey.Contains("15Owner Street Address"))
                    {
                        form.SetField(fieldKey, ownerStreetAddress);
                    }
                    if (fieldKey.Contains("16 City State ZIP"))
                    {
                        form.SetField(fieldKey, ownerCityStateZip);
                    }
                    if (fieldKey.Contains("17Carrier Name"))
                    {
                        form.SetField(fieldKey, carrierName);
                    }
                    if (fieldKey.Contains("18Carrier Street Address"))
                    {
                        form.SetField(fieldKey, carrierStreetAddress);
                    }
                    if (fieldKey.Contains("19City State ZIP"))
                    {
                        form.SetField(fieldKey, carrierCityStateZip);
                    }
                    if (fieldKey.Contains("20Owner USDOT"))
                    {
                        form.SetField(fieldKey, ownerUsdotNum);
                    }
                    if (fieldKey.Contains("21 Carrier USDOT"))
                    {
                        form.SetField(fieldKey, carrierUsDotNum);
                    }
                    if (fieldKey.Contains("22Inspector Name"))
                    {
                        form.SetField(fieldKey, inspectorName);
                    }
                    if (fieldKey.Contains("23Inspector"))
                    {
                        form.SetField(fieldKey, inspectorNumber);
                    }
                }
                //The below will make sure the fields are not editable in
                //the output PDF.
                //stamper.FormFlattening = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(accountNumInputBox.Text))
            {
                string pattern = @"^(?<Prefix>\d{3})(?:[-\.\s]?)(?<Suffix>\d{4})(?!\d)";

                Match match = Regex.Match(accountNumInputBox.Text, pattern);

                if (match.Success)
                {
                    var accountNumberInput = accountNumInputBox.Text;

                    var unitInfoDataTable = getUnitInfoFromDb(accountNumberInput);
                    var listOfUnits = unitList(unitInfoDataTable);
                    var listOfUnitsSortedByUnitNumber = listOfUnits.OrderBy(u => u.UnitNumber).ToList();

                    unitSelector.DataSource = listOfUnitsSortedByUnitNumber;
                    unitSelector.DisplayMember = "Name";
                    unitSelector.ValueMember = "UnitNumber";

                    var accountInfoDataTable = getAccountInfoFromDb(accountNumberInput);
                    var listOfAccountProperties = customerList(accountInfoDataTable, accountNumberInput);

                    ownerInputBox.Text = listOfAccountProperties[0].CustomerName;
                    ownerCityStZipInputBox.Text = listOfAccountProperties[0].CustomerCityStateZip;
                    ownerStreetAddressInputBox.Text = listOfAccountProperties[0].CustomerStreetAddress;
                    carrierNameInputBox.Text = listOfAccountProperties[0].CustomerName;
                    carrierStreetAddressInputBox.Text = listOfAccountProperties[0].CustomerStreetAddress2;
                    carrierCityStZipInputBox.Text = listOfAccountProperties[0].CustomerCityStateZip;

                    inspectorNameInputBox.Items.Add("James");
                    inspectorNameInputBox.Items.Add("Kody");
                    inspectorNameInputBox.Items.Add("Kory");

                    if (accountNumberInput == "751-5413")
                    {
                        carrierUsdotNumInputBox.Text = "075545";
                        ownerUsdotNumInputBox.Text = "075545";
                    }
                }
                else
                {
                    ownerInputBox.Text = "PHONE NUM ERROR";
                }
            }


        }

        private void ownerIsCarrierButton_Click(object sender, EventArgs e)
        {
            carrierNameInputBox.Text = ownerInputBox.Text;
            carrierStreetAddressInputBox.Text = ownerStreetAddressInputBox.Text;
            carrierCityStZipInputBox.Text = ownerCityStZipInputBox.Text;
        }

        private void addToFormButton_Click(object sender, EventArgs e)
        {
            string date = DateTime.Now.ToString("MM/dd/yyyy");
            string inspectionLocation = "15635 US-2";
            string inspectionCityStateZip = "Bagley, MN, 56621";
            string timeIn = DateTime.Now.ToString("h:mm");
            string amOrPm = DateTime.Now.ToString("tt");
            string timeOut = "";
            string timeOutAmOrPm = "";
            string vehMake = makeInputBox.Text;
            string year = yearInputBox.Text;
            string vin = vinInputBox.Text;
            string unitNum = unitSelector.SelectedText;
            string lic = licenseInputBox.Text;
            string ownerName = ownerInputBox.Text;
            string ownerStreetAddress = ownerStreetAddressInputBox.Text;
            string ownerCityStateZip = ownerCityStZipInputBox.Text;
            string carrierName = carrierNameInputBox.Text;
            string carrierStreetAddress = carrierStreetAddressInputBox.Text;
            string carrierCityStateZip = carrierCityStZipInputBox.Text;
            string ownerUsdotNum = ownerUsdotNumInputBox.Text;
            string carrierUsDotNum = carrierUsdotNumInputBox.Text;
            string inspectorName = "";
            string inspectorNumber = "";

            if (inspectorNameInputBox.Text == "James")
            {
                inspectorName = "James Shongo";
                inspectorNumber = "101482";
            }

            if (inspectorNameInputBox.Text == "Kody")
            {
                inspectorName = "Kody Anderson";
                inspectorNumber = "133772";
            }

            if (inspectorNameInputBox.Text == "Kory")
            {
                inspectorName = "Kory Anderson";
                inspectorNumber = "926056";
            }

            writePdf(date, inspectionLocation, inspectionCityStateZip, timeIn, amOrPm, timeOut, timeOutAmOrPm, vehMake, year, vin, unitNum, lic, ownerName, ownerStreetAddress, ownerCityStateZip, carrierName, carrierStreetAddress, carrierCityStateZip, ownerUsdotNum, carrierUsDotNum, inspectorName, inspectorNumber);
        }

        private void unitSelector_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string accountNumber = accountNumInputBox.Text;
            string unitNumber = unitSelector.SelectedText.ToString();

            var unitTable = getSpecificUnitInfo(accountNumber, unitNumber);
            string make = unitTable.Rows[0][0].ToString();
            string vin = unitTable.Rows[0][1].ToString();
            string year = unitTable.Rows[0][2].ToString();
            string license = unitTable.Rows[0][3].ToString();

            makeInputBox.Text = make;
            yearInputBox.Text = year;
            vinInputBox.Text = vin;
            licenseInputBox.Text = license;

        }

        private void unitSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string accountNumber = accountNumInputBox.Text;
                string unitNumber = unitSelector.SelectedText.ToString();

                var unitTable = getSpecificUnitInfo(accountNumber, unitNumber);
                string make = unitTable.Rows[0][0].ToString();
                string vin = unitTable.Rows[0][1].ToString();
                string year = unitTable.Rows[0][2].ToString();
                string license = unitTable.Rows[0][3].ToString();

                makeInputBox.Text = make;
                yearInputBox.Text = year;
                vinInputBox.Text = vin;
                licenseInputBox.Text = license;
            }
            catch (Exception)
            {
                
            }
        }
    }
}