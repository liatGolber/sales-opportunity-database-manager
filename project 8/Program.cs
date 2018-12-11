using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace project_8
{
    public struct User
    {
        public string name;
        public string lastN;
        public string ID;
        public string password;
        public bool isAdmin;
    }

    public struct Opp
    {
        public string name;
        public string lastN;
        public string ID;
        public string phone;
        public string email;
        public DateTime treatedAt;
        public string status;
        public User treatedBy;
        public string comment;
        public string hID;//for history uses
    }

    public struct Package
    {
        public string ID;
        public string lineNum;
        public int packageType;
        public DateTime startD;
        public DateTime endD;
        public string hID;//for history uses
    }

    static class Program
    {
        //vars
        //databases paths
        public static string userDB = Environment.CurrentDirectory + @"\Users.xlsx";//1-id 2-name 3-last 4-password 5-isadmin
        public static string historyDB = Environment.CurrentDirectory + @"\History.xlsx";
        public static string leadsDB = Environment.CurrentDirectory + @"\Leads.xlsx";
        public static string opportunitesDB = Environment.CurrentDirectory + @"\Opportunites.xlsx";

        public static List<User> userList;
        public static List<Opp> opportunites;
        public static List<Package> packages;
        public static User currentUser;

        //functions


        public static Dictionary<string, object> MakeExcelEss(string dbPath, int sheetNum = 1)
        {
            Dictionary<string, object> ret = new Dictionary<string, object>();
            ret["App"] = new Excel.Application();
            ret["Book"] = null;
            // the reference to the worksheet,
            // we'll assume the first sheet in the book.
            ret["Sheet"] = null;
            ret["Range"] = null;
            // the range object is used to hold the data
            // we'll be reading from and to find the range of data.
            (ret["App"] as Excel.Application).Visible = false;
            (ret["App"] as Excel.Application).ScreenUpdating = false;
            (ret["App"] as Excel.Application).DisplayAlerts = false;
            ret["Book"] = (ret["App"] as Excel.Application).Workbooks.Open(dbPath,
       Type.Missing, Type.Missing, Type.Missing,
       Type.Missing, Type.Missing, Type.Missing, Type.Missing,
       Type.Missing, Type.Missing, Type.Missing, Type.Missing,
       Type.Missing, Type.Missing, Type.Missing);
            ret["Sheet"] = (Excel.Worksheet)(ret["Book"] as Excel.Workbook).Worksheets[sheetNum];
            //ret["Range"] = (ret["Sheet"] as Excel.Worksheet).get_Range("A1", Type.Missing);
            //ret["Range"] = (ret["Range"] as Excel.Range).get_End(Excel.XlDirection.xlToRight);
            //ret["Range"] = (ret["Range"] as Excel.Range).get_End(Excel.XlDirection.xlDown);
            //string downAddress = (ret["Range"] as Excel.Range).get_Address(
            //    false, false, Excel.XlReferenceStyle.xlA1,
            //    Type.Missing, Type.Missing);
            //ret["Range"] = (ret["Sheet"] as Excel.Worksheet).get_Range("A1", downAddress);
            ret["Range"] = (ret["Sheet"] as Excel.Worksheet).UsedRange;
            return ret;
        }

        public static void CleanExcelEss(ref Dictionary<string, object> ret)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(ret["Range"]);
            Marshal.ReleaseComObject(ret["Sheet"]);

            ret["Range"] = null;
            ret["Sheet"] = null;
            if (ret["Book"] != null)
            {
                (ret["Book"] as Excel.Workbook).Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(ret["Book"]);
            }
            ret["Book"] = null;
            if (ret["App"] != null)
            {
                (ret["App"] as Excel.Application).Quit();
                Marshal.ReleaseComObject(ret["App"]);
            }

            ret["App"] = null;
            ret = null;
        }

        public static User GetUserByID(string id)
        {
            #region oldcode
            //User ret = new User();
            //Excel.Application MyApp = new Excel.Application();
            //Excel.Workbook MyBook = MyApp.Workbooks.Open(userDB);
            //Excel.Worksheet MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            //Excel.Range xlRange = MySheet.UsedRange;
            //for (int i = 1; i <= xlRange.Rows.Count; i++)
            //{
            //    if (xlRange.Cells[i, 1].Value.ToString() == id)
            //    {
            //        ret.ID = xlRange.Cells[i, 1].Value.ToString();
            //        ret.name = xlRange.Cells[i, 2].Value.ToString();
            //        ret.lastN = xlRange.Cells[i, 3].Value.ToString();
            //        ret.password = xlRange.Cells[i, 4].Value.ToString();
            //        ret.isAdmin = Convert.ToBoolean(xlRange.Cells[i, 5].Value.ToString());

            //    }
            //}

            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(MySheet);

            //MyBook.Close();
            //Marshal.ReleaseComObject(MyBook);
            //MyApp.Quit();
            //Marshal.ReleaseComObject(MyApp);

            //return ret;
            #endregion

            foreach (User u in userList)
            {
                if (u.ID == id)
                    return u;
            }
            return new User();
        }

        public static Opp GetOpByID(string id)
        {

            foreach (Opp o in opportunites)
            {
                if (o.ID == id)
                    return o;
            }
            return new Opp();
        }

        public static List<Package> GetPackagesByID(string id)
        {
            List<Package> ret = new List<Package>();
            foreach (Package p in packages)
            {
                if (p.ID == id)
                    ret.Add(p);
            }
            return ret;
        }

        public static void RemovePackage(Package p)
        {
            Dictionary<string, object> ess = MakeExcelEss(opportunitesDB, 2);
            object[,] values = (ess["Range"] as Excel.Range).Value2;
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (values[i, 1].ToString() == p.ID && values[i, 2].ToString() == p.lineNum)
                {
                    Excel.Range range = (ess["Sheet"] as Excel.Worksheet).get_Range(string.Format("A{0}:A{1}", i, i), Type.Missing).EntireRow;
                    range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    Marshal.ReleaseComObject(range);
                    (ess["Book"] as Excel.Workbook).Save();
                    break;
                }
            }
            CleanExcelEss(ref ess);
        }

        public static void InsertNewUser(string id, string name, string lName, string password, bool isAdmin)
        {
            Dictionary<string, object> ess = MakeExcelEss(userDB);
            object[,] values = (ess["Range"] as Excel.Range).Value2;
            int r = values.GetLength(0) + 1;
            (ess["Sheet"] as Excel.Worksheet).Cells[r, 1] = id;
            (ess["Sheet"] as Excel.Worksheet).Cells[r, 2] = name;
            (ess["Sheet"] as Excel.Worksheet).Cells[r, 3] = lName;
            (ess["Sheet"] as Excel.Worksheet).Cells[r, 4] = password;
            (ess["Sheet"] as Excel.Worksheet).Cells[r, 5] = isAdmin;
            (ess["Book"] as Excel.Workbook).Save();
            CleanExcelEss(ref ess);
        }

        public static void InsertUpdateOpp(string id, string name, string lName, string phone, string email, DateTime treatedAt, string status, string treatedBy, string comment)
        {
            //must have for excel handeling
            Dictionary<string, object> ess = MakeExcelEss(opportunitesDB);
            object[,] values = (ess["Range"] as Excel.Range).Value2;

            //
            bool flag = true;
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (values[i, 1].ToString() == id)
                {
                    values[i, 2] = name;
                    values[i, 3] = lName;
                    values[i, 4] = phone;
                    values[i, 5] = email;
                    values[i, 7] = treatedBy;
                    values[i, 8] = treatedAt.Date;
                    values[i, 9] = comment;
                    flag = false;
                    (ess["Range"] as Excel.Range).Value2 = values;
                    break;
                }
            }
            if (flag)
            {
                int i = values.GetLength(0) + 1;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 1] = id;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 2] = name;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 3] = lName;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 4] = phone;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 5] = email;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 6] = status;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 7] = treatedBy;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 8] = treatedAt.Date;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 9] = comment;
            }
            (ess["Book"] as Excel.Workbook).Save();
            CleanExcelEss(ref ess);
        }

        public static void InsertUpdatePackage(string id, string phone , int type,bool close)
        {
            //must have for excel handeling
            Dictionary<string, object> ess = MakeExcelEss(opportunitesDB,2);
            object[,] values = (ess["Range"] as Excel.Range).Value2;

            //
            bool flag = true;
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (values[i, 1].ToString() == id&&values[1,2].ToString()==phone)
                {
                    values[i, 2] = phone;
                    values[i, 3] = type;
                    if(close)
                    {
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 1] = id;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 2] = phone;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 3] = type;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 4] = DateTime.Now.Date;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 5] = DateTime.Now.AddMonths(12).Date;
                    }
                    else (ess["Range"] as Excel.Range).Value2 = values;
                    flag = false;
                
                    break;
                }
            }
            if (flag)
            {
                int i = values.GetLength(0) + 1;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 1] = id;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 2] = phone;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 3] = type;
                if (close)
                {
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 4] = DateTime.Now.Date;
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 3] = DateTime.Now.AddMonths(12).Date;
                }

            }
           (ess["Book"] as Excel.Workbook).Save();
            CleanExcelEss(ref ess);
        }

        public static void MovetHistory(string id)
        {
            Opp o = GetOpByID(id);
            if (o.ID == null)
                return;
            Dictionary<string, object> ess = MakeExcelEss(historyDB);
            object[,] values = (ess["Range"] as Excel.Range).Value2;
            // we choose for pin a random number
            try
            {
                Random rnd = new Random();
                string pin = "";
                for (int j = 0; j < 10; j++)
                {
                    //if get 1 or 0, w will add for the pin number- a number
                    if (rnd.Next(2) == 0)
                        pin += rnd.Next(10).ToString();
                    else
                        // else we will add cher
                        pin += Convert.ToChar(rnd.Next(Convert.ToInt32('A'), Convert.ToInt32('Z') + 1)).ToString();
                }
                int i = values.GetLength(0);
                if (values[i, 1] != null) i++;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 1] = id;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 2] = o.name;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 3] = o.lastN;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 4] = o.phone;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 5] = o.email;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 6] = o.status;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 7] = o.treatedBy.ID;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 8] = o.treatedAt.Date;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 9] = o.comment;
                (ess["Sheet"] as Excel.Worksheet).Cells[i, 10] = pin;
                (ess["Book"] as Excel.Workbook).Save();
                CleanExcelEss(ref ess);
                values = null;

                ess = MakeExcelEss(historyDB, 2);
                values = (ess["Range"] as Excel.Range).Value2;
                ////

                i = values.GetLength(0);
                if (values[i, 1] != null) i++;
                foreach (Package p in packages)
                {
                    if (p.ID == id)
                    {
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 1] = id;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 2] = p.lineNum;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 3] = p.packageType;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 4] = p.startD.Date;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i, 5] = p.endD.Date;
                        (ess["Sheet"] as Excel.Worksheet).Cells[i++, 6] = pin;
                    }
                }
               (ess["Book"] as Excel.Workbook).Save();
                CleanExcelEss(ref ess);
                values = null;

                ess = MakeExcelEss(opportunitesDB);
                values = (ess["Range"] as Excel.Range).Value2;
                for (int j = 1; j <= values.GetLength(0); j++)
                {
                    if (values[j, 1].ToString() == id)
                    {
                        Excel.Range range = (ess["Sheet"] as Excel.Worksheet).get_Range(string.Format("A{0}:A{1}", j, j), Type.Missing).EntireRow;
                        range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        break;
                    }
                }
                (ess["Book"] as Excel.Workbook).Save();
                CleanExcelEss(ref ess);
                values = null;

                ess = MakeExcelEss(opportunitesDB, 2);
                values = (ess["Range"] as Excel.Range).Value2;
                for (int j = 1; j <= values.GetLength(0); j++)
                {
                    if (values[j, 1].ToString() == id)
                    {
                        Excel.Range range = (ess["Sheet"] as Excel.Worksheet).get_Range(string.Format("A{0}:A{1}", j, j), Type.Missing).EntireRow;
                        range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        values = (ess["Range"] as Excel.Range).Value2;
                        Marshal.ReleaseComObject(range);
                        j--;
                    }
                }

             (ess["Book"] as Excel.Workbook).Save();
                CleanExcelEss(ref ess);
                values = null;
            }
            catch
            {
                MessageBox.Show("Couldnt transfer " + id + " to history");
            }
        }

        public static void UpdateUserList()
        {
            if (userList != null)
                userList.Clear();
            else
                userList = new List<User>();
            Dictionary<string, object> ess = MakeExcelEss(userDB);
            try
            {
                object[,] values = (object[,])(ess["Range"] as Excel.Range).Value2;
                for (int i = 1; i <= values.GetLength(0); i++)
                {
                    User ret = new User();
                    ret.ID = values[i, 1].ToString();
                    ret.name = values[i, 2].ToString();
                    ret.lastN = values[i, 3].ToString();
                    ret.password = values[i, 4].ToString();
                    ret.isAdmin = Convert.ToBoolean(values[i, 5].ToString());
                    userList.Add(ret);
                }
            }
            catch { }
            finally
            {
                CleanExcelEss(ref ess);
            }
        }

        public static void UpdateOppList()
        {
            if (opportunites != null)
                opportunites.Clear();
            else
                opportunites = new List<Opp>();
            Dictionary<string, object> ess = MakeExcelEss(opportunitesDB);
            try
            {
                object[,] values = (object[,])(ess["Range"] as Excel.Range).Value2;
                for (int i = 1; i <= values.GetLength(0); i++)
                {
                    Opp ret = new Opp();
                    ret.ID = values[i, 1].ToString();
                    ret.name = values[i, 2].ToString();
                    ret.lastN = values[i, 3].ToString();
                    ret.phone = values[i, 4].ToString();
                    ret.email = values[i, 5].ToString();
                    ret.status = values[i, 6].ToString();
                    ret.treatedBy = GetUserByID(values[i, 7].ToString());
                    ret.treatedAt = DateTime.FromOADate((double)values[i, 8]);
                    ret.comment = values[i, 9].ToString();
                    opportunites.Add(ret);
                }
                CleanExcelEss(ref ess);

                ess = MakeExcelEss(historyDB);
                values = null;
                values = (object[,])(ess["Range"] as Excel.Range).Value2;
                for (int i = 1; i <= values.GetLength(0); i++)
                {
                    Opp ret = new Opp();
                    ret.ID = values[i, 1].ToString();
                    ret.name = values[i, 2].ToString();
                    ret.lastN = values[i, 3].ToString();
                    ret.phone = values[i, 4].ToString();
                    ret.email = values[i, 5].ToString();
                    ret.status = values[i, 6].ToString();
                    ret.treatedBy = GetUserByID(values[i, 7].ToString());
                    ret.treatedAt = DateTime.FromOADate((double)values[i, 8]);
                    ret.comment = values[i, 9].ToString();
                    ret.hID = values[i, 10].ToString();
                    opportunites.Add(ret);
                }
            }
            catch { }
            finally
            {
                CleanExcelEss(ref ess);
            }

        }

        public static void UpdatePacList()
        {
            if (packages != null)
                packages.Clear();
            else
                packages = new List<Package>();

            Dictionary<string, object> ess = MakeExcelEss(opportunitesDB, 2);
            try
            {
                object[,] values = (object[,])(ess["Range"] as Excel.Range).Value2;
                for (int i = 1; i <= values.GetLength(0); i++)
                {

                    Package ret = new Package();
                    ret.ID = values[i, 1].ToString();
                    ret.lineNum = values[i, 2].ToString();
                    ret.packageType = Convert.ToInt32(values[i, 3].ToString());
                    packages.Add(ret);
                }
                CleanExcelEss(ref ess);
                values = null;
                ess = MakeExcelEss(historyDB, 2);
                values = (object[,])(ess["Range"] as Excel.Range).Value2;
                for (int i = 1; i <= values.GetLength(0); i++)
                {
                    Package ret = new Package();
                    ret.ID = values[i, 1].ToString();
                    ret.lineNum = values[i, 2].ToString();
                    ret.packageType = Convert.ToInt32(values[i, 3].ToString());
                    ret.hID = values[i, 6].ToString();
                    packages.Add(ret);
                }
            }
            catch { }
            finally
            {
                CleanExcelEss(ref ess);
            }

        }

        private static void init()
        {
            UpdateUserList();
            UpdateOppList();
            UpdatePacList();
        }

        private static void GenerateRndPack2(string id, int count, DateTime treatedAt, Excel.Workbook MyBook)
        {
            //must have for excel handeling
            Excel.Worksheet MySheet = (Excel.Worksheet)MyBook.Sheets[2];
            Excel.Range xlRange = MySheet.UsedRange;
            object[,] values = (object[,])xlRange.Value2;
            int r = values.GetLength(0);
            if (values[r, 1] != null) r++;
            Random rnd = new Random();
            //
            try
            {
                for (int i = r; i < count + r; i++)
                {
                    MySheet.Cells[i, 1] = id;
                    string s = "05";
                    for (int j = 0; j < 8; j++)
                        s += rnd.Next(10);
                    MySheet.Cells[i, 2] = s;
                    MySheet.Cells[i, 3] = rnd.Next(1, 4).ToString();
                    DateTime d = treatedAt.Date.AddDays(-7 + rnd.Next(8));
                    MySheet.Cells[i, 4] = d.Date;
                    MySheet.Cells[i, 5] = d.AddMonths(rnd.Next(6, 13));
                }
                MyBook.Save();
            }
            catch { throw; }
            finally
            {
                //must have for excel handeling
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(MySheet);

            }
            //
        }

        private static void GenerateRndOpp(int k)
        {
            UpdateUserList();
            #region namesArrs
            string[] names = { "Aaran", "Aaren", "Aarez", "Aarman", "Aaron", "Aaron-James", "Aarron", "Aaryan", "Aaryn", "Aayan", "Aazaan", "Abaan", "Abbas", "Abdallah", "Abdalroof", "Abdihakim", "Abdirahman", "Abdisalam", "Abdul", "Abdul-Aziz", "Abdulbasir", "Abdulkadir", "Abdulkarem", "Abdulkhader", "Abdullah", "Abdul-Majeed", "Abdulmalik", "Abdul-Rehman", "Abdur", "Abdurraheem", "Abdur-Rahman", "Abdur-Rehmaan", "Abel", "Abhinav", "Abhisumant", "Abid", "Abir", "Abraham", "Abu", "Abubakar", "Ace", "Adain", "Adam", "Adam-James", "Addison", "Addisson", "Adegbola", "Adegbolahan", "Aden", "Adenn", "Adie", "Adil", "Aditya", "Adnan", "Adrian", "Adrien", "Aedan", "Aedin", "Aedyn", "Aeron", "Afonso", "Ahmad", "Ahmed", "Ahmed-Aziz", "Ahoua", "Ahtasham", "Aiadan", "Aidan", "Aiden", "Aiden-Jack", "Aiden-Vee", "Aidian", "Aidy", "Ailin", "Aiman", "Ainsley", "Ainslie", "Airen", "Airidas", "Airlie", "AJ", "Ajay", "A-Jay", "Ajayraj", "Akan", "Akram", "Al", "Ala", "Alan", "Alanas", "Alasdair", "Alastair", "Alber", "Albert", "Albie", "Aldred", "Alec", "Aled", "Aleem", "Aleksandar", "Aleksander", "Aleksandr", "Aleksandrs", "Alekzander", "Alessandro", "Alessio", "Alex", "Alexander", "Alexei", "Alexx", "Alexzander", "Alf", "Alfee", "Alfie", "Alfred", "Alfy", "Alhaji", "Al-Hassan", "Ali", "Aliekber", "Alieu", "Alihaider", "Alisdair", "Alishan", "Alistair", "Alistar", "Alister", "Aliyaan", "Allan", "Allan-Laiton", "Allen", "Allesandro", "Allister", "Ally", "Alphonse", "Altyiab", "Alum", "Alvern", "Alvin", "Alyas", "Amaan", "Aman", "Amani", "Ambanimoh", "Ameer", "Amgad", "Ami", "Amin", "Amir", "Ammaar", "Ammar", "Ammer", "Amolpreet", "Amos", "Amrinder", "Amrit", "Amro", "Anay", "Andrea", "Andreas", "Andrei", "Andrejs", "Andrew", "Andy", "Anees", "Anesu", "Angel", "Angelo", "Angus", "Anir", "Anis", "Anish", "Anmolpreet", "Annan", "Anndra", "Anselm", "Anthony", "Anthony-John", "Antoine", "Anton", "Antoni", "Antonio", "Antony", "Antonyo", "Anubhav", "Aodhan", "Aon", "Aonghus", "Apisai", "Arafat", "Aran", "Arandeep", "Arann", "Aray", "Arayan", "Archibald", "Archie", "Arda", "Ardal", "Ardeshir", "Areeb", "Areez", "Aref", "Arfin", "Argyle", "Argyll", "Ari", "Aria", "Arian", "Arihant", "Aristomenis", "Aristotelis", "Arjuna", "Arlo", "Armaan", "Arman", "Armen", "Arnab", "Arnav", "Arnold", "Aron", "Aronas", "Arran", "Arrham", "Arron", "Arryn", "Arsalan", "Artem", "Arthur", "Artur", "Arturo", "Arun", "Arunas", "Arved", "Arya", "Aryan", "Aryankhan", "Aryian", "Aryn", "Asa", "Asfhan", "Ash", "Ashlee-jay", "Ashley", "Ashton", "Ashton-Lloyd", "Ashtyn", "Ashwin", "Asif", "Asim", "Aslam", "Asrar", "Ata", "Atal", "Atapattu", "Ateeq", "Athol", "Athon", "Athos-Carlos", "Atli", "Atom", "Attila", "Aulay", "Aun", "Austen", "Austin", "Avani", "Averon", "Avi", "Avinash", "Avraham", "Awais", "Awwal", "Axel", "Ayaan", "Ayan", "Aydan", "Ayden", "Aydin", "Aydon", "Ayman", "Ayomide", "Ayren", "Ayrton", "Aytug", "Ayub", "Ayyub", "Azaan", "Azedine", "Azeem", "Azim", "Aziz", "Azlan", "Azzam", "Azzedine", "Babatunmise", "Babur", "Bader", "Badr", "Badsha", "Bailee", "Bailey", "Bailie", "Bailley", "Baillie", "Baley", "Balian", "Banan", "Barath", "Barkley", "Barney", "Baron", "Barrie", "Barry", "Bartlomiej", "Bartosz", "Basher", "Basile", "Baxter", "Baye", "Bayley", "Beau", "Beinn", "Bekim", "Believe", "Ben", "Bendeguz", "Benedict", "Benjamin", "Benjamyn", "Benji", "Benn", "Bennett", "Benny", "Benoit", "Bentley", "Berkay", "Bernard", "Bertie", "Bevin", "Bezalel", "Bhaaldeen", "Bharath", "Bilal", "Bill", "Billy", "Binod", "Bjorn", "Blaike", "Blaine", "Blair", "Blaire", "Blake", "Blazej", "Blazey", "Blessing", "Blue", "Blyth", "Bo", "Boab", "Bob", "Bobby", "Bobby-Lee", "Bodhan", "Boedyn", "Bogdan", "Bohbi", "Bony", "Bowen", "Bowie", "Boyd", "Bracken", "Brad", "Bradan", "Braden", "Bradley", "Bradlie", "Bradly", "Brady", "Bradyn", "Braeden", "Braiden", "Brajan", "Brandan", "Branden", "Brandon", "Brandonlee", "Brandon-Lee", "Brandyn", "Brannan", "Brayden", "Braydon", "Braydyn", "Breandan", "Brehme", "Brendan", "Brendon", "Brendyn", "Breogan", "Bret", "Brett", "Briaddon", "Brian", "Brodi", "Brodie", "Brody", "Brogan", "Broghan", "Brooke", "Brooklin", "Brooklyn", "Bruce", "Bruin", "Bruno", "Brunon", "Bryan", "Bryce", "Bryden", "Brydon", "Brydon-Craig", "Bryn", "Brynmor", "Bryson", "Buddy", "Bully", "Burak", "Burhan", "Butali", "Butchi", "Byron", "Cabhan", "Cadan", "Cade", "Caden", "Cadon", "Cadyn", "Caedan", "Caedyn", "Cael", "Caelan", "Caelen", "Caethan", "Cahl", "Cahlum", "Cai", "Caidan", "Caiden", "Caiden-Paul", "Caidyn", "Caie", "Cailaen", "Cailean", "Caileb-John", "Cailin", "Cain", "Caine", "Cairn", "Cal", "Calan", "Calder", "Cale", "Calean", "Caleb", "Calen", "Caley", "Calib", "Calin", "Callahan", "Callan", "Callan-Adam", "Calley", "Callie", "Callin", "Callum", "Callun", "Callyn", "Calum", "Calum-James", "Calvin", "Cambell", "Camerin", "Cameron", "Campbel", "Campbell", "Camron", "Caolain", "Caolan", "Carl", "Carlo", "Carlos", "Carrich", "Carrick", "Carson", "Carter", "Carwyn", "Casey", "Casper", "Cassy", "Cathal", "Cator", "Cavan", "Cayden", "Cayden-Robert", "Cayden-Tiamo", "Ceejay", "Ceilan", "Ceiran", "Ceirin", "Ceiron", "Cejay", "Celik", "Cephas", "Cesar", "Cesare", "Chad", "Chaitanya", "Chang-Ha", "Charles", "Charley", "Charlie", "Charly", "Chase", "Che", "Chester", "Chevy", "Chi", "Chibudom", "Chidera", "Chimsom", "Chin", "Chintu", "Chiqal", "Chiron", "Chris", "Chris-Daniel", "Chrismedi", "Christian", "Christie", "Christoph", "Christopher", "Christopher-Lee", "Christy", "Chu", "Chukwuemeka", "Cian", "Ciann", "Ciar", "Ciaran", "Ciarian", "Cieran", "Cillian", "Cillin", "Cinar", "CJ", "C-Jay", "Clark", "Clarke", "Clayton", "Clement", "Clifford", "Clyde", "Cobain", "Coban", "Coben", "Cobi", "Cobie", "Coby", "Codey", "Codi", "Codie", "Cody", "Cody-Lee", "Coel", "Cohan", "Cohen", "Colby", "Cole", "Colin", "Coll", "Colm", "Colt", "Colton", "Colum", "Colvin", "Comghan", "Conal", "Conall", "Conan", "Conar", "Conghaile", "Conlan", "Conley", "Conli", "Conlin", "Conlly", "Conlon", "Conlyn", "Connal", "Connall", "Connan", "Connar", "Connel", "Connell", "Conner", "Connolly", "Connor", "Connor-David", "Conor", "Conrad", "Cooper", "Copeland", "Coray", "Corben", "Corbin", "Corey", "Corey-James", "Corey-Jay", "Cori", "Corie", "Corin", "Cormac", "Cormack", "Cormak", "Corran", "Corrie", "Cory", "Cosmo", "Coupar", "Craig", "Craig-James", "Crawford", "Creag", "Crispin", "Cristian", "Crombie", "Cruiz", "Cruz", "Cuillin", "Cullen", "Cullin", "Curtis", "Cyrus", "Daanyaal", "Daegan", "Daegyu", "Dafydd", "Dagon", "Dailey", "Daimhin", "Daithi", "Dakota", "Daksh", "Dale", "Dalong", "Dalton", "Damian", "Damien", "Damon", "Dan", "Danar", "Dane", "Danial", "Daniel", "Daniele", "Daniel-James", "Daniels", "Daniil", "Danish", "Daniyal", "Danniel", "Danny", "Dante", "Danyal", "Danyil", "Danys", "Daood", "Dara", "Darach", "Daragh", "Darcy", "D'arcy", "Dareh", "Daren", "Darien", "Darius", "Darl", "Darn", "Darrach", "Darragh", "Darrel", "Darrell", "Darren", "Darrie", "Darrius", "Darroch", "Darryl", "Darryn", "Darwyn", "Daryl", "Daryn", "Daud", "Daumantas", "Davi", "David", "David-Jay", "David-Lee", "Davie", "Davis", "Davy", "Dawid", "Dawson", "Dawud", "Dayem", "Daymian", "Deacon", "Deagan", "Dean", "Deano", "Decklan", "Declain", "Declan", "Declyan", "Declyn", "Dedeniseoluwa", "Deecan", "Deegan", "Deelan", "Deklain-Jaimes", "Del", "Demetrius", "Denis", "Deniss", "Dennan", "Dennin", "Dennis", "Denny", "Dennys", "Denon", "Denton", "Denver", "Denzel", "Deon", "Derek", "Derick", "Derin", "Dermot", "Derren", "Derrie", "Derrin", "Derron", "Derry", "Derryn", "Deryn", "Deshawn", "Desmond", "Dev", "Devan", "Devin", "Devlin", "Devlyn", "Devon", "Devrin", "Devyn", "Dex", "Dexter", "Dhani", "Dharam", "Dhavid", "Dhyia", "Diarmaid", "Diarmid", "Diarmuid", "Didier", "Diego", "Diesel", "Diesil", "Digby", "Dilan", "Dilano", "Dillan", "Dillon", "Dilraj", "Dimitri", "Dinaras", "Dion", "Dissanayake", "Dmitri", "Doire", "Dolan", "Domanic", "Domenico", "Domhnall", "Dominic", "Dominick", "Dominik", "Donald", "Donnacha", "Donnie", "Dorian", "Dougal", "Douglas", "Dougray", "Drakeo", "Dre", "Dregan", "Drew", "Dugald", "Duncan", "Duriel", "Dustin", "Dylan", "Dylan-Jack", "Dylan-James", "Dylan-John", "Dylan-Patrick", "Dylin", "Dyllan", "Dyllan-James", "Dyllon", "Eadie", "Eagann", "Eamon", "Eamonn", "Eason", "Eassan", "Easton", "Ebow", "Ed", "Eddie", "Eden", "Ediomi", "Edison", "Eduardo", "Eduards", "Edward", "Edwin", "Edwyn", "Eesa", "Efan", "Efe", "Ege", "Ehsan", "Ehsen", "Eiddon", "Eidhan", "Eihli", "Eimantas", "Eisa", "Eli", "Elias", "Elijah", "Eliot", "Elisau", "Eljay", "Eljon", "Elliot", "Elliott", "Ellis", "Ellisandro", "Elshan", "Elvin", "Elyan", "Emanuel", "Emerson", "Emil", "Emile", "Emir", "Emlyn", "Emmanuel", "Emmet", "Eng", "Eniola", "Enis", "Ennis", "Enrico", "Enrique", "Enzo", "Eoghain", "Eoghan", "Eoin", "Eonan", "Erdehan", "Eren", "Erencem", "Eric", "Ericlee", "Erik", "Eriz", "Ernie-Jacks", "Eroni", "Eryk", "Eshan", "Essa", "Esteban", "Ethan", "Etienne", "Etinosa", "Euan", "Eugene", "Evan", "Evann", "Ewan", "Ewen", "Ewing", "Exodi", "Ezekiel", "Ezra", "Fabian", "Fahad", "Faheem", "Faisal", "Faizaan", "Famara", "Fares", "Farhaan", "Farhan", "Farren", "Farzad", "Fauzaan", "Favour", "Fawaz", "Fawkes", "Faysal", "Fearghus", "Feden", "Felix", "Fergal", "Fergie", "Fergus", "Ferre", "Fezaan", "Fiachra", "Fikret", "Filip", "Filippo", "Finan", "Findlay", "Findlay-James", "Findlie", "Finlay", "Finley", "Finn", "Finnan", "Finnean", "Finnen", "Finnlay", "Finnley", "Fintan", "Fionn", "Firaaz", "Fletcher", "Flint", "Florin", "Flyn", "Flynn", "Fodeba", "Folarinwa", "Forbes", "Forgan", "Forrest", "Fox", "Francesco", "Francis", "Francisco", "Franciszek", "Franco", "Frank", "Frankie", "Franklin", "Franko", "Fraser", "Frazer", "Fred", "Freddie", "Frederick", "Fruin", "Fyfe", "Fyn", "Fynlay", "Fynn", "Gabriel", "Gallagher", "Gareth", "Garren", "Garrett", "Garry", "Gary", "Gavin", "Gavin-Lee", "Gene", "Geoff", "Geoffrey", "Geomer", "Geordan", "Geordie", "George", "Georgia", "Georgy", "Gerard", "Ghyll", "Giacomo", "Gian", "Giancarlo", "Gianluca", "Gianmarco", "Gideon", "Gil", "Gio", "Girijan", "Girius", "Gjan", "Glascott", "Glen", "Glenn", "Gordon", "Grady", "Graeme", "Graham", "Grahame", "Grant", "Grayson", "Greg", "Gregor", "Gregory", "Greig", "Griffin", "Griffyn", "Grzegorz", "Guang", "Guerin", "Guillaume", "Gurardass", "Gurdeep", "Gursees", "Gurthar", "Gurveer", "Gurwinder", "Gus", "Gustav", "Guthrie", "Guy", "Gytis", "Habeeb", "Hadji", "Hadyn", "Hagun", "Haiden", "Haider", "Hamad", "Hamid", "Hamish", "Hamza", "Hamzah", "Han", "Hansen", "Hao", "Hareem", "Hari", "Harikrishna", "Haris", "Harish", "Harjeevan", "Harjyot", "Harlee", "Harleigh", "Harley", "Harman", "Harnek", "Harold", "Haroon", "Harper", "Harri", "Harrington", "Harris", "Harrison", "Harry", "Harvey", "Harvie", "Harvinder", "Hasan", "Haseeb", "Hashem", "Hashim", "Hassan", "Hassanali", "Hately", "Havila", "Hayden", "Haydn", "Haydon", "Haydyn", "Hcen", "Hector", "Heddle", "Heidar", "Heini", "Hendri", "Henri", "Henry", "Herbert", "Heyden", "Hiro", "Hirvaansh", "Hishaam", "Hogan", "Honey", "Hong", "Hope", "Hopkin", "Hosea", "Howard", "Howie", "Hristomir", "Hubert", "Hugh", "Hugo", "Humza", "Hunter", "Husnain", "Hussain", "Hussan", "Hussnain", "Hussnan", "Hyden", "I", "Iagan", "Iain", "Ian", "Ibraheem", "Ibrahim", "Idahosa", "Idrees", "Idris", "Iestyn", "Ieuan", "Igor", "Ihtisham", "Ijay", "Ikechukwu", "Ikemsinachukwu", "Ilyaas", "Ilyas", "Iman", "Immanuel", "Inan", "Indy", "Ines", "Innes", "Ioannis", "Ireayomide", "Ireoluwa", "Irvin", "Irvine", "Isa", "Isaa", "Isaac", "Isaiah", "Isak", "Isher", "Ishwar", "Isimeli", "Isira", "Ismaeel", "Ismail", "Israel", "Issiaka", "Ivan", "Ivar", "Izaak", "J", "Jaay", "Jac", "Jace", "Jack", "Jacki", "Jackie", "Jack-James", "Jackson", "Jacky", "Jacob", "Jacques", "Jad", "Jaden", "Jadon", "Jadyn", "Jae", "Jagat", "Jago", "Jaheim", "Jahid", "Jahy", "Jai", "Jaida", "Jaiden", "Jaidyn", "Jaii", "Jaime", "Jai-Rajaram", "Jaise", "Jak", "Jake", "Jakey", "Jakob", "Jaksyn", "Jakub", "Jamaal", "Jamal", "Jameel", "Jameil", "James", "James-Paul", "Jamey", "Jamie", "Jan", "Jaosha", "Jardine", "Jared", "Jarell", "Jarl", "Jarno", "Jarred", "Jarvi", "Jasey-Jay", "Jasim", "Jaskaran", "Jason", "Jasper", "Jaxon", "Jaxson", "Jay", "Jaydan", "Jayden", "Jayden-James", "Jayden-Lee", "Jayden-Paul", "Jayden-Thomas", "Jaydn", "Jaydon", "Jaydyn", "Jayhan", "Jay-Jay", "Jayke", "Jaymie", "Jayse", "Jayson", "Jaz", "Jazeb", "Jazib", "Jazz", "Jean", "Jean-Lewis", "Jean-Pierre", "Jebadiah", "Jed", "Jedd", "Jedidiah", "Jeemie", "Jeevan", "Jeffrey", "Jensen", "Jenson", "Jensyn", "Jeremy", "Jerome", "Jeronimo", "Jerrick", "Jerry", "Jesse", "Jesuseun", "Jeswin", "Jevan", "Jeyun", "Jez", "Jia", "Jian", "Jiao", "Jimmy", "Jincheng", "JJ", "Joaquin", "Joash", "Jock", "Jody", "Joe", "Joeddy", "Joel", "Joey", "Joey-Jack", "Johann", "Johannes", "Johansson", "John", "Johnathan", "Johndean", "Johnjay", "John-Michael", "Johnnie", "Johnny", "Johnpaul", "John-Paul", "John-Scott", "Johnson", "Jole", "Jomuel", "Jon", "Jonah", "Jonatan", "Jonathan", "Jonathon", "Jonny", "Jonothan", "Jon-Paul", "Jonson", "Joojo", "Jordan", "Jordi", "Jordon", "Jordy", "Jordyn", "Jorge", "Joris", "Jorryn", "Josan", "Josef", "Joseph", "Josese", "Josh", "Joshiah", "Joshua", "Josiah", "Joss", "Jostelle", "Joynul", "Juan", "Jubin", "Judah", "Jude", "Jules", "Julian", "Julien", "Jun", "Junior", "Jura", "Justan", "Justin", "Justinas", "Kaan", "Kabeer", "Kabir", "Kacey", "Kacper", "Kade", "Kaden", "Kadin", "Kadyn", "Kaeden", "Kael", "Kaelan", "Kaelin", "Kaelum", "Kai", "Kaid", "Kaidan", "Kaiden", "Kaidinn", "Kaidyn", "Kaileb", "Kailin", "Kain", "Kaine", "Kainin", "Kainui", "Kairn", "Kaison", "Kaiwen", "Kajally", "Kajetan", "Kalani", "Kale", "Kaleb", "Kaleem", "Kal-el", "Kalen", "Kalin", "Kallan", "Kallin", "Kalum", "Kalvin", "Kalvyn", "Kameron", "Kames", "Kamil", "Kamran", "Kamron", "Kane", "Karam", "Karamvir", "Karandeep", "Kareem", "Karim", "Karimas", "Karl", "Karol", "Karson", "Karsyn", "Karthikeya", "Kasey", "Kash", "Kashif", "Kasim", "Kasper", "Kasra", "Kavin", "Kayam", "Kaydan", "Kayden", "Kaydin", "Kaydn", "Kaydyn", "Kaydyne", "Kayleb", "Kaylem", "Kaylum", "Kayne", "Kaywan", "Kealan", "Kealon", "Kean", "Keane", "Kearney", "Keatin", "Keaton", "Keavan", "Keayn", "Kedrick", "Keegan", "Keelan", "Keelin", "Keeman", "Keenan", "Keenan-Lee", "Keeton", "Kehinde", "Keigan", "Keilan", "Keir", "Keiran", "Keiren", "Keiron", "Keiryn", "Keison", "Keith", "Keivlin", "Kelam", "Kelan", "Kellan", "Kellen", "Kelso", "Kelum", "Kelvan", "Kelvin", "Ken", "Kenan", "Kendall", "Kendyn", "Kenlin", "Kenneth", "Kensey", "Kenton", "Kenyon", "Kenzeigh", "Kenzi", "Kenzie", "Kenzo", "Kenzy", "Keo", "Ker", "Kern", "Kerr", "Kevan", "Kevin", "Kevyn", "Kez", "Khai", "Khalan", "Khaleel", "Khaya", "Khevien", "Khizar", "Khizer", "Kia", "Kian", "Kian-James", "Kiaran", "Kiarash", "Kie", "Kiefer", "Kiegan", "Kienan", "Kier", "Kieran", "Kieran-Scott", "Kieren", "Kierin", "Kiern", "Kieron", "Kieryn", "Kile", "Killian", "Kimi", "Kingston", "Kinneil", "Kinnon", "Kinsey", "Kiran", "Kirk", "Kirwin", "Kit", "Kiya", "Kiyonari", "Kjae", "Klein", "Klevis", "Kobe", "Kobi", "Koby", "Koddi", "Koden", "Kodi", "Kodie", "Kody", "Kofi", "Kogan", "Kohen", "Kole", "Konan", "Konar", "Konnor", "Konrad", "Koray", "Korben", "Korbyn", "Korey", "Kori", "Korrin", "Kory", "Koushik", "Kris", "Krish", "Krishan", "Kriss", "Kristian", "Kristin", "Kristofer", "Kristoffer", "Kristopher", "Kruz", "Krzysiek", "Krzysztof", "Ksawery", "Ksawier", "Kuba", "Kurt", "Kurtis", "Kurtis-Jae", "Kyaan", "Kyan", "Kyde", "Kyden", "Kye", "Kyel", "Kyhran", "Kyie", "Kylan", "Kylar", "Kyle", "Kyle-Derek", "Kylian", "Kym", "Kynan", "Kyral", "Kyran", "Kyren", "Kyrillos", "Kyro", "Kyron", "Kyrran", "Lachlainn", "Lachlan", "Lachlann", "Lael", "Lagan", "Laird", "Laison", "Lakshya", "Lance", "Lancelot", "Landon", "Lang", "Lasse", "Latif", "Lauchlan", "Lauchlin", "Laughlan", "Lauren", "Laurence", "Laurie", "Lawlyn", "Lawrence", "Lawrie", "Lawson", "Layne", "Layton", "Lee", "Leigh", "Leigham", "Leighton", "Leilan", "Leiten", "Leithen", "Leland", "Lenin", "Lennan", "Lennen", "Lennex", "Lennon", "Lennox", "Lenny", "Leno", "Lenon", "Lenyn", "Leo", "Leon", "Leonard", "Leonardas", "Leonardo", "Lepeng", "Leroy", "Leven", "Levi", "Levon", "Levy", "Lewie", "Lewin", "Lewis", "Lex", "Leydon", "Leyland", "Leylann", "Leyton", "Liall", "Liam", "Liam-Stephen", "Limo", "Lincoln", "Lincoln-John", "Lincon", "Linden", "Linton", "Lionel", "Lisandro", "Litrell", "Liyonela-Elam", "LLeyton", "Lliam", "Lloyd", "Lloyde", "Loche", "Lochlan", "Lochlann", "Lochlan-Oliver", "Lock", "Lockey", "Logan", "Logann", "Logan-Rhys", "Loghan", "Lokesh", "Loki", "Lomond", "Lorcan", "Lorenz", "Lorenzo", "Lorne", "Loudon", "Loui", "Louie", "Louis", "Loukas", "Lovell", "Luc", "Luca", "Lucais", "Lucas", "Lucca", "Lucian", "Luciano", "Lucien", "Lucus", "Luic", "Luis", "Luk", "Luka", "Lukas", "Lukasz", "Luke", "Lukmaan", "Luqman", "Lyall", "Lyle", "Lyndsay", "Lysander", "Maanav", "Maaz", "Mac", "Macallum", "Macaulay", "Macauley", "Macaully", "Machlan", "Maciej", "Mack", "Mackenzie", "Mackenzy", "Mackie", "Macsen", "Macy", "Madaki", "Maddison", "Maddox", "Madison", "Madison-Jake", "Madox", "Mael", "Magnus", "Mahan", "Mahdi", "Mahmoud", "Maias", "Maison", "Maisum", "Maitlind", "Majid", "Makensie", "Makenzie", "Makin", "Maksim", "Maksymilian", "Malachai", "Malachi", "Malachy", "Malakai", "Malakhy", "Malcolm", "Malik", "Malikye", "Malo", "Ma'moon", "Manas", "Maneet", "Manmohan", "Manolo", "Manson", "Mantej", "Manuel", "Manus", "Marc", "Marc-Anthony", "Marcel", "Marcello", "Marcin", "Marco", "Marcos", "Marcous", "Marcquis", "Marcus", "Mario", "Marios", "Marius", "Mark", "Marko", "Markus", "Marley", "Marlin", "Marlon", "Maros", "Marshall", "Martin", "Marty", "Martyn", "Marvellous", "Marvin", "Marwan", "Maryk", "Marzuq", "Mashhood", "Mason", "Mason-Jay", "Masood", "Masson", "Matas", "Matej", "Mateusz", "Mathew", "Mathias", "Mathu", "Mathuyan", "Mati", "Matt", "Matteo", "Matthew", "Matthew-William", "Matthias", "Max", "Maxim", "Maximilian", "Maximillian", "Maximus", "Maxwell", "Maxx", "Mayeul", "Mayson", "Mazin", "Mcbride", "McCaulley", "McKade", "McKauley", "McKay", "McKenzie", "McLay", "Meftah", "Mehmet", "Mehraz", "Meko", "Melville", "Meshach", "Meyzhward", "Micah", "Michael", "Michael-Alexander", "Michael-James", "Michal", "Michat", "Micheal", "Michee", "Mickey", "Miguel", "Mika", "Mikael", "Mikee", "Mikey", "Mikhail", "Mikolaj", "Miles", "Millar", "Miller", "Milo", "Milos", "Milosz", "Mir", "Mirza", "Mitch", "Mitchel", "Mitchell", "Moad", "Moayd", "Mobeen", "Modoulamin", "Modu", "Mohamad", "Mohamed", "Mohammad", "Mohammad-Bilal", "Mohammed", "Mohanad", "Mohd", "Momin", "Momooreoluwa", "Montague", "Montgomery", "Monty", "Moore", "Moosa", "Moray", "Morgan", "Morgyn", "Morris", "Morton", "Moshy", "Motade", "Moyes", "Msughter", "Mueez", "Muhamadjavad", "Muhammad", "Muhammed", "Muhsin", "Muir", "Munachi", "Muneeb", "Mungo", "Munir", "Munmair", "Munro", "Murdo", "Murray", "Murrough", "Murry", "Musa", "Musse", "Mustafa", "Mustapha", "Muzammil", "Muzzammil", "Mykie", "Myles", "Mylo", "Nabeel", "Nadeem", "Nader", "Nagib", "Naif", "Nairn", "Narvic", "Nash", "Nasser", "Nassir", "Natan", "Nate", "Nathan", "Nathanael", "Nathanial", "Nathaniel", "Nathan-Rae", "Nawfal", "Nayan", "Neco", "Neil", "Nelson", "Neo", "Neshawn", "Nevan", "Nevin", "Ngonidzashe", "Nial", "Niall", "Nicholas", "Nick", "Nickhill", "Nicki", "Nickson", "Nicky", "Nico", "Nicodemus", "Nicol", "Nicolae", "Nicolas", "Nidhish", "Nihaal", "Nihal", "Nikash", "Nikhil", "Niki", "Nikita", "Nikodem", "Nikolai", "Nikos", "Nilav", "Niraj", "Niro", "Niven", "Noah", "Noel", "Nolan", "Noor", "Norman", "Norrie", "Nuada", "Nyah", "Oakley", "Oban", "Obieluem", "Obosa", "Odhran", "Odin", "Odynn", "Ogheneochuko", "Ogheneruno", "Ohran", "Oilibhear", "Oisin", "Ojima-Ojo", "Okeoghene", "Olaf", "Ola-Oluwa", "Olaoluwapolorimi", "Ole", "Olie", "Oliver", "Olivier", "Oliwier", "Ollie", "Olurotimi", "Oluwadamilare", "Oluwadamiloju", "Oluwafemi", "Oluwafikunayomi", "Oluwalayomi", "Oluwatobiloba", "Oluwatoni", "Omar", "Omri", "Oran", "Orin", "Orlando", "Orley", "Orran", "Orrick", "Orrin", "Orson", "Oryn", "Oscar", "Osesenagha", "Oskar", "Ossian", "Oswald", "Otto", "Owain", "Owais", "Owen", "Owyn", "Oz", "Ozzy", "Pablo", "Pacey", "Padraig", "Paolo", "Pardeepraj", "Parkash", "Parker", "Pascoe", "Pasquale", "Patrick", "Patrick-John", "Patrikas", "Patryk", "Paul", "Pavit", "Pawel", "Pawlo", "Pearce", "Pearse", "Pearsen", "Pedram", "Pedro", "Peirce", "Peiyan", "Pele", "Peni", "Peregrine", "Peter", "Phani", "Philip", "Philippos", "Phinehas", "Phoenix", "Phoevos", "Pierce", "Pierre-Antoine", "Pieter", "Pietro", "Piotr", "Porter", "Prabhjoit", "Prabodhan", "Praise", "Pranav", "Pravin", "Precious", "Prentice", "Presley", "Preston", "Preston-Jay", "Prinay", "Prince", "Prithvi", "Promise", "Puneetpaul", "Pushkar", "Qasim", "Qirui", "Quinlan", "Quinn", "Radmiras", "Raees", "Raegan", "Rafael", "Rafal", "Rafferty", "Rafi", "Raheem", "Rahil", "Rahim", "Rahman", "Raith", "Raithin", "Raja", "Rajab-Ali", "Rajan", "Ralfs", "Ralph", "Ramanas", "Ramit", "Ramone", "Ramsay", "Ramsey", "Rana", "Ranolph", "Raphael", "Rasmus", "Rasul", "Raul", "Raunaq", "Ravin", "Ray", "Rayaan", "Rayan", "Rayane", "Rayden", "Rayhan", "Raymond", "Rayne", "Rayyan", "Raza", "Reace", "Reagan", "Reean", "Reece", "Reed", "Reegan", "Rees", "Reese", "Reeve", "Regan", "Regean", "Reggie", "Rehaan", "Rehan", "Reice", "Reid", "Reigan", "Reilly", "Reily", "Reis", "Reiss", "Remigiusz", "Remo", "Remy", "Ren", "Renars", "Reng", "Rennie", "Reno", "Reo", "Reuben", "Rexford", "Reynold", "Rhein", "Rheo", "Rhett", "Rheyden", "Rhian", "Rhoan", "Rholmark", "Rhoridh", "Rhuairidh", "Rhuan", "Rhuaridh", "Rhudi", "Rhy", "Rhyan", "Rhyley", "Rhyon", "Rhys", "Rhys-Bernard", "Rhyse", "Riach", "Rian", "Ricards", "Riccardo", "Ricco", "Rice", "Richard", "Richey", "Richie", "Ricky", "Rico", "Ridley", "Ridwan", "Rihab", "Rihan", "Rihards", "Rihonn", "Rikki", "Riley", "Rio", "Rioden", "Rishi", "Ritchie", "Rivan", "Riyadh", "Riyaj", "Roan", "Roark", "Roary", "Rob", "Robbi", "Robbie", "Robbie-lee", "Robby", "Robert", "Robert-Gordon", "Robertjohn", "Robi", "Robin", "Rocco", "Roddy", "Roderick", "Rodrigo", "Roen", "Rogan", "Roger", "Rohaan", "Rohan", "Rohin", "Rohit", "Rokas", "Roman", "Ronald", "Ronan", "Ronan-Benedict", "Ronin", "Ronnie", "Rooke", "Roray", "Rori", "Rorie", "Rory", "Roshan", "Ross", "Ross-Andrew", "Rossi", "Rowan", "Rowen", "Roy", "Ruadhan", "Ruaidhri", "Ruairi", "Ruairidh", "Ruan", "Ruaraidh", "Ruari", "Ruaridh", "Ruben", "Rubhan", "Rubin", "Rubyn", "Rudi", "Rudy", "Rufus", "Rui", "Ruo", "Rupert", "Ruslan", "Russel", "Russell", "Ryaan", "Ryan", "Ryan-Lee", "Ryden", "Ryder", "Ryese", "Ryhs", "Rylan", "Rylay", "Rylee", "Ryleigh", "Ryley", "Rylie", "Ryo", "Ryszard", "Saad", "Sabeen", "Sachkirat", "Saffi", "Saghun", "Sahaib", "Sahbian", "Sahil", "Saif", "Saifaddine", "Saim", "Sajid", "Sajjad", "Salahudin", "Salman", "Salter", "Salvador", "Sam", "Saman", "Samar", "Samarjit", "Samatar", "Sambrid", "Sameer", "Sami", "Samir", "Sami-Ullah", "Samual", "Samuel", "Samuela", "Samy", "Sanaullah", "Sandro", "Sandy", "Sanfur", "Sanjay", "Santiago", "Santino", "Satveer", "Saul", "Saunders", "Savin", "Sayad", "Sayeed", "Sayf", "Scot", "Scott", "Scott-Alexander", "Seaan", "Seamas", "Seamus", "Sean", "Seane", "Sean-James", "Sean-Paul", "Sean-Ray", "Seb", "Sebastian", "Sebastien", "Selasi", "Seonaidh", "Sephiroth", "Sergei", "Sergio", "Seth", "Sethu", "Seumas", "Shaarvin", "Shadow", "Shae", "Shahmir", "Shai", "Shane", "Shannon", "Sharland", "Sharoz", "Shaughn", "Shaun", "Shaunpaul", "Shaun-Paul", "Shaun-Thomas", "Shaurya", "Shaw", "Shawn", "Shawnpaul", "Shay", "Shayaan", "Shayan", "Shaye", "Shayne", "Shazil", "Shea", "Sheafan", "Sheigh", "Shenuk", "Sher", "Shergo", "Sheriff", "Sherwyn", "Shiloh", "Shiraz", "Shreeram", "Shreyas", "Shyam", "Siddhant", "Siddharth", "Sidharth", "Sidney", "Siergiej", "Silas", "Simon", "Sinai", "Skye", "Sofian", "Sohaib", "Sohail", "Soham", "Sohan", "Sol", "Solomon", "Sonneey", "Sonni", "Sonny", "Sorley", "Soul", "Spencer", "Spondon", "Stanislaw", "Stanley", "Stefan", "Stefano", "Stefin", "Stephen", "Stephenjunior", "Steve", "Steven", "Steven-lee", "Stevie", "Stewart", "Stewarty", "Strachan", "Struan", "Stuart", "Su", "Subhaan", "Sudais", "Suheyb", "Suilven", "Sukhi", "Sukhpal", "Sukhvir", "Sulayman", "Sullivan", "Sultan", "Sung", "Sunny", "Suraj", "Surien", "Sweyn", "Syed", "Sylvain", "Symon", "Szymon", "Tadd", "Taddy", "Tadhg", "Taegan", "Taegen", "Tai", "Tait", "Taiwo", "Talha", "Taliesin", "Talon", "Talorcan", "Tamar", "Tamiem", "Tammam", "Tanay", "Tane", "Tanner", "Tanvir", "Tanzeel", "Taonga", "Tarik", "Tariq-Jay", "Tate", "Taylan", "Taylar", "Tayler", "Taylor", "Taylor-Jay", "Taylor-Lee", "Tayo", "Tayyab", "Tayye", "Tayyib", "Teagan", "Tee", "Teejay", "Tee-jay", "Tegan", "Teighen", "Teiyib", "Te-Jay", "Temba", "Teo", "Teodor", "Teos", "Terry", "Teydren", "Theo", "Theodore", "Thiago", "Thierry", "Thom", "Thomas", "Thomas-Jay", "Thomson", "Thorben", "Thorfinn", "Thrinei", "Thumbiko", "Tiago", "Tian", "Tiarnan", "Tibet", "Tieran", "Tiernan", "Timothy", "Timucin", "Tiree", "Tisloh", "Titi", "Titus", "Tiylar", "TJ", "Tjay", "T-Jay", "Tobey", "Tobi", "Tobias", "Tobie", "Toby", "Todd", "Tokinaga", "Toluwalase", "Tom", "Tomas", "Tomasz", "Tommi-Lee", "Tommy", "Tomson", "Tony", "Torin", "Torquil", "Torran", "Torrin", "Torsten", "Trafford", "Trai", "Travis", "Tre", "Trent", "Trey", "Tristain", "Tristan", "Troy", "Tubagus", "Turki", "Turner", "Ty", "Ty-Alexander", "Tye", "Tyelor", "Tylar", "Tyler", "Tyler-James", "Tyler-Jay", "Tyllor", "Tylor", "Tymom", "Tymon", "Tymoteusz", "Tyra", "Tyree", "Tyrnan", "Tyrone", "Tyson", "Ubaid", "Ubayd", "Uchenna", "Uilleam", "Umair", "Umar", "Umer", "Umut", "Urban", "Uri", "Usman", "Uzair", "Uzayr", "Valen", "Valentin", "Valentino", "Valery", "Valo", "Vasyl", "Vedantsinh", "Veeran", "Victor", "Victory", "Vinay", "Vince", "Vincent", "Vincenzo", "Vinh", "Vinnie", "Vithujan", "Vladimir", "Vladislav", "Vrishin", "Vuyolwethu", "Wabuya", "Wai", "Walid", "Wallace", "Walter", "Waqaas", "Warkhas", "Warren", "Warrick", "Wasif", "Wayde", "Wayne", "Wei", "Wen", "Wesley", "Wesley-Scott", "Wiktor", "Wilkie", "Will", "William", "William-John", "Willum", "Wilson", "Windsor", "Wojciech", "Woyenbrakemi", "Wyatt", "Wylie", "Wynn", "Xabier", "Xander", "Xavier", "Xiao", "Xida", "Xin", "Xue", "Yadgor", "Yago", "Yahya", "Yakup", "Yang", "Yanick", "Yann", "Yannick", "Yaseen", "Yasin", "Yasir", "Yassin", "Yoji", "Yong", "Yoolgeun", "Yorgos", "Youcef", "Yousif", "Youssef", "Yu", "Yuanyu", "Yuri", "Yusef", "Yusuf", "Yves", "Zaaine", "Zaak", "Zac", "Zach", "Zachariah", "Zacharias", "Zacharie", "Zacharius", "Zachariya", "Zachary", "Zachary-Marc", "Zachery", "Zack", "Zackary", "Zaid", "Zain", "Zaine", "Zaineddine", "Zainedin", "Zak", "Zakaria", "Zakariya", "Zakary", "Zaki", "Zakir", "Zakk", "Zamaar", "Zander", "Zane", "Zarran", "Zayd", "Zayn", "Zayne", "Ze", "Zechariah", "Zeek", "Zeeshan", "Zeid", "Zein", "Zen", "Zendel", "Zenith", "Zennon", "Zeph", "Zerah", "Zhen", "Zhi", "Zhong", "Zhuo", "Zi", "Zidane", "Zijie", "Zinedine", "Zion", "Zishan", "Ziya", "Ziyaan", "Zohaib", "Zohair", "Zoubaeir", "Zubair", "Zubayr", "Zuriel" };
            string[] lastName = { "Anderson", "Ashwoon", "Aikin", "Bateman", "Bongard", "Bowers", "Boyd", "Cannon", "Cast", "Deitz", "Dewalt", "Ebner", "Frick", "Hancock", "Haworth", "Hesch", "Hoffman", "Kassing", "Knutson", "Lawless", "Lawicki", "Mccord", "McCormack", "Miller", "Myers", "Nugent", "Ortiz", "Orwig", "Ory", "Paiser", "Pak", "Pettigrew", "Quinn", "Quizoz", "Ramachandran", "Resnick", "Sagar", "Schickowski", "Schiebel", "Sellon", "Severson", "Shaffer", "Solberg", "Soloman", "Sonderling", "Soukup", "Soulis", "Stahl", "Sweeney", "Tandy", "Trebil", "Trusela", "Trussel", "Turco", "Uddin", "Uflan", "Ulrich", "Upson", "Vader", "Vail", "Valente", "Van Zandt", "Vanderpoel", "Ventotla", "Vogal", "Wagle", "Wagner", "Wakefield", "Weinstein", "Weiss", "Woo", "Yang", "Yates", "Yocum", "Zeaser", "Zeller", "Ziegler", "Bauer", "Baxster", "Casal", "Cataldi", "Caswell", "Celedon", "Chambers", "Chapman", "Christensen", "Darnell", "Davidson", "Davis", "DeLorenzo", "Dinkins", "Doran", "Dugelman", "Dugan", "Duffman", "Eastman", "Ferro", "Ferry", "Fletcher", "Fietzer", "Hylan", "Hydinger", "Illingsworth", "Ingram", "Irwin", "Jagtap", "Jenson", "Johnson", "Johnsen", "Jones", "Jurgenson", "Kalleg", "Kaskel", "Keller", "Leisinger", "LePage", "Lewis", "Linde", "Lulloff", "Maki", "Martin", "McGinnis", "Mills", "Moody", "Moore", "Napier", "Nelson", "Norquist", "Nuttle", "Olson", "Ostrander", "Reamer", "Reardon", "Reyes", "Rice", "Ripka", "Roberts", "Rogers", "Root", "Sandstrom", "Sawyer", "Schlicht", "Schmitt", "Schwager", "Schutz", "Schuster", "Tapia", "Thompson", "Tiernan", "Tisler" };
            string[] status = { "New (10%)", "First contact(10%)", "In process (30%)", "Proposal/Price Quote (50%)", "Negotiation/Review (90%)", "Closed won(100%)", "Closed lost(0%)" };
            //string[] status = { "Closed won(100%)", "Closed lost(0%)" };
            #endregion
            //must have for excel handeling
            Dictionary<string, object> ess = MakeExcelEss(opportunitesDB);
            object[,] values = (object[,])(ess["Range"] as Excel.Range).Value2;
            int r = values.GetLength(0);
            if (values[r, 1] != null) r++;
            Random rnd = new Random();
            //
            try
            {
                for (int i = r; i < k + r; i++)
                {
                    string s = "";
                    for (int j = 0; j < 9; j++)
                        s += rnd.Next(j == 0 ? 5 : 10);
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 1] = s;
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 2] = names[rnd.Next(0, names.Length)];
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 3] = lastName[rnd.Next(0, lastName.Length)];
                    s = "05";
                    for (int j = 0; j < 8; j++)
                        s += rnd.Next(10);
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 4] = s;
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 5] = (ess["Sheet"] as Excel.Worksheet).Cells[i, 2].Value.ToString() + "_" + (ess["Sheet"] as Excel.Worksheet).Cells[i, 3].Value.ToString() + "@" + Convert.ToChar(rnd.Next(Convert.ToInt32('A'), Convert.ToInt32('Z') + 1)) + "mail.com";
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 6] = status[rnd.Next(0, status.Length)];
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 7] = userList[rnd.Next(userList.Count)].ID;
                    int sub = rnd.Next(31) * -1;
                    DateTime d = DateTime.Now.AddDays(sub);
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 8] = d.Date;
                    (ess["Sheet"] as Excel.Worksheet).Cells[i, 9] = "don’t know what he wants in his life";
                    GenerateRndPack2((ess["Sheet"] as Excel.Worksheet).Cells[i, 1].Value, rnd.Next(1, 4), d, (ess["Book"] as Excel.Workbook));
                }
               (ess["Book"] as Excel.Workbook).Save();
                CleanExcelEss(ref ess);
            }
            catch
            {
                MessageBox.Show("coudlnt generage");
                return;
            }
            finally
            {
                //must have for excel handeling

            }
            //
        }

        private static void GenerateRndPack(int k)
        {
            //must have for excel handeling
            Excel.Application MyApp = new Excel.Application();
            Excel.Workbook MyBook = MyApp.Workbooks.Open(opportunitesDB);
            Excel.Worksheet MySheet = (Excel.Worksheet)MyBook.Sheets[2];
            Excel.Range xlRange = MySheet.UsedRange;
            int r = xlRange.Rows.Count;
            if (r != 0)
                r++;
            Random rnd = new Random();
            //
            for (int i = r; i < k + r; i++)
            {
                int rn = rnd.Next(0, opportunites.Count);
                MySheet.Cells[i, 1] = opportunites[rn].ID;
                string s = "05";
                for (int j = 0; j < 8; j++)
                    s += rnd.Next(10);
                MySheet.Cells[i, 2] = s;
                MySheet.Cells[i, 3] = rnd.Next(1, 4).ToString();
                DateTime d = opportunites[rn].treatedAt.Date.AddDays(-7 + rnd.Next(8));
                MySheet.Cells[i, 4] = d.Date;
                MySheet.Cells[i, 5] = d.AddMonths(rnd.Next(6, 13));
            }
            MyBook.Save();
            //must have for excel handeling
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(MySheet);
            MyBook.Close();
            Marshal.ReleaseComObject(MyBook);
            MyApp.Quit();
            Marshal.ReleaseComObject(MyApp);
            //
        }

        public static int GetStatusPrec(string status)
        {
            return Convert.ToInt32(status.Substring(status.IndexOf('(') + 1, status.IndexOf('%') - 1 - status.IndexOf('(')));
        }

        public static float GetPackagePrice(int packageType)
        {
            switch (packageType)
            {
                case 1:
                    {
                        return 30;
                    }
                case 2:
                    {
                        return 20;
                    }
                case 3:
                    {
                        return 10;
                    }
                default:
                    {
                        return 0;
                    }
            }
        }

        public static float[] GetStatistics(User u)//0-status
        {
            float[] ret = { 0, 0, 0, 0, 0 };
            float[] overall = GetStatistics();
            foreach (Opp p in opportunites)
            {
                if (p.treatedBy.ID == u.ID)
                {
                    ret[3]++;//number of deals
                    ret[0] += GetStatusPrec(p.status); //status
                    switch (GetStatusPrec(p.status))
                    {
                        case 0:
                            {
                                ret[2]++;
                                break;
                            }
                        case 100:
                            {
                                ret[1]++;
                                break;
                            }
                    }
                    foreach (Package pc in packages)
                        if (pc.ID == p.ID)
                            ret[4] += GetPackagePrice(pc.packageType);
                }
            }
            ret[1] /= overall[1];
            ret[1] *= 100;
            ret[2] /= overall[2];
            ret[2] *= 100;
            ret[0] /= ret[3] != 0 ? ret[3] : 1;
            ret[4] /= ret[3] != 0 ? ret[3] : 1;
            return ret;
        }

        public static float[] GetStatistics()//0-status
        {
            float[] ret = { 0, 0, 0, 0, 0 };
            foreach (Opp p in opportunites)
            {
                ret[3]++;//number of deals
                ret[0] += GetStatusPrec(p.status); //status
                switch (GetStatusPrec(p.status))
                {
                    case 0:
                        {
                            ret[2]++;
                            break;
                        }
                    case 100:
                        {
                            ret[1]++;
                            break;
                        }
                }
            }
            foreach (Package pc in packages)
                ret[4] += GetPackagePrice(pc.packageType);
            ret[0] /= ret[3];
            ret[4] /= ret[3];
            return ret;
        }
        [STAThread]
        static void Main()
        {
            init();
            //GenerateRndOpp(400);
            //init();
            //foreach (Opp o in opportunites)
            //    if (o.status.ToUpper().Contains("CLOSED"))
            //        MovetHistory(o.ID);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            new LogIn().ShowDialog();
            if (currentUser.ID != null)
                Application.Run(new MainWin());
        }
    }
}
