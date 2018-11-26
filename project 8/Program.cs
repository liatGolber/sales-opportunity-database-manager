﻿using System;
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

    static class Program
    {
        //databases paths
        public static string userDB = Environment.CurrentDirectory + @"\Users.xlsx";//1-id 2-name 3-last 4-password 5-isadmin
        public static string historyDB = Environment.CurrentDirectory + @"\History.xlsx";
        public static string leadsDB = Environment.CurrentDirectory + @"\Leads.xlsx";
        public static string opportunitesDB = Environment.CurrentDirectory + @"\Opportunites.xlsx";

        public static User currentUser;

        public static User GetUserByID(string id)
        {
            User ret = new User();
            Excel.Application MyApp = new Excel.Application();
            Excel.Workbook MyBook = MyApp.Workbooks.Open(userDB);
            Excel.Worksheet MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            Excel.Range xlRange = MySheet.UsedRange;
            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (xlRange.Cells[i, 1].Value.ToString() == id)
                {
                    ret.ID = xlRange.Cells[i, 1].Value.ToString();
                    ret.name = xlRange.Cells[i, 2].Value.ToString();
                    ret.lastN = xlRange.Cells[i, 3].Value.ToString();
                    ret.password = xlRange.Cells[i, 4].Value.ToString();
                    ret.isAdmin = Convert.ToBoolean(xlRange.Cells[i, 5].Value.ToString());

                }
            }

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(MySheet);

            MyBook.Close();
            Marshal.ReleaseComObject(MyBook);
            MyApp.Quit();
            Marshal.ReleaseComObject(MyApp);

            return ret;
        }
        public static User InsertNewUser(string id, string name, string lName, string password, bool isAdmin)
        {
            User ret = new User();
            Excel.Application MyApp = new Excel.Application();
            Excel.Workbook MyBook = MyApp.Workbooks.Open(userDB);
            Excel.Worksheet MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            Excel.Range xlRange = MySheet.UsedRange;
            int r = xlRange.Rows.Count + 1;
            MySheet.Cells[r, 1] = id;
            MySheet.Cells[r, 2] = name;
            MySheet.Cells[r, 3] = lName;
            MySheet.Cells[r, 4] = password;
            MySheet.Cells[r, 5] = isAdmin;
            MyBook.Save();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(MySheet);

            MyBook.Close();
            Marshal.ReleaseComObject(MyBook);
            MyApp.Quit();
            Marshal.ReleaseComObject(MyApp);

            return ret;
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            new LogIn().ShowDialog();
            if (currentUser.ID != null)
                Application.Run(new MainWin());
        }
    }
}
