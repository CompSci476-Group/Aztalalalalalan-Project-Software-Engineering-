using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.Reflection;
using Microsoft.Office.Interop.Excel;

public class CreateExcelWorksheet
{
    static void Main()
    {

        Console.Write("please enter a number\n");
        string choice =Console.ReadLine();
        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        if (xlApp == null)
        {
            Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
            return;
        }
        xlApp.Visible = true;

        Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        Worksheet ws = (Worksheet)wb.Worksheets[1];

        if (ws == null)
        {
            Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
        }

        // Select the Excel cells, in the range c1 to c7 in the worksheet.
        Range aRange = ws.get_Range("C1", "C7");//this is used for the selection of what is changing

        if (aRange == null)
        {
            Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
        }

        // Fill the cells in the C1 to C7 range of the worksheet with the number 6.
        Object[] args = new Object[1];
        args[0] = 6;
        aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);

        // Change the cells in the C1 to C7 range of the worksheet to the number 8.
        aRange.Value2 = choice;


        //saving the file
        Console.WriteLine("would you like to save? (Y or N)");
        string userIn = Console.ReadLine();
        if (userIn.Equals("Y"))
        {
            save(xlApp, wb);
        }
        else
        {
        }

    }

    static void save(Application Excel, Workbook THING)
    {
        string path = @"C:\Users\Sam Kromm\Documents\Aztalalalalalan-Project-Software-Engineering-\CreateExcelWorksheet\CreateExcelWorksheet";
        Excel.DisplayAlerts = false;
        THING.Application.DefaultFilePath = path;
        THING.SaveAs("Test", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
    }
}

