using OpenExcelExport;
using OpenExcelSdk;
using System.Reflection;

Assembly assembly = Assembly.GetExecutingAssembly();
Version version = assembly.GetName().Version;

string vers = version.ToString();

// check arguments
if (args.Length==0)
{
    HelpPrinter.PrintHelp(vers);
    return;
}

// parse arguments
string argLine = string.Join(" ", args);


foreach (string s in args)
{
    Console.WriteLine("arg : " + s);
}
//Console.WriteLine("args : " + argLine);

//--XXXEDEBUG:
//string argLine = " -excel = \"CellFormat.xlsx\"  -out =  \"out\"  ";
//string argLine = "-excel=\"file.xlsx\"  -out=\"out.xlsx\"";

//argLine = "-excel=\"aaa\"";


if (!ArgsParser.Parse(argLine, out ProgParams progParams, out string errMsg))
{
    Console.WriteLine(errMsg);
    return;
}

// add xlsx extension if not exists
if (Path.GetExtension(progParams.OutputExcelFile).Length==0) 
{
    progParams.OutputExcelFile += ".xlsx";
}

Console.WriteLine("Ok, Will analyse the Excel file : " + progParams.InputExcelFile);

// remove previous result file
if(File.Exists(progParams.OutputExcelFile))
{
    Console.WriteLine("Remove previous result file : " + progParams.OutputExcelFile);
    File.Delete(progParams.OutputExcelFile);
}

// check input file exists
if (File.Exists(progParams.InputExcelFile) == false)
{
    Console.WriteLine("Error, input excel file does not exist : " + progParams.InputExcelFile);
    return;
}

ExcelProcessor excelProcessor = new ExcelProcessor();

try
{
    excelProcessor.ExportAllStyles(progParams.InputExcelFile, progParams.OutputExcelFile);
    Console.WriteLine("=> Ok, analysis done, result : " + progParams.OutputExcelFile);
}
catch (Exception ex)
{
    Console.WriteLine("Error, exception occurs during excel styles export : " + ex.Message);
    return;
}
