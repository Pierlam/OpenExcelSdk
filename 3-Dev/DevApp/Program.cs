using DevApp;
using DocumentFormat.OpenXml.Math;
using OpenExcelSdk;
using OpenExcelSdk.System;

void ConvertDouble()
{
    string value = "45927.524259259262";
    value = value.Replace('.', ',');

    //string value = "45927,524";
    double valDouble = double.Parse(value);

}

Console.WriteLine("=> OpenExcelSdk DevApp:");

//CellReader.Read();

ConvertDouble();

Console.WriteLine("=> Ok, Ends." );
