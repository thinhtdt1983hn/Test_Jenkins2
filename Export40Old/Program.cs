using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


// See https://aka.ms/new-console-template for more information
Console.WriteLine("Thong ke 40 so trong 6 so tiep theo!");

//Create COM Objects. Create a COM object for everything that is referenced
Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Downloads\Demo_data\Data_temp.xlsx");
_Worksheet xlWorksheet = xlWorkbook.Sheets[1];
Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

int rowCount = xlRange.Rows.Count;
//int colCount = xlRange.Columns.Count;

//====================================
var _1G1_data = new List<int>();// Lưu các số đã ra, sắp xếp từ mới đến cũ
var _1DB_data = new List<int>();// Lưu các số đã ra, sắp xếp từ mới đến cũ
var _1B_data = new List<int>();// Lưu các số đã ra, sắp xếp từ mới đến cũ

var _1G1_dataT = new List<int>();// Lưu 10 số hàng chục
var _1G1_dataP = new List<int>();// Lưu 10 số hàng đơn vị
var _1DB_dataT = new List<int>();// Lưu 10 số hàng chục
var _1DB_dataP = new List<int>();// Lưu 10 số hàng chục
var _1B_dataT = new List<int>();// Lưu 10 số hàng chục ngàn
var _1B_dataP = new List<int>();// Lưu 10 số hàng ngàn
//====================================

// Đọc data các số đã ra từ file Excel
// Lưu vào các biến: _1G1_data, _1DB_data, _1B_data
for (int i = rowCount; i >= 1; i--)
{
    string _1G1_temp = xlRange.Cells[i, 1].Value2.ToString().Substring(3);
    string _temp = xlRange.Cells[i, 2].Value2.ToString();

    string _1B_temp = _temp.Substring(0, 2);
    string _1DB_temp = string.Empty;

    if (_temp.Length > 5) _1DB_temp = _temp.Substring(4);
    else _1DB_temp = _temp.Substring(3);


    if (!_1G1_data.Contains(Int32.Parse(_1G1_temp)) && _1G1_data.Count <= 100) _1G1_data.Add(Int32.Parse(_1G1_temp));

    if (!_1DB_data.Contains(Int32.Parse(_1DB_temp)) && _1DB_data.Count <= 100) _1DB_data.Add(Int32.Parse(_1DB_temp));

    if (!_1B_data.Contains(Int32.Parse(_1B_temp)) && _1B_data.Count <= 100) _1B_data.Add(Int32.Parse(_1B_temp));


    if (!_1G1_dataT.Contains(Int32.Parse(_1G1_temp.Substring(0, 1)))) _1G1_dataT.Add(Int32.Parse(_1G1_temp.Substring(0, 1)));
    if (!_1G1_dataP.Contains(Int32.Parse(_1G1_temp.Substring(1)))) _1G1_dataP.Add(Int32.Parse(_1G1_temp.Substring(1)));

    if (!_1DB_dataT.Contains(Int32.Parse(_1DB_temp.Substring(0, 1)))) _1DB_dataT.Add(Int32.Parse(_1DB_temp.Substring(0, 1)));
    if (!_1DB_dataP.Contains(Int32.Parse(_1DB_temp.Substring(1)))) _1DB_dataP.Add(Int32.Parse(_1DB_temp.Substring(1)));

    if (!_1B_dataT.Contains(Int32.Parse(_1B_temp.Substring(0, 1)))) _1B_dataT.Add(Int32.Parse(_1B_temp.Substring(0, 1)));
    if (!_1B_dataP.Contains(Int32.Parse(_1B_temp.Substring(1)))) _1B_dataP.Add(Int32.Parse(_1B_temp.Substring(1)));
}

//for (int i = rowCount; i >= 1; i--)
//{
//    string _1G1_temp = xlRange.Cells[i, 1].Value2.ToString().Substring(3);
//    string _temp1 = xlRange.Cells[i, 2].Value2.ToString();

//    string _1B_temp = _temp1.Substring(0, 2);
//    string _1DB_temp = string.Empty;

//    if (_temp1.Length > 5) _1DB_temp = _temp1.Substring(4);
//    else _1DB_temp = _temp1.Substring(3);


//    if (!_1G1_dataT.Contains(Int32.Parse(_1G1_temp.Substring(0, 1)))) _1G1_dataT.Add(Int32.Parse(_1G1_temp.Substring(0, 1)));
//    if (!_1G1_dataP.Contains(Int32.Parse(_1G1_temp.Substring(1)))) _1G1_dataP.Add(Int32.Parse(_1G1_temp.Substring(1)));

//    if (!_1DB_dataT.Contains(Int32.Parse(_1DB_temp.Substring(0, 1)))) _1DB_dataT.Add(Int32.Parse(_1DB_temp.Substring(0, 1)));
//    if (!_1DB_dataP.Contains(Int32.Parse(_1DB_temp.Substring(1)))) _1DB_dataP.Add(Int32.Parse(_1DB_temp.Substring(1)));

//    if (!_1B_dataT.Contains(Int32.Parse(_1B_temp.Substring(0, 1)))) _1B_dataT.Add(Int32.Parse(_1B_temp.Substring(0, 1)));
//    if (!_1B_dataP.Contains(Int32.Parse(_1B_temp.Substring(1)))) _1B_dataP.Add(Int32.Parse(_1B_temp.Substring(1)));
//}


//cleanup
GC.Collect();
GC.WaitForPendingFinalizers();

//rule of thumb for releasing com objects:
//  never use two dots, all COM objects must be referenced and released individually
//  ex: [somthing].[something].[something] is bad

//release com objects to fully kill excel process from running in the background
Marshal.ReleaseComObject(xlRange);
Marshal.ReleaseComObject(xlWorksheet);

//close and release
xlWorkbook.Close();
Marshal.ReleaseComObject(xlWorkbook);

//quit and release
xlApp.Quit();
Marshal.FinalReleaseComObject(xlApp);


// Bổ sung nếu chưa ra đủ 100 số
if (_1G1_data.Count < 100)
{
    for (int i = 0; i < 100; i++)
    {
        if (!_1G1_data.Contains(i)) _1G1_data.Add(i);
    }
}
if (_1DB_data.Count < 100)
{
    for (int i = 0; i < 100; i++)
    {
        if (!_1DB_data.Contains(i)) _1DB_data.Add(i);
    }
}
if (_1B_data.Count < 100)
{
    for (int i = 0; i < 100; i++)
    {
        if (!_1B_data.Contains(i)) _1B_data.Add(i);
    }
}


_1G1_dataT.Remove(_1G1_dataT[0]);
_1G1_dataT.Remove(_1G1_dataT[0]);
_1G1_dataT.Remove(_1G1_dataT[0]);
_1G1_dataT.Remove(_1G1_dataT[0]);

// Chuyển thành chuỗi số
string _1G1_KQ_T = string.Empty;
foreach (var item in _1G1_dataT)
{
    for (int i = 0; i < 10; i++)
    {
        _1G1_KQ_T += item + i.ToString() + ",";
    }
}

_1G1_dataP.Remove(_1G1_dataP[0]);
_1G1_dataP.Remove(_1G1_dataP[0]);
_1G1_dataP.Remove(_1G1_dataP[0]);
_1G1_dataP.Remove(_1G1_dataP[0]);

// Chuyển thành chuỗi số
string _1G1_KQ_P = string.Empty;
foreach (var item in _1G1_dataP)
{
    for (int i = 0; i < 10; i++)
    {
        _1G1_KQ_P += item + i.ToString() + ",";
    }
}

_1DB_dataT.Remove(_1DB_dataT[0]);
_1DB_dataT.Remove(_1DB_dataT[0]);
_1DB_dataT.Remove(_1DB_dataT[0]);
_1DB_dataT.Remove(_1DB_dataT[0]);

// Chuyển thành chuỗi số
string _1DB_KQ_T = string.Empty;
foreach (var item in _1DB_dataT)
{
    for (int i = 0; i < 10; i++)
    {
        _1DB_KQ_T += item + i.ToString() + ",";
    }
}

_1DB_dataP.Remove(_1DB_dataP[0]);
_1DB_dataP.Remove(_1DB_dataP[0]);
_1DB_dataP.Remove(_1DB_dataP[0]);
_1DB_dataP.Remove(_1DB_dataP[0]);

// Chuyển thành chuỗi số
string _1DB_KQ_P = string.Empty;
foreach (var item in _1DB_dataP)
{
    for (int i = 0; i < 10; i++)
    {
        _1DB_KQ_P += item + i.ToString() + ",";
    }
}

_1B_dataT.Remove(_1B_dataT[0]);
_1B_dataT.Remove(_1B_dataT[0]);
_1B_dataT.Remove(_1B_dataT[0]);
_1B_dataT.Remove(_1B_dataT[0]);

// Chuyển thành chuỗi số
string _1B_KQ_T = string.Empty;
foreach (var item in _1B_dataT)
{
    for (int i = 0; i < 10; i++)
    {
        _1B_KQ_T += item + i.ToString() + ",";
    }
}

_1B_dataP.Remove(_1B_dataP[0]);
_1B_dataP.Remove(_1B_dataP[0]);
_1B_dataP.Remove(_1B_dataP[0]);
_1B_dataP.Remove(_1B_dataP[0]);

// Chuyển thành chuỗi số
string _1B_KQ_P = string.Empty;
foreach (var item in _1B_dataP)
{
    for (int i = 0; i < 10; i++)
    {
        _1B_KQ_P += item + i.ToString() + ",";
    }
}

// Loại 10 số gan nhất
int _countDeleT = 0;
int _countDeleP = 0;
for (int i = _1G1_data.Count - 1; i >= 0; i--)
{
    string _temp = _1G1_data[i].ToString() + ",";
    if (_1G1_data[i] < 10) _temp = "0" + _1G1_data[i].ToString() + ",";
    else _temp = _1G1_data[i].ToString() + ",";


    if (_1G1_KQ_T.Contains(_temp) && _countDeleT < 10)
    {
        _1G1_KQ_T = _1G1_KQ_T.Replace(_temp, "");
        _countDeleT++;
    }

    if (_1G1_KQ_P.Contains(_temp) && _countDeleP < 10)
    {
        _1G1_KQ_P = _1G1_KQ_P.Replace(_temp, "");
        _countDeleP++;
    }
}

_countDeleT = 0;
_countDeleP = 0;
for (int i = _1DB_data.Count - 1; i >= 0; i--)
{
    string _temp = _1DB_data[i].ToString() + ",";
    if (_1DB_data[i] < 10) _temp = "0" + _1DB_data[i].ToString() + ",";
    else _temp = _1DB_data[i].ToString() + ",";


    if (_1DB_KQ_T.Contains(_temp) && _countDeleT < 10)
    {
        _1DB_KQ_T = _1DB_KQ_T.Replace(_temp, "");
        _countDeleT++;
    }

    if (_1DB_KQ_P.Contains(_temp) && _countDeleP < 10)
    {
        _1DB_KQ_P = _1DB_KQ_P.Replace(_temp, "");
        _countDeleP++;
    }
}

_countDeleT = 0;
_countDeleP = 0;
for (int i = _1B_data.Count - 1; i >= 0; i--)
{
    string _temp = _1B_data[i].ToString() + ",";
    if (_1B_data[i] < 10) _temp = "0" + _1B_data[i].ToString() + ",";
    else _temp = _1B_data[i].ToString() + ",";


    if (_1B_KQ_T.Contains(_temp) && _countDeleT < 10)
    {
        _1B_KQ_T = _1B_KQ_T.Replace(_temp, "");
        _countDeleT++;
    }

    if (_1B_KQ_P.Contains(_temp) && _countDeleP < 10)
    {
        _1B_KQ_P = _1B_KQ_P.Replace(_temp, "");
        _countDeleP++;
    }
}


// Loại 10 số mới ra gần nhất
_countDeleT = 0;
_countDeleP = 0;
for (int i = 0; i < _1G1_data.Count; i++)
{
    string _temp = _1G1_data[i].ToString() + ",";
    if (_1G1_data[i] < 10) _temp = "0" + _1G1_data[i].ToString() + ",";
    else _temp = _1G1_data[i].ToString() + ",";


    if (_1G1_KQ_T.Contains(_temp) && _countDeleT < 10)
    {
        _1G1_KQ_T = _1G1_KQ_T.Replace(_temp, "");
        _countDeleT++;
    }

    if (_1G1_KQ_P.Contains(_temp) && _countDeleP < 10)
    {
        _1G1_KQ_P = _1G1_KQ_P.Replace(_temp, "");
        _countDeleP++;
    }
}

_countDeleT = 0;
_countDeleP = 0;
for (int i = 0; i< _1DB_data.Count; i++)
{
    string _temp = _1DB_data[i].ToString() + ",";
    if (_1DB_data[i] < 10) _temp = "0" + _1DB_data[i].ToString() + ",";
    else _temp = _1DB_data[i].ToString() + ",";


    if (_1DB_KQ_T.Contains(_temp) && _countDeleT < 10)
    {
        _1DB_KQ_T = _1DB_KQ_T.Replace(_temp, "");
        _countDeleT++;
    }

    if (_1DB_KQ_P.Contains(_temp) && _countDeleP < 10)
    {
        _1DB_KQ_P = _1DB_KQ_P.Replace(_temp, "");
        _countDeleP++;
    }
}

_countDeleT = 0;
_countDeleP = 0;
for (int i = 0; i < _1B_data.Count; i++)
{
    string _temp = _1B_data[i].ToString() + ",";
    if (_1B_data[i] < 10) _temp = "0" + _1B_data[i].ToString() + ",";
    else _temp = _1B_data[i].ToString() + ",";


    if (_1B_KQ_T.Contains(_temp) && _countDeleT < 10)
    {
        _1B_KQ_T = _1B_KQ_T.Replace(_temp, "");
        _countDeleT++;
    }

    if (_1B_KQ_P.Contains(_temp) && _countDeleP < 10)
    {
        _1B_KQ_P = _1B_KQ_P.Replace(_temp, "");
        _countDeleP++;
    }
}


// Show ra màn hình
Console.WriteLine("1G1_T: 40 so trong 6 so cuoi ({0},{1},{2},{3},{4},{5}) HANG CHUC cua GIAI NHAT. KET QUA: ", _1G1_dataT[0], _1G1_dataT[1], _1G1_dataT[2], _1G1_dataT[3], _1G1_dataT[4], _1G1_dataT[5]);
Console.WriteLine(_1G1_KQ_T);
Console.WriteLine(" ");
Console.WriteLine("------------------------------------");
Console.WriteLine(" ");

Console.WriteLine("1G1_P: 40 so trong 6 so cuoi ({0},{1},{2},{3},{4},{5}) HANG DON VI cua GIAI NHAT. KET QUA: ", _1G1_dataP[0], _1G1_dataP[1], _1G1_dataP[2], _1G1_dataP[3], _1G1_dataP[4], _1G1_dataP[5]);
Console.WriteLine(_1G1_KQ_P);
Console.WriteLine(" ");
Console.WriteLine("------------------------------------");
Console.WriteLine(" ");

Console.WriteLine("1DB_T: 40 so trong 6 so cuoi ({0},{1},{2},{3},{4},{5}) HANG CHUC cua GIAI DAC BIET. KET QUA: ", _1DB_dataT[0], _1DB_dataT[1], _1DB_dataT[2], _1DB_dataT[3], _1DB_dataT[4], _1DB_dataT[5]);
Console.WriteLine(_1DB_KQ_T);
Console.WriteLine(" ");
Console.WriteLine("------------------------------------");
Console.WriteLine(" ");

Console.WriteLine("1DB_P: 40 so trong 6 so cuoi ({0},{1},{2},{3},{4},{5}) HANG DON VI cua GIAI DAC BIET. KET QUA: ", _1DB_dataP[0], _1DB_dataP[1], _1DB_dataP[2], _1DB_dataP[3], _1DB_dataP[4], _1DB_dataP[5]);
Console.WriteLine(_1DB_KQ_P);
Console.WriteLine(" ");
Console.WriteLine("------------------------------------");
Console.WriteLine(" ");

Console.WriteLine("1B_T: 40 so trong 6 so cuoi ({0},{1},{2},{3},{4},{5}) HANG CHUC NGAN cua GIAI DAC BIET. KET QUA: ", _1B_dataT[0], _1B_dataT[1], _1B_dataT[2], _1B_dataT[3], _1B_dataT[4], _1B_dataT[5]);
Console.WriteLine(_1B_KQ_T);
Console.WriteLine(" ");
Console.WriteLine("------------------------------------");
Console.WriteLine(" ");

Console.WriteLine("1B_P: 40 so trong 6 so cuoi ({0},{1},{2},{3},{4},{5}) HANG NGAN cua GIAI DAC BIET. KET QUA: ", _1B_dataP[0], _1B_dataP[1], _1B_dataP[2], _1B_dataP[3], _1B_dataP[4], _1B_dataP[5]);
Console.WriteLine(_1B_KQ_P);
Console.WriteLine(" ");
Console.WriteLine(" ");
Console.WriteLine("=================================================");