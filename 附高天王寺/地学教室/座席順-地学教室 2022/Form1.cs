using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;



namespace 座席順_地学教室_2022
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)  //ロード時
        {
            MessageBox.Show("大天才上尾先輩大感謝", "ペア作成ツール 2022 - 地学教室 Edition", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void files_Click(object sender, EventArgs e)  //ファイル
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)  //ヘルプ
        {
            System.Diagnostics.Process.Start("cmd", "/C start https://www.bing.com/search?q=windows+%e3%81%a7%e3%83%98%e3%83%ab%e3%83%97%e3%82%92%e8%a1%a8%e7%a4%ba%e3%81%99%e3%82%8b%e6%96%b9%e6%b3%95&filters=guid:%224028227-ja-dia%22%20lang:%22ja%22&form=S00028");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            main();
        }

        private void main()  //偶数奇数選別
        {
            string thank = "大天才上尾先輩大感謝";
            string title = "ペア作成ツール 2022 - 地学教室 Edition";
            string students_txt = textBox1.Text;
            int students = int.Parse(textBox1.Text);  //studentsとは出席番号最後の人、即ちクラスの人数である。
            if(students <= 2 || students > 40)
            {
                string naiyo = "「" + students_txt + "」は無効な文字列です。\n\r半角数字で入力して下さい。";
                MessageBox.Show(naiyo, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Application.Exit();
            }

            string naiyo2 = students_txt + "で続行します。"; 
            MessageBox.Show(naiyo2, title, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            
            DateTime dt = DateTime.Now; //時刻を取得
            string timenow = dt.ToString("yyyy_MM_dd HH_mm_ss") +  " 生徒" + students_txt + "人 ";  //時刻をフォーマット


            string font = "HGSｺﾞｼｯｸM";

            if (students % 2 == 1)  //奇数のとき
            {
                for (int i = 1; i <= students; i++) //iは1~studentsまでの数 第i回目の授業
                {

                    var workbook = new XLWorkbook();
                    var worksheet_2 = workbook.Worksheets.Add("座席表");  //座席表 シート作成
                    var worksheet_1 = workbook.Worksheets.Add("ペア");  //ペア シート作成

                    //以下座席表初期設定
                    worksheet_2.Column(1).Width = 0;
                    worksheet_2.Column(4).Width = 4;
                    worksheet_2.Column(9).Width = 4;

                    worksheet_2.Column(2).Width = 16;
                    worksheet_2.Column(3).Width = 16;
                    worksheet_2.Column(5).Width = 16;
                    worksheet_2.Column(6).Width = 16;
                    worksheet_2.Column(7).Width = 16;
                    worksheet_2.Column(8).Width = 16;
                    worksheet_2.Column(10).Width = 16;
                    worksheet_2.Column(11).Width = 16;
                    worksheet_2.Column(12).Width = 16;
                    worksheet_2.Column(13).Width = 16;

                    worksheet_2.Row(1).Height = 48;
                    worksheet_2.Row(2).Height = 48;
                    worksheet_2.Row(3).Height = 48;
                    worksheet_2.Row(4).Height = 48;
                    worksheet_2.Row(5).Height = 48;
                    worksheet_2.Row(6).Height = 48;
                    worksheet_2.Row(7).Height = 48;
                    worksheet_2.Row(8).Height = 48;
                    worksheet_2.Row(9).Height = 48;

                    worksheet_2.Range("B4:C7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("E4:F7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("G4:H7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("J4:K7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("L4:M7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);

                    worksheet_2.Range("C4:C7").Style.Border.SetLeftBorder(XLBorderStyleValues.MediumDashed);
                    worksheet_2.Range("F4:F7").Style.Border.SetLeftBorder(XLBorderStyleValues.MediumDashed);
                    worksheet_2.Range("G4:G7").Style.Border.SetRightBorder(XLBorderStyleValues.MediumDashed);
                    worksheet_2.Range("J4:J7").Style.Border.SetRightBorder(XLBorderStyleValues.MediumDashed);
                    worksheet_2.Range("L4:L7").Style.Border.SetRightBorder(XLBorderStyleValues.MediumDashed);

                    worksheet_2.Range("B4:C7").Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("E4:H7").Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("J4:M7").Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);

                    worksheet_2.Range("B2:M2").Merge();
                    worksheet_2.Range("B2:M2").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Cell("B2").Value = "教卓";
                    worksheet_2.Range("B9:M9").Merge();
                    worksheet_2.Cell("B9").Value = "大天才上尾先輩大感謝";

                    worksheet_2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet_2.Range("B2:M9").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet_2.Range("B1:M1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

                    worksheet_1.Range("A1:Z100").Style.Font.SetFontName(font);
                    worksheet_2.Range("A1:Z100").Style.Font.SetFontName(font);

                    worksheet_2.Range("A1:Z100").Style.Font.SetFontSize(28);

                    worksheet_2.Range("B1:E1").Merge();
                    worksheet_2.Cell("B1").Value = students_txt + "人 / 第" + i.ToString() + "回";
                    worksheet_2.Cell("B1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet_2.Cell("H1").Value = "年";
                    worksheet_2.Cell("K1").Value = "月";
                    worksheet_2.Cell("M1").Value = "日";
                    worksheet_2.Range("A1:M1").Style.Font.SetFontSize(20);
                    //座席表終わり

                    for (int j = 1; j <= students; j++)  //jは1~studentsまでの数, 1列目に1から最終出席番号まで数値入力
                    {
                        worksheet_1.Cell(j, 1).Value = j;
                    }
                    int l = 1;
                    for (int k = i; k >= 1; k--)
                    {
                        worksheet_1.Cell(l, 2).Value = k;
                        l = l + 1;
                    }
                    l = i + 1;
                    for (int m = students; m > i; m--)
                    {
                        worksheet_1.Cell(l, 2).Value = m;
                        l = l + 1;
                    }

                    int s = (students - 1) / 2;
                    if (i % 2 == 1)  //奇数回目のとき
                    {
                        int r = (i + 1) / 2 + s;
                        for (int t = 1; t <= s+1; t++)
                        {
                            worksheet_1.Row(r).Delete();
                            r--;
                        }
                        worksheet_1.Cell("A20").Value = (i + 1) / 2;
                    }
                    else  //偶数回目のとき
                    {
                        int r = i / 2 + 1 + s;
                        for (int t = 1; t <= s+1; t++)
                        {
                            worksheet_1.Row(r).Delete();
                            r--;
                        }
                        worksheet_1.Cell("A20").Value = i / 2 + s + 1;
                    }

                    
                    //ランダム配列
                    int[] ary0 = new int[(students - 1) / 2];
                    for(int num1 = 0; num1 <= (students -3) / 2; num1++)
                    {
                        ary0[num1] = num1 + 1;
                    }
                    ary0 = ary0.OrderBy(i => Guid.NewGuid()).ToArray();
                    int arn = 41;
                    int[] ary = new int[19] {arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn};
                    for (int num2 = 0; num2 <= (students - 3)/2; num2++)
                    {
                        int jk = ary0[num2];
                        ary[num2] = jk;
                    }
                    //ランダム配列終わり
                    
                    string a0 = ary[0].ToString();
                    string a1 = ary[1].ToString();
                    string a2 = ary[2].ToString();
                    string a3 = ary[3].ToString();
                    string a4 = ary[4].ToString();
                    string a5 = ary[5].ToString();
                    string a6 = ary[6].ToString();
                    string a7 = ary[7].ToString();
                    string a8 = ary[8].ToString();
                    string a9 = ary[9].ToString();

                    string b0 = ary[10].ToString();
                    string b1 = ary[11].ToString();
                    string b2 = ary[12].ToString();
                    string b3 = ary[13].ToString();
                    string b4 = ary[14].ToString();
                    string b5 = ary[15].ToString();
                    string b6 = ary[16].ToString();
                    string b7 = ary[17].ToString();
                    string b8 = ary[18].ToString();

                    worksheet_2.Cell("E4").FormulaA1 = "ペア!A" + a0;
                    worksheet_2.Cell("F4").FormulaA1 = "ペア!B" + a0;
                    worksheet_2.Cell("G4").FormulaA1 = "ペア!B" + a1;
                    worksheet_2.Cell("H4").FormulaA1 = "ペア!A" + a1;
                    worksheet_2.Cell("B4").FormulaA1 = "ペア!A" + a2;
                    worksheet_2.Cell("C4").FormulaA1 = "ペア!B" + a2;
                    worksheet_2.Cell("J4").FormulaA1 = "ペア!B" + a3;
                    worksheet_2.Cell("K4").FormulaA1 = "ペア!A" + a3;
                    worksheet_2.Cell("E5").FormulaA1 = "ペア!A" + a4;
                    worksheet_2.Cell("F5").FormulaA1 = "ペア!B" + a4;
                    worksheet_2.Cell("G5").FormulaA1 = "ペア!B" + a5;
                    worksheet_2.Cell("H5").FormulaA1 = "ペア!A" + a5;
                    worksheet_2.Cell("B5").FormulaA1 = "ペア!A" + a6;
                    worksheet_2.Cell("C5").FormulaA1 = "ペア!B" + a6;
                    worksheet_2.Cell("J5").FormulaA1 = "ペア!B" + a7;
                    worksheet_2.Cell("K5").FormulaA1 = "ペア!A" + a7;
                    worksheet_2.Cell("E6").FormulaA1 = "ペア!A" + a8;
                    worksheet_2.Cell("F6").FormulaA1 = "ペア!B" + a8;
                    worksheet_2.Cell("G6").FormulaA1 = "ペア!B" + a9;
                    worksheet_2.Cell("H6").FormulaA1 = "ペア!A" + a9;

                    worksheet_2.Cell("B6").FormulaA1 = "ペア!B" + b0;
                    worksheet_2.Cell("C6").FormulaA1 = "ペア!A" + b0;
                    worksheet_2.Cell("J6").FormulaA1 = "ペア!A" + b1;
                    worksheet_2.Cell("K6").FormulaA1 = "ペア!B" + b1;
                    worksheet_2.Cell("L6").FormulaA1 = "ペア!B" + b2;
                    worksheet_2.Cell("M6").FormulaA1 = "ペア!A" + b2;
                    worksheet_2.Cell("E7").FormulaA1 = "ペア!A" + b3;
                    worksheet_2.Cell("F7").FormulaA1 = "ペア!B" + b3;
                    worksheet_2.Cell("G7").FormulaA1 = "ペア!B" + b4;
                    worksheet_2.Cell("H7").FormulaA1 = "ペア!A" + b4;
                    worksheet_2.Cell("B7").FormulaA1 = "ペア!A" + b5;
                    worksheet_2.Cell("C7").FormulaA1 = "ペア!B" + b5;
                    worksheet_2.Cell("J7").FormulaA1 = "ペア!B" + b6;
                    worksheet_2.Cell("K7").FormulaA1 = "ペア!A" + b6;
                    worksheet_2.Cell("L7").FormulaA1 = "ペア!A" + b7;
                    worksheet_2.Cell("M7").FormulaA1 = "ペア!B" + b7;
                    if (students != 39)
                    {
                        worksheet_2.Cell("L5").FormulaA1 = "ペア!A20";
                    }
                    else
                    {
                        worksheet_2.Cell("L5").FormulaA1 = "ペア!A" + b8;
                        worksheet_2.Cell("M5").FormulaA1 = "ペア!B" + b8;
                        worksheet_2.Cell("L4").FormulaA1 = "ペア!A20";
                    }

                    for (int ko = 4; ko <= 8; ko++)
                    {
                        for (int ko2 = 2; ko2 <= 13; ko2++)
                        {
                            if (worksheet_2.Cell(ko, ko2).Value.ToString() == "0")
                            {
                                worksheet_2.Cell(ko, ko2).Value = "";
                            }
                        }
                    }
                    if (i < 10)
                    {
                        string asl = "0" + i;
                        workbook.SaveAs(@"C:\大天才上尾先輩大感謝\" + timenow + "大天才上尾先輩大感謝-地学教室座席順\\大天才上尾先輩大感謝-地学教室座席順 第" + asl + "回.xlsx");
                    }
                    else
                    {
                        string asl = "" + i;
                        workbook.SaveAs(@"C:\大天才上尾先輩大感謝\" + timenow + "大天才上尾先輩大感謝-地学教室座席順\\大天才上尾先輩大感謝-地学教室座席順 第" + asl + "回.xlsx");
                    }
                }
            }

            else  //偶数のとき
            {
                students = students - 1;
                for (int i = 1; i <= students; i++) //iは1~studentsまでの数 第i回目の授業
                {

                    var workbook = new XLWorkbook();
                    var worksheet_2 = workbook.Worksheets.Add("座席表");  //座席表 シート作成
                    var worksheet_1 = workbook.Worksheets.Add("ペア");  //ペア シート作成

                    //以下座席表初期設定
                    worksheet_2.Column(1).Width = 0;
                    worksheet_2.Column(4).Width = 4;
                    worksheet_2.Column(9).Width = 4;

                    worksheet_2.Column(2).Width = 16;
                    worksheet_2.Column(3).Width = 16;
                    worksheet_2.Column(5).Width = 16;
                    worksheet_2.Column(6).Width = 16;
                    worksheet_2.Column(7).Width = 16;
                    worksheet_2.Column(8).Width = 16;
                    worksheet_2.Column(10).Width = 16;
                    worksheet_2.Column(11).Width = 16;
                    worksheet_2.Column(12).Width = 16;
                    worksheet_2.Column(13).Width = 16;

                    worksheet_2.Row(1).Height = 48;
                    worksheet_2.Row(2).Height = 48;
                    worksheet_2.Row(3).Height = 48;
                    worksheet_2.Row(4).Height = 48;
                    worksheet_2.Row(5).Height = 48;
                    worksheet_2.Row(6).Height = 48;
                    worksheet_2.Row(7).Height = 48;
                    worksheet_2.Row(8).Height = 48;
                    worksheet_2.Row(9).Height = 48;

                    worksheet_2.Range("B4:C7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("E4:F7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("G4:H7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("J4:K7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("L4:M7").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);

                    worksheet_2.Range("C4:C7").Style.Border.SetLeftBorder(XLBorderStyleValues.MediumDashed);
                    worksheet_2.Range("F4:F7").Style.Border.SetLeftBorder(XLBorderStyleValues.MediumDashed);
                    worksheet_2.Range("G4:G7").Style.Border.SetRightBorder(XLBorderStyleValues.MediumDashed);
                    worksheet_2.Range("J4:J7").Style.Border.SetRightBorder(XLBorderStyleValues.MediumDashed);
                    worksheet_2.Range("L4:L7").Style.Border.SetRightBorder(XLBorderStyleValues.MediumDashed);

                    worksheet_2.Range("B4:C7").Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("E4:H7").Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Range("J4:M7").Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);

                    worksheet_2.Range("B2:M2").Merge();
                    worksheet_2.Range("B2:M2").Style.Border.SetOutsideBorder(XLBorderStyleValues.Medium);
                    worksheet_2.Cell("B2").Value = "教卓";
                    worksheet_2.Range("B9:M9").Merge();
                    worksheet_2.Cell("B9").Value = "大天才上尾先輩大感謝";

                    worksheet_2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet_2.Range("B2:M9").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet_2.Range("B1:M1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

                    worksheet_1.Range("A1:Z100").Style.Font.SetFontName(font);
                    worksheet_2.Range("A1:Z100").Style.Font.SetFontName(font);

                    worksheet_2.Range("A1:Z100").Style.Font.SetFontSize(28);

                    worksheet_2.Range("B1:E1").Merge();
                    worksheet_2.Cell("B1").Value = students_txt + "人 / 第" + i.ToString() + "回";
                    worksheet_2.Cell("B1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet_2.Cell("H1").Value = "年";
                    worksheet_2.Cell("K1").Value = "月";
                    worksheet_2.Cell("M1").Value = "日";
                    worksheet_2.Range("A1:M1").Style.Font.SetFontSize(20);
                    //座席表終わり

                    for (int j = 1; j <= students; j++)  //jは1~studentsまでの数, 1列目に1から最終出席番号まで数値入力
                    {
                        worksheet_1.Cell(j, 1).Value = j;
                    }
                    int l = 1;
                    for (int k = i; k >= 1; k--)
                    {
                        worksheet_1.Cell(l, 2).Value = k;
                        l = l + 1;
                    }
                    l = i + 1;
                    for (int m = students; m > i; m--)
                    {
                        worksheet_1.Cell(l, 2).Value = m;
                        l = l + 1;
                    }

                    int s = (students - 1) / 2;
                    if (i % 2 == 1)  //奇数回目のとき
                    {
                        int r = (i + 1) / 2 + s;
                        for (int t = 1; t <= s; t++)
                        {
                            worksheet_1.Row(r).Delete();
                            r--;
                        }
                    }
                    else  //偶数回目のとき
                    {
                        int r = i / 2 + s;
                        for (int t = 1; t <= s; t++)
                        {
                            worksheet_1.Row(r).Delete();
                            r--;
                        }
                    }
                    for (int Zn = 1; Zn <= (students + 1) / 2; Zn++)
                    {
                        string sc1 = worksheet_1.Cell(Zn, 1).Value.ToString();
                        string sc2 = worksheet_1.Cell(Zn, 2).Value.ToString();
                        if(sc1 == sc2)
                        {
                            worksheet_1.Cell(Zn, 2).Value = students + 1;
                        }
                    }

                    //ランダム配列
                    int[] ary0 = new int[(students + 2) / 2];
                    for (int num1 = 0; num1 <= students  / 2; num1++)
                    {
                        ary0[num1] = num1 + 1;
                    }
                    ary0 = ary0.OrderBy(i => Guid.NewGuid()).ToArray();
                    int arn = 41;
                    int[] ary = new int[20] {arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn, arn };
                    for (int num2 = 0; num2 <= students / 2; num2++)
                    {
                        int jk = ary0[num2];
                        ary[num2] = jk;
                    }
                    //ランダム配列終わり

                    string a0 = ary[0].ToString();
                    string a1 = ary[1].ToString();
                    string a2 = ary[2].ToString();
                    string a3 = ary[3].ToString();
                    string a4 = ary[4].ToString();
                    string a5 = ary[5].ToString();
                    string a6 = ary[6].ToString();
                    string a7 = ary[7].ToString();
                    string a8 = ary[8].ToString();
                    string a9 = ary[9].ToString();

                    string b0 = ary[10].ToString();
                    string b1 = ary[11].ToString();
                    string b2 = ary[12].ToString();
                    string b3 = ary[13].ToString();
                    string b4 = ary[14].ToString();
                    string b5 = ary[15].ToString();
                    string b6 = ary[16].ToString();
                    string b7 = ary[17].ToString();
                    string b8 = ary[18].ToString();
                    string b9 = ary[19].ToString();

                    worksheet_2.Cell("E4").FormulaA1 = "ペア!A" + a0;
                    worksheet_2.Cell("F4").FormulaA1 = "ペア!B" + a0;
                    worksheet_2.Cell("G4").FormulaA1 = "ペア!B" + a1;
                    worksheet_2.Cell("H4").FormulaA1 = "ペア!A" + a1;
                    worksheet_2.Cell("B4").FormulaA1 = "ペア!A" + a2;
                    worksheet_2.Cell("C4").FormulaA1 = "ペア!B" + a2;
                    worksheet_2.Cell("J4").FormulaA1 = "ペア!B" + a3;
                    worksheet_2.Cell("K4").FormulaA1 = "ペア!A" + a3;
                    worksheet_2.Cell("E5").FormulaA1 = "ペア!A" + a4;
                    worksheet_2.Cell("F5").FormulaA1 = "ペア!B" + a4;
                    worksheet_2.Cell("G5").FormulaA1 = "ペア!B" + a5;
                    worksheet_2.Cell("H5").FormulaA1 = "ペア!A" + a5;
                    worksheet_2.Cell("B5").FormulaA1 = "ペア!A" + a6;
                    worksheet_2.Cell("C5").FormulaA1 = "ペア!B" + a6;
                    worksheet_2.Cell("J5").FormulaA1 = "ペア!B" + a7;
                    worksheet_2.Cell("K5").FormulaA1 = "ペア!A" + a7;
                    worksheet_2.Cell("E6").FormulaA1 = "ペア!A" + a8;
                    worksheet_2.Cell("F6").FormulaA1 = "ペア!B" + a8;
                    worksheet_2.Cell("G6").FormulaA1 = "ペア!B" + a9;
                    worksheet_2.Cell("H6").FormulaA1 = "ペア!A" + a9;

                    worksheet_2.Cell("B6").FormulaA1 = "ペア!B" + b0;
                    worksheet_2.Cell("C6").FormulaA1 = "ペア!A" + b0;
                    worksheet_2.Cell("J6").FormulaA1 = "ペア!A" + b1;
                    worksheet_2.Cell("K6").FormulaA1 = "ペア!B" + b1;
                    worksheet_2.Cell("L6").FormulaA1 = "ペア!B" + b2;
                    worksheet_2.Cell("M6").FormulaA1 = "ペア!A" + b2;
                    worksheet_2.Cell("E7").FormulaA1 = "ペア!A" + b3;
                    worksheet_2.Cell("F7").FormulaA1 = "ペア!B" + b3;
                    worksheet_2.Cell("G7").FormulaA1 = "ペア!B" + b4;
                    worksheet_2.Cell("H7").FormulaA1 = "ペア!A" + b4;
                    worksheet_2.Cell("B7").FormulaA1 = "ペア!A" + b5;
                    worksheet_2.Cell("C7").FormulaA1 = "ペア!B" + b5;
                    worksheet_2.Cell("J7").FormulaA1 = "ペア!B" + b6;
                    worksheet_2.Cell("K7").FormulaA1 = "ペア!A" + b6;
                    worksheet_2.Cell("L7").FormulaA1 = "ペア!A" + b7;
                    worksheet_2.Cell("M7").FormulaA1 = "ペア!B" + b7;
                    worksheet_2.Cell("L5").FormulaA1 = "ペア!B" + b8;
                    worksheet_2.Cell("M5").FormulaA1 = "ペア!A" + b8;
                    worksheet_2.Cell("L4").FormulaA1 = "ペア!A" + b9;
                    worksheet_2.Cell("M4").FormulaA1 = "ペア!B" + b9;

                    for(int ko = 4; ko <= 8; ko++)
                    {
                        for(int ko2 = 2; ko2 <= 13; ko2++)
                        {
                            if (worksheet_2.Cell(ko, ko2).Value.ToString() == "0")
                            {
                                worksheet_2.Cell(ko, ko2).Value = "";
                            }
                        }
                    }
                    if (i < 10)
                    {
                        string asl = "0" + i;
                        workbook.SaveAs(@"C:\大天才上尾先輩大感謝\" + timenow + "大天才上尾先輩大感謝-地学教室座席順\\大天才上尾先輩大感謝-地学教室座席順 第" + asl + "回.xlsx");
                    }
                    else
                    {
                        string asl = "" + i;
                        workbook.SaveAs(@"C:\大天才上尾先輩大感謝\" + timenow + "大天才上尾先輩大感謝-地学教室座席順\\大天才上尾先輩大感謝-地学教室座席順 第" + asl + "回.xlsx");
                    }
                }
            }

            


            MessageBox.Show("\"C:\\大天才上尾先輩大感謝\\" + timenow + " 大天才上尾先輩大感謝-地学教室座席順\"にExcelブックを作成しました。", title, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            MessageBox.Show("アプリケーションを終了します。", title, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            Application.Exit();
        }

        private void 環境設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("そんな機能は無いよ!!!", "座席順 - 地学教室", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)  //拡張機能
        {
            MessageBox.Show("そんな機能は無いよ!!!", "座席順 - 地学教室", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void toolStripButton2_Click(object sender, EventArgs e)  //ツール
        {
            MessageBox.Show("そんな機能は無いよ!!!", "座席順 - 地学教室", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}