using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Sudoku
{
    class Program
    {
        static void Main(string[] args)
        {
            Application xlsApp = null;
            Workbook wb = null;
            Worksheet ws = null;
            Worksheet ws_Solution = null;

            List<Coordinate_Sudoku> Origin = new List<Coordinate_Sudoku>();
            List<Coordinate_Sudoku> Solution = new List<Coordinate_Sudoku>();
            string path = System.Environment.CurrentDirectory + @"\Sudoku.xlsx";

            if (File.Exists(path))
            {
                try
                {
                    xlsApp = new Application();
                    xlsApp.DisplayAlerts = false;
                    xlsApp.AskToUpdateLinks = false;
                    wb = xlsApp.Workbooks.Open(path);
                    ws = wb.Worksheets[1];
                    ws.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                    ws_Solution = wb.Worksheets[wb.Worksheets.Count];

                    Origin = getOrigin(ws);

                    Console.WriteLine("原題目：([Value:(Row,Column)-Block])");
                    for (int i = 1; i < 10;i++ )
                    {
                        for(int j = 1;j < 10;j++)
                        {
                            var temp = Origin.Where(c => c.Row == i && c.Column == j).FirstOrDefault();
                            if (temp != null)
                                Console.Write(temp.Value);
                            else
                                Console.Write("O");
                        }
                        Console.WriteLine();
                    }

                        Solution = SolveSolution(Origin, Solution, 1, 1, (81 - Origin.Count));

                    if(Solution.Count > 0)
                    {
                        foreach(Coordinate_Sudoku cs in Solution)
                        {
                            ws_Solution.Cells[cs.Row, cs.Column].Value2 = cs.Value;
                        }
                        Console.WriteLine("\n\nComplete!!\nPress Any Keys to continue...");
                    }
                    else
                    {
                        Console.WriteLine("No Solution.");
                    }
                    Console.ReadLine();

                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                finally
                {
                    wb.Close(true);
                    xlsApp.DisplayAlerts = true;
                    xlsApp.AskToUpdateLinks = true;
                    if (ws != null)
                        Marshal.FinalReleaseComObject(ws);
                    if (ws_Solution != null)
                        Marshal.FinalReleaseComObject(ws_Solution);
                    if (wb != null)
                        Marshal.FinalReleaseComObject(wb);
                    if (xlsApp != null)
                        Marshal.FinalReleaseComObject(xlsApp);
                }
            }
            else
                Console.WriteLine("找不到指定檔案");
        }

        static List<Coordinate_Sudoku> getOrigin(Worksheet ws)
        {
            List<Coordinate_Sudoku> Origin = new List<Coordinate_Sudoku>();
            for (int i = 1; i < 10; i++)
            {
                Origin.AddRange(getBlock(ws, i));
            }
            return Origin;
        }

        static List<Coordinate_Sudoku> getBlock(Worksheet ws, int block)
        {
            List<Coordinate_Sudoku> b = new List<Coordinate_Sudoku>();
            int row = 0, col = 0;
            returnRowColumn(block, ref row, ref col);
            
            for (int i = row; i < row + 3; i++)
            {
                for (int j = col; j < col + 3; j++)
                {
                    if (ws.Cells[i, j].Value2 != null)
                    {
                        b.Add(new Coordinate_Sudoku(i, j, Convert.ToInt32(ws.Cells[i, j].Value2), block));
                    }
                }
            }
            return b;
        }

        static bool isSafe(List<Coordinate_Sudoku> Origin,List<Coordinate_Sudoku> Solution, Coordinate_Sudoku cs)
        {
            List<Coordinate_Sudoku> temp1 = Origin.Where(c => c.Value == cs.Value).ToList();
            List<Coordinate_Sudoku> temp2 = Solution.Where(c => c.Value == cs.Value).ToList();
            foreach (Coordinate_Sudoku c in temp1)
            {
                if (c.Row == cs.Row || c.Column == cs.Column || c.Block == cs.Block)
                    return false;
            }
            foreach (Coordinate_Sudoku c in temp2)
            {
                if (c.Row == cs.Row || c.Column == cs.Column || c.Block == cs.Block)
                    return false;
            }
            return true;
        }
        static List<Coordinate_Sudoku> SolveSolution(List<Coordinate_Sudoku> Origin, List<Coordinate_Sudoku> Solution,int row,int col, int SolutionCount)
        {
            if(!Origin.Exists(c => c.Row == row && c.Column == col))
            {
                for(int value = 1;value < 10;value++)
                {
                    Coordinate_Sudoku temp = new Coordinate_Sudoku(row, col, value, returnBlock(row, col));
                    if(isSafe(Origin,Solution,temp))
                    {
                        Solution.Add(temp);
                        if(Solution.Count < SolutionCount)
                        {
                            if (col < 9)
                                SolveSolution(Origin, Solution, row, col + 1, SolutionCount);
                            else
                                SolveSolution(Origin, Solution, row + 1, 1, SolutionCount);
                        }
                        if (Solution.Count < SolutionCount && Solution.Count > 0)
                            Solution.RemoveAt(Solution.Count - 1);
                    }
                }
            }
            else
            {
                if (Solution.Count < SolutionCount)
                {
                    if (col < 9)
                        SolveSolution(Origin, Solution, row, col + 1, SolutionCount);
                    else
                        SolveSolution(Origin, Solution, row + 1, 1, SolutionCount);
                }
            }

            
            return Solution;
        }
        static int returnBlock(int row, int col)
        {
            if (row < 4)
            {
                if (col < 4)
                    return 1;
                else if (col > 6)
                    return 3;
                else
                    return 2;
            }
            else if (row > 6)
            {
                if (col < 4)
                    return 7;
                else if (col > 6)
                    return 9;
                else
                    return 8;
            }
            else
            {
                if (col < 4)
                    return 4;
                else if (col > 6)
                    return 6;
                else
                    return 5;
            }
        }

        static void returnRowColumn(int block,ref int row,ref int col)
        {
            switch(block)
            {
                case 1:
                    row = 1;
                    col = 1;
                    break;
                case 2:
                    row = 1;
                    col = 4;
                    break;
                case 3:
                    row = 1;
                    col = 7;
                    break;
                case 4:
                    row = 4;
                    col = 1;
                    break;
                case 5:
                    row = 4;
                    col = 4;
                    break;
                case 6:
                    row = 4;
                    col = 7;
                    break;
                case 7:
                    row = 7;
                    col = 1;
                    break;
                case 8:
                    row = 7;
                    col = 4;
                    break;
                case 9:
                    row = 7;
                    col = 7;
                    break;
            }
        }
    }
}
