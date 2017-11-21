using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sudoku
{
    class Coordinate_Sudoku
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int Value { get; set; }
        public int Block { get; set; }

        public Coordinate_Sudoku(int row,int col,int value,int block)
        {
            this.Row = row;
            this.Column = col;
            this.Value = value;
            this.Block = block;
        }

        public override string ToString()
        {
            return "[" + Value + ":(" + Row + "," + Column + ")-" + Block + "]";
        }
    }
}
