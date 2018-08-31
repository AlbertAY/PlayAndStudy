using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExeclTool.Model
{
    /// <summary>
    /// 图片信息对象
    /// </summary>
    public class PicturesInfo
    {
        /// <summary>
        /// 最小开始行
        /// </summary>
        public int MinRow { get; set; }
        /// <summary>
        /// 最大开始行
        /// </summary>
        public int MaxRow { get; set; }
        /// <summary>
        /// 最小开始列
        /// </summary>
        public int MinCol { get; set; }
        /// <summary>
        /// 最大开始列
        /// </summary>
        public int MaxCol { get; set; }
        /// <summary>
        /// 图片数据
        /// </summary>
        public Byte[] PictureData { get; private set; }
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="pictureData"></param>
        public PicturesInfo(int minRow, int maxRow, int minCol, int maxCol, Byte[] pictureData)
        {
            this.MinRow = minRow;
            this.MaxRow = maxRow;
            this.MinCol = minCol;
            this.MaxCol = maxCol;
            this.PictureData = pictureData;
        }
    }
}
