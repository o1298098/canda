using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Media.Media3D;

namespace _3dbinpscking
{
    public class sa3dpacking
    {
        packing pack = new packing();
        List<itembox> itembox = new List<itembox>();
        List<itembox>  Bitembox = new List<itembox>();

        /// <summary>
        /// 模拟退火算法
        /// </summary>
        /// <param name="pbox">内箱</param>
        /// <param name="cbox">外箱</param>
        /// <returns></returns>
        public double packing(List<itembox> pbox,List<itembox>cbox)
        {
            itembox = Clone<itembox>(pbox);
            IEnumerable<itembox> itemboxsort = itembox.OrderByDescending(itembox => itembox.height * itembox.lenght * itembox.width);
            List<itembox> newboxes = new List<itembox>();
            foreach (var q in itemboxsort) { newboxes.Add(q); }
            itembox = newboxes;
            double Fbest = pack.Setpoint(itembox,cbox);//最优填充率
            double fz = Fbest;
            Bitembox = pack.getbox();//最优装箱方式
            //for (int i = 0; i < 1; i++)
            //{
                double t = 1;//初始温度
                double Lt = 0;//当前邻域长度
                double Et = 0.01;//结束温度
                int dL = itembox.GroupBy(x => new { x.type }).Count();//箱子的种类
                double dt = 0.92;//温度衰减系数
                while (t >= Et)
                {
                    for (int j = 0; j < Lt; j++)
                    {
                        List<itembox> Bx = niegbourhood(itembox);//获取itembox的邻域
                        double f = pack.Setpoint(Bx,cbox);
                        double df = f - fz;
                        if (df > 0)
                        {
                            fz = f;
                            itembox = Clone<itembox>(Bx);
                            if (fz > Fbest)
                            {
                                
                                Fbest = fz;
                                Bitembox = pack.getbox();
                            }
                        }
                        else
                        {
                            long tick = DateTime.Now.Ticks;
                            Random random = new Random((int)(tick & 0xffffffffL) | (int)(tick >> 32));
                            double numx = random.NextDouble();
                            if (numx < Math.Pow((1 - 6*dt / t),1 /6))
                            {
                              
                                fz = f;
                                itembox = Clone<itembox>(Bx);
                            }
                        }
                    }
                    Lt += dL;
                    //t *= dt;
                    t = t * Math.Pow(dt, Lt);
                    //t = t * Math.Exp(-dt * Math.Pow(Lt - 1, 1 / 2));
                    //t = t * Math.Exp(-dt * (Lt - 1));
                }
            //}
           
            return Fbest;
            
        }
        /// <summary>
        /// 计算商品箱的邻域
        /// </summary>
        /// <param name="itembox"></param>
        /// <returns></returns>
        List<itembox> niegbourhood(List<itembox> itembox)
        {
          
           
            List<itembox> nitembox = itembox;
            long tick = DateTime.Now.Ticks;
            Random random = new Random((int)(tick & 0xffffffffL) | (int)(tick >> 32));
            var boxtype = itembox.GroupBy(x => new { x.type });//获取箱子的种类
            int boxnum = boxtype.Count();
            int box = random.Next(1, boxnum + 1);
            string newbox = string.Empty;
            int c = 1;
            int turntype = random.Next(1, 4);
            foreach (var q in boxtype)
            {

                if (c == box) { newbox = q.Key.type; break; }

                c++;
            }
            for (int i = 0; i < itembox.Count; i++)
            {
                itembox[i].flag = 0;
                itembox[i].point = new Point3D(0, 0, 0);
                if (itembox[i].type == newbox)
                {
                    double midx = itembox[i].width;
                    double midy = itembox[i].lenght;
                    double midz = itembox[i].height;
                    switch (turntype)
                    {
                        case 1:

                            nitembox[i].width = midy;
                            nitembox[i].lenght = midx;
                            break;
                        case 2:

                            nitembox[i].width = midz;
                            nitembox[i].height = midx;
                            break;
                        case 3:

                            nitembox[i].height = midy;
                            nitembox[i].lenght = midz;
                            break;
                    }//根据随机数对调随意两边的长度

                }


            }
            return nitembox;
        }


        public List<itembox> getbox()
        {
            return Bitembox;
        }
      
        public static List<T> Clone<T>(object List)
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, List);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as List<T>;
            }
        }
    }
}
