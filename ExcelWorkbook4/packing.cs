using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Media.Media3D;

namespace _3dbinpscking
{
    public class packing
    {

        List<itembox> newcbox = new List<_3dbinpscking.itembox>();
        List<itembox> cbox = new List<_3dbinpscking.itembox>();
        Point3DCollection points = new Point3DCollection();
        List<itembox> itembox = new List<_3dbinpscking.itembox>();

        public double Setpoint(List<itembox> pbox, List<itembox> wbox)
        {
            double Vs = 0;
            double Vm = 0;
            double Bestv = 0;
            double xom = 0;
            List<itembox> newboxes = new List<_3dbinpscking.itembox>();
            List<itembox> Bestbox = Clone<itembox>(pbox);
            itembox = Clone<itembox>(pbox);
            cbox = Clone<itembox>(wbox);
            IEnumerable<itembox> itemboxsort = itembox.OrderByDescending(itembox => itembox.height * itembox.lenght * itembox.width).ThenBy(itembox => itembox.type);
            foreach (var q in itemboxsort) { newboxes.Add(q); }
            Bestbox = Clone<itembox>(newboxes);
            int docount = 0;//循环计数器
            while (true)
            {

                Bestv = 0;
                docount = docount + 1;
                newboxes = Clone<itembox>(Bestbox);
                for (int k = 0; k < cbox.Count; k++)
                {
                    double Lx = 0;
                    double Lz = 0;
                    double H = cbox[k].width;
                    double W = cbox[k].lenght;
                    double D = cbox[k].height;
                    itembox = Clone<itembox>(newboxes);
                    Vm = H * W * D;
                    points.Clear();
                    points.Add(new Point3D(0, 0, 0));
                    for (int i = 0; i < itembox.Count(); i++)
                    {
                        if (itembox[i].flag != 0) { continue; }
                        double Itemx = itembox[i].width;
                        double Itemy = itembox[i].lenght;
                        double Itemz = itembox[i].height;
                        for (int j = 0; j < points.Count; j++)
                        {
                            var Xd = points[j].X + Itemx;
                            var Zd = points[j].Z + Itemz;

                            if (itembox[i].flag == 0)
                            {
                                if (Itemx <= H - points[j].X && Itemz <= D - points[j].Z && Itemy <= W - points[j].Y && ifinbox(i, j, itembox, docount) && Xd <= Lx && Zd <= Lz)
                                {
                                    itembox[i].flag = docount;
                                    itembox[i].point = points[j];
                                    itembox[i].inbox = cbox[k].type + docount;
                                    points = updatapoints(Itemy, j, Xd, Zd);
                                    break;

                                }
                                if (Lx == 0 || Lx == H)
                                {

                                    if (Itemx <= H - Lx && Itemz <= D - Lz && Itemy <= W - points[j].Y && ifinbox(i, j, itembox, docount))
                                    {
                                        itembox[i].flag = docount;
                                        itembox[i].point = points[j];
                                        itembox[i].inbox = cbox[k].type + docount;
                                        Lz = Lz + Itemz;
                                        Lx = Itemx;
                                        points = updatapoints(Itemy, j, Xd, Zd);
                                        break;

                                    }
                                    else if (Lz < D)
                                    {
                                        Lz = D;
                                        Lx = H;
                                        i = i - 1;
                                        break;
                                    }

                                }
                                else if (points[j].X == Lx && points[j].Y == 0)
                                {
                                    if (Itemx <= H - Lx && Itemz <= D - Lz && Itemy <= W - points[j].Y && ifinbox(i, j, itembox, docount) && points[j].Z + Itemy <= Lz)
                                    {
                                        itembox[i].flag = docount;
                                        itembox[i].point = points[j];
                                        itembox[i].inbox = cbox[k].type + docount;
                                        Lx = Lx + Itemx;
                                        points = updatapoints(Itemy, j, Xd, Zd);
                                        break;
                                    }
                                    if (itembox[i].flag == 0 && i != 0)
                                    {
                                        Lx = H;
                                        i = i - 1;
                                        break;
                                    }
                                }

                            }

                        }


                    }

                    double allv = 0;
                    foreach (var inbox in itembox)
                    {
                        if (inbox.flag == docount)
                        {
                            allv = allv + inbox.width * inbox.height * inbox.lenght;

                        }
                    }
                    xom = allv / Vm;//计算体积装填率
                    if (xom >= Bestv)
                    {

                        Bestv = xom;
                        Bestbox = Clone<itembox>(itembox);
                    }

                }
                Vs = Vs + Bestv;
                if (allinbox(Bestbox) == itembox.Count || docount == 100)
                { break; }

            }
            xom = Vs / docount;
            itembox = Clone<itembox>(Bestbox);
            return xom;
        }

        private Point3DCollection updatapoints(double Itemy, int j, double Xd, double Zd)
        {
            double Px = points[j].X;
            double Py = points[j].Y;
            double Pz = points[j].Z;
            if (samepoint(Xd, Py, Pz)) { points.Add(new Point3D(Xd, Py, Pz)); }
            if (samepoint(Px, Py + Itemy, Pz)) { points.Add(new Point3D(Px, Py + Itemy, Pz)); }
            if (samepoint(Px, Py, Zd)) { points.Add(new Point3D(Px, Py, Zd)); }
            points.Remove(points[j]);
            List<Point3D> a = new List<Point3D>();
            for (int i = 0; i < points.Count; i++)
            {
                if (points[i].X == Px && points[i].Z == Pz && points[i].Y < Py)
                {
                    points.RemoveAt(i);
                }
                else if (points[i].X == Px && points[i].Z == Pz && points[i].Y < Py + Itemy)
                {
                    points.RemoveAt(i);
                }
                else if (points[i].X == Px && points[i].Y == Py && points[i].Z < Zd)
                {
                    points.RemoveAt(i);
                }
                else if (points[i].Z == Pz && points[i].Y == Py && points[i].X < Xd)
                {
                    points.RemoveAt(i);
                }

            }
            Point3DCollection newpoints = new Point3DCollection();
            IEnumerable<Point3D> sortp = points.OrderBy(points => points.X).ThenBy(points => points.Y).ThenBy(points => points.Z);
            foreach (var p in sortp)
            {
                newpoints.Add(p);
            }
            return newpoints;
        }
        /// <summary>
        /// 判断点集合中是否含有相同的点
        /// </summary>
        /// <param name="px"></param>
        /// <param name="py"></param>
        /// <param name="pz"></param>
        /// <returns></returns>
        private bool samepoint(double px, double py, double pz)
        {
            bool result = true;
            foreach (var a in points)
            {
                if (a.X == px && a.Y == py && a.Z == pz)
                {
                    result = false;
                    break;
                }
                else if (a.X > px && a.Y == py && a.Z == pz)
                {
                    result = false;
                    break;
                }
                //else if (a.X == px && a.Y > py && a.Z == pz)
                //{
                //    result = false;
                //    break;
                //}

            }
            return result;
        }

        public List<itembox> getbox()
        {
            return itembox;


        }
        bool ifinbox(int i, int j, List<itembox> ifbox, int count)
        {
            double Px = points[j].X;
            double Py = points[j].Y;
            double Pz = points[j].Z;
            bool result = true;
            foreach (var q in ifbox)
            {
                //    if (Px + ifbox[i].width > q.point.X && q.point.X + q.width > Px && Py + ifbox[i].lenght > q.point.Y && q.point.Y + q.lenght > Py && q.flag == count && Pz == q.point.Z)

                //    {
                //        result = false;
                //        break;
                //    }
                if (2 * Math.Abs((Px + ifbox[i].width / 2) - (q.point.X + q.width / 2)) < ifbox[i].width + q.width && 2 * Math.Abs((Py + ifbox[i].lenght / 2) - (q.point.Y + q.lenght / 2)) < ifbox[i].lenght + q.lenght && q.flag == count && Pz == q.point.Z)
                {
                    result = false;
                    break;
                }
            }
            return result;

        }
        private int allinbox(List<itembox> ifbox)
        {
            int count = 0;
            foreach (var a in ifbox)
            {

                if (a.flag != 0)
                {
                    count++;

                }
            }
            return count;
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
