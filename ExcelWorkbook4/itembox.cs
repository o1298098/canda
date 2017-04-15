using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Media.Media3D;

namespace _3dbinpscking
{
    [Serializable]
    public class itembox
    {
        private double _width;
        private double _lenght;
        private double _height;
        private double _flag;
        private string _type;
        private Point3D _point;
        private string _inbox;
        public double width
        {
            get { return _width; }
            set { _width = value; }
        }
        public double lenght
        {
            get { return _lenght; }
            set { _lenght = value; }
        }
        public double height
        {
            get { return _height; }
            set { _height = value; }
        }
        public double flag
        {
            get { return _flag; }
            set { _flag = value; }
        }

        public string type
        {
            get { return _type; }
            set { _type = value; }
        }
        public Point3D point
        {
            get { return _point; }
            set { _point = value; }
        }
        public string inbox
        {
            get { return _inbox; }
            set { _inbox = value; }
        }

        public object Clone()
        {
            BinaryFormatter formatter = new BinaryFormatter(null, new System.Runtime.Serialization.StreamingContext(System.Runtime.Serialization.StreamingContextStates.Clone));
            MemoryStream stream = new MemoryStream();
            formatter.Serialize(stream, this);
            stream.Position = 0;
            object clonedObj = formatter.Deserialize(stream);
            stream.Close();
            return clonedObj;
        }
    }
}
