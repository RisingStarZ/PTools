///Copyright (C) 2017  RisingStarZ
///This program is free software: you can redistribute it and/or modify it under the terms 
///of the GNU Affero General Public License as published by the Free Software Foundation, 
///either version 3 of the License, or(at your option) any later version.
///
///This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
///without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
///See the GNU Affero General Public License for more details.
///
///You should have received a copy of the GNU Affero General Public License along with this program.
///If not, see<http://www.gnu.org/licenses/>.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace PTools
{
    public class Roll : IEnumerator<SortedRoll>, IEnumerable<SortedRoll>
    {
        private int index = -1;
        private List<SortedRoll> _roll;

        public SortedRoll this[int index]
        {
            get
            {
                return _roll[index];
            }
            set
            {
                _roll[index] = value;
            }
        }

        public Roll(string file)
        {
            _roll = new List<SortedRoll>();
            using (var sr = new StreamReader(file))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    _roll.Add(new SortedRoll(line.Split('|')));
                }
            }
        }

        public SortedRoll Current
        {
            get
            {
                return _roll[index];
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return _roll[index];
            }
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        public IEnumerator<SortedRoll> GetEnumerator()
        {
            return this;
        }

        public bool MoveNext()
        {
            if (index == _roll.Count - 1)
            {
                Reset();
                return false;
            }
            index++;
            return true;
        }

        public void Reset()
        {
            index = -1;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _roll.GetEnumerator();
        }

        public int Length
        {
            get { return _roll.Count; }
        }
    }

    public class Addressee
    {
        private int _index;
        private string _address;
        private string _name;
        private string _barcode;

        public int Index
        {
            get
            {
                return _index;
            }
            set
            {
                _index = value;
            }
        }

        public string Address
        {
            get
            {
                return _address;
            }
            set
            {
                _address = value;
            }
        }

        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                _name = value;
            }
        }

        public string Barcode
        {
            get
            {
                return _barcode;
            }
            set
            {
                _barcode = value;
            }
        }

        public Addressee(string Name, string Address, int Index, string Barcode = "")
        {
            this.Name = Name;
            this.Address = Address;
            this.Index = Index;
            this.Barcode = Barcode;
        }
    }

    public class SortedRoll
    {
        public string Destination { get; set; }
        public string Num { get; set; }
        public string Id { get; set; }
        public string Barcode { get; set; }
        
        public SortedRoll() { }
        public SortedRoll(string[] input)
        {
            Destination = input[1];
            Num = input[0];
            Id = input[2];
            Barcode = input[3];
        }
    }
}
