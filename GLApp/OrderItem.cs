using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GLApp
{
    class OrderItem
    {
        public enum State
        {
            NoInfo = -2,
            Rejected = -1,
            Success = 0,
            Waiting = 1,
            InProgress = 2
        }

        public Int32 ID;
        public Int32 ClientID;
        public double Price;
        public string FromAddress;
        public string ToAddress;
        public DateTime PickTime;
        public bool[] Additions; //Детское сиденье, Кондиционер, Пьяный пассажир, Перевозка животного
        public string ClData;
        public string DrData;
        public State Status;
    }
}
