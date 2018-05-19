using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTemplateConverterLib
{
    public sealed class Dataset : IEquatable<Dataset>
    {
        public struct Indicators : IEquatable<Indicators>
        {
            public double sum, amount;
            public Indicators(int s, int a)
            {
                sum = s;
                amount = a;
            }

            public override bool Equals(object obj)
            {
                return obj is Indicators && Equals((Indicators)obj);
            }

            public bool Equals(Indicators other)
            {
                return sum == other.sum &&
                       amount == other.amount;
            }

            public override int GetHashCode()
            {
                var hashCode = 609155889;
                hashCode = hashCode * -1521134295 + base.GetHashCode();
                hashCode = hashCode * -1521134295 + sum.GetHashCode();
                hashCode = hashCode * -1521134295 + amount.GetHashCode();
                return hashCode;
            }

            public static bool operator ==(Indicators indicators1, Indicators indicators2)
            {
                return indicators1.Equals(indicators2);
            }

            public static bool operator !=(Indicators indicators1, Indicators indicators2)
            {
                return !(indicators1 == indicators2);
            }
        }
        public struct Balance : IEquatable<Balance>
        {
            public Indicators debit, credit;
            public Balance(Indicators d, Indicators c)
            {
                debit = d;
                credit = c;
            }

            public override bool Equals(object obj)
            {
                return obj is Balance && Equals((Balance)obj);
            }

            public bool Equals(Balance other)
            {
                return debit.Equals(other.debit) &&
                       credit.Equals(other.credit);
            }

            public override int GetHashCode()
            {
                var hashCode = 264963065;
                hashCode = hashCode * -1521134295 + base.GetHashCode();
                hashCode = hashCode * -1521134295 + debit.GetHashCode();
                hashCode = hashCode * -1521134295 + credit.GetHashCode();
                return hashCode;
            }

            public static bool operator ==(Balance balance1, Balance balance2)
            {
                return balance1.Equals(balance2);
            }

            public static bool operator !=(Balance balance1, Balance balance2)
            {
                return !(balance1 == balance2);
            }
        }

        private string invoice = "";
        private string name = "";
        private string inventoryNumber = "";
        private string kfo = "";
        public Balance startPeriodBalance = new Balance();
        public Balance turnover = new Balance();
        public Balance endPeriodBalance = new Balance();

        public string Invoice { get => invoice; set => invoice = value; }
        public string Name { get => name; set => name = value; }
        public string InventoryNumber { get => inventoryNumber; set => inventoryNumber = value; }
        public string KFO { get => kfo; set => kfo = value; }

        public override bool Equals(object obj)
        {
            return Equals(obj as Dataset);
        }

        public bool Equals(Dataset other)
        {
            return other != null &&
                   invoice == other.invoice &&
                   name == other.name &&
                   inventoryNumber == other.inventoryNumber &&
                   kfo == other.kfo &&
                   startPeriodBalance.Equals(other.startPeriodBalance) &&
                   turnover.Equals(other.turnover) &&
                   endPeriodBalance.Equals(other.endPeriodBalance);
        }

        public override int GetHashCode()
        {
            var hashCode = 1134822336;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(invoice);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(name);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(inventoryNumber);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(kfo);
            hashCode = hashCode * -1521134295 + EqualityComparer<Balance>.Default.GetHashCode(startPeriodBalance);
            hashCode = hashCode * -1521134295 + EqualityComparer<Balance>.Default.GetHashCode(turnover);
            hashCode = hashCode * -1521134295 + EqualityComparer<Balance>.Default.GetHashCode(endPeriodBalance);
            return hashCode;
        }

        public static bool operator ==(Dataset dataset1, Dataset dataset2)
        {
            return EqualityComparer<Dataset>.Default.Equals(dataset1, dataset2);
        }

        public static bool operator !=(Dataset dataset1, Dataset dataset2)
        {
            return !(dataset1 == dataset2);
        }
    }
}
