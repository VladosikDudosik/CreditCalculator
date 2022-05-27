using System;
using System.Windows.Forms;
class Credit
{
    public double sum { get; private set; }
    private double p;
    private double n;
    public int overpayment { get; private set; }
    public int finalPrice { get; private set; }
    public int AmountOfPayment { get; private set;}
    public Credit(int credit, double p, int n)
    {
        overpayment = 0;
        this.sum = credit;
        this.p = p/100;
        this.n = n*12;
        AmountOfPayment = (int)Math.Round((sum * this.p / 12) / (1 - Math.Pow(1 + this.p / 12, -this.n))) ;
        finalPrice = (int)((double)AmountOfPayment * this.n);
    }
    
    public string Table()
    {
        string table = String.Format("|{0,5}|  |{1,5}|  |{2,5}|  |{3,5}|", "Sum", "Amount", "Body of credit", "Procents") + '\n';
        for (int i = 0; i < n; i++)
        {
            table += String.Format("|{0,5}|  |{1,5}|  |{2,5}|               |{3,5}|", sum, AmountOfPayment, AmountOfPayment - Math.Round(sum * p / 12), Math.Round(sum * p / 12)) + "\n";
            sum -= Math.Round(AmountOfPayment - sum * p / 12);
        }
        return table;
    }
    public void CreateTable( ref DataGridView table)
    {
        for (int i = 0; i < n; i++)
        {
            overpayment += (int)Math.Round(sum * p / 12);
            string[] row =
            {
                Convert.ToString(i+1),
                Convert.ToString(AmountOfPayment),
                Convert.ToString(AmountOfPayment - Math.Round(sum * p / 12)),
                Convert.ToString(Math.Round(sum * p / 12)),
                Convert.ToString(sum - Math.Round(AmountOfPayment - sum * p / 12))
            };
            table.Rows.Add(row);
            sum -= Math.Round(AmountOfPayment - sum * p / 12)  ;
        }
    }
}

