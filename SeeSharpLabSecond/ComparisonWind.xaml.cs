using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SeeSharpLabSecond
{
    /// <summary>
    /// Логика взаимодействия для ComparisonWind.xaml
    /// </summary>
    public partial class ComparisonWind : Window
    {
        private List<Threat> BeforeBase { set; get; }
        private List<Threat> AfterBase { set; get; }

        public ComparisonWind()
        {
            InitializeComponent();

        }
        public void UpdateWindows(List<Threat> a, List<Threat> b)
        {
            BeforeBase = a;
            AfterBase = b;
            beforeWindow.ItemsSource = BeforeBase;
            AfterWindow.ItemsSource = AfterBase;
        }
        public int CompareBases()
        {
            List<Threat> tempAfterBase = new List<Threat>();
            List<Threat> tempBeforeBase = new List<Threat>();
            int counter = 0;
            int length = 0;
            
            if (BeforeBase.Count > AfterBase.Count) length = AfterBase.Count;
            else length = BeforeBase.Count;

            for (int i = 0; i < length; i++)
            {
                if (AfterBase[i].Name != BeforeBase[i].Name
                    || AfterBase[i].Description != BeforeBase[i].Description
                    || AfterBase[i].Availability != BeforeBase[i].Availability
                    || AfterBase[i].Confidence != BeforeBase[i].Confidence
                    || AfterBase[i].ChangedDate != BeforeBase[i].ChangedDate
                    || AfterBase[i].AddedDate != BeforeBase[i].AddedDate
                    || AfterBase[i].Source != BeforeBase[i].Source
                    || AfterBase[i].Integrity != BeforeBase[i].Integrity
                    || AfterBase[i].Id != BeforeBase[i].Id
                    || AfterBase[i].Target != BeforeBase[i].Target)
                { tempAfterBase.Add(AfterBase[i]); tempBeforeBase.Add(BeforeBase[i]); counter++; }
            }
            if (BeforeBase.Count > AfterBase.Count) tempBeforeBase.AddRange(BeforeBase.GetRange(AfterBase.Count, BeforeBase.Count - AfterBase.Count));
            else if (AfterBase.Count > BeforeBase.Count) tempAfterBase.AddRange(AfterBase.GetRange(BeforeBase.Count, AfterBase.Count - BeforeBase.Count));
            UpdateWindows(tempBeforeBase,tempAfterBase);

            counter += Math.Abs(BeforeBase.Count - AfterBase.Count);
            return counter;
        }


    }
}
