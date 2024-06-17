﻿using System.Windows;
using SKToolsAddins.Utils;
using SKToolsAddins.ViewModel;
using Window = System.Windows.Window;

namespace SKToolsAddins.Forms
{
    public partial class SelectElementsVer1NumberingRuleWpfWindow : Window
    {
        public SelectElementsVer1NumberingRuleWpfWindow(SelectElementsVer1ViewModel viewModel)
        {
            InitializeComponent();
            this.DataContext = viewModel;
        }
        
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
        private void Numbering_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }


    }
}
